import csv
import io
import re
import zlib
from datetime import date, datetime
from decimal import Decimal, InvalidOperation

import openpyxl
import pandas as pd
import streamlit as st

from utils.ui_components import _callout

APP_TITLE = "Paycom Deduction Sanity Check"

# =========================================================
# Paycom Deduction Sanity Check
# - Input: one Paycom Employee Scheduled Deductions export (.csv / .xlsx)
# - Removes rows:
#     * blank EE Code
#     * Stop Date before the 1st of the current month (00/00/0000 and blank stay)
#     * no non-zero Amount and no non-zero Percent (0 / blank on both sides)
#     * every row whose Deduction Desc the user ticks (garnishments,
#       reimbursements and earned-wage access are pre-ticked)
# - Changes ONE thing: Percent x 100 when the file writes percents as
#   fractions (0.04 = 4%). The onboarding API passes Percent through as-is
#   (EmployeeDeductionMapper: Double.parseDouble(percent)), so 0.04 would be
#   set up as 0.04%. Files that already hold whole percents are left alone.
# - Everything else the API would reject is FLAGGED, not changed.
#
# API rules (onboarding-service, PayComEmployeeDeductionValidator):
#   EE Code / Deduction Code / Deduction Desc mandatory; exactly one of
#   Amount / Percent; Amount >= 0; 0 <= Percent <= 100; Tax Treatment must
#   start "A -", "B -", ... "H -"; Start / Stop Date MM/dd/yyyy with
#   00/00/0000 = none, a blank Start Date counting as TODAY, and Stop Date not
#   before it; Limit "No" or "Yes (amount)"; Limit Accum numeric. The mapping
#   is looked up by the exact Deduction Desc.
# =========================================================

COL_ID = "EE Code"
COL_CODE = "Deduction Code"
COL_DESC = "Deduction Desc"
COL_AMT = "Amount"
COL_PCT = "Percent"
COL_TAX = "Tax Treatment"
COL_START = "Start Date"
COL_STOP = "Stop Date"
COL_LIMIT = "Limit"
COL_ACCUM = "Limit Accum"
COL_MATCH = "Company Match"
COL_NAME = "EE Name"
REQUIRED_COLUMNS = [COL_ID, COL_CODE, COL_DESC, COL_AMT, COL_PCT, COL_TAX, COL_STOP]
OPTIONAL_COLUMNS = [COL_START, COL_LIMIT, COL_ACCUM, COL_MATCH, COL_NAME]

CAT_DD = "Direct deposit"
CAT_GARN = "Garnishment"
CAT_EWA = "Reimbursement / earned-wage access"
CAT_USER = "Ticked by you"
CAT_BLANK_ID = "Blank EE Code"
CAT_STOPPED = "Stop Date before the current month"
CAT_ZERO = "Amount and Percent both 0 or blank"

TAX_PREFIXES = ("A -", "B -", "C -", "D -", "E -", "F -", "G -", "H -")
NO_DATE = "00/00/0000"

_NUMBER = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
_MDY = re.compile(r"^(\d{1,2})/(\d{1,2})/(\d{4})$")
_LIMIT = re.compile(r"^Yes \([0-9,.]+\)$")


def _s(v):
    return "" if v is None else str(v)


def _num(v):
    """Decimal value of a plain number (commas dropped), else None."""
    t = _s(v).strip().replace(",", "")
    if not t or not _NUMBER.match(t):
        return None
    try:
        return Decimal(t)
    except InvalidOperation:
        return None


def parse_mdy(v):
    """A real date from M/D/YYYY, the only format the API reads. Else None."""
    m = _MDY.match(_s(v).strip())
    if not m:
        return None
    try:
        return date(int(m.group(3)), int(m.group(1)), int(m.group(2)))
    except ValueError:
        return None


def _no_date(v):
    return _s(v).strip() in ("", NO_DATE)


# ---------------------------------------------------------------- reading

def _cell_text(v):
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%m/%d/%Y")
    if isinstance(v, date):
        return v.strftime("%m/%d/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _find_header(rows, limit=20):
    for i, row in enumerate(rows[:limit]):
        cells = {str(c).strip().lower() for c in row if str(c).strip()}
        if COL_ID.lower() in cells and COL_DESC.lower() in cells:
            return i
    return None


def _frame(rows, header_idx):
    head = [str(h) for h in rows[header_idx]]
    body = []
    for r in rows[header_idx + 1:]:
        r = list(r) + [""] * (len(head) - len(r))
        if any(str(c).strip() for c in r):
            body.append([str(c) for c in r[:len(head)]])
    return pd.DataFrame(body, columns=head, dtype=str)


def read_deduction_file(file):
    """The Paycom export as all-text rows. Returns ({"df", "header_row"}, error)."""
    name = (getattr(file, "name", "") or "").lower()
    data = file.getvalue()
    sheets = []
    try:
        if data[:2] == b"PK":
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            for ws in wb.worksheets:
                sheets.append([[_cell_text(v) for v in row]
                               for row in ws.iter_rows(values_only=True)])
        elif data[:4] == bytes.fromhex("d0cf11e0") or name.endswith(".xls"):
            for df in pd.read_excel(io.BytesIO(data), sheet_name=None, header=None,
                                    dtype=str, keep_default_na=False).values():
                sheets.append(df.values.tolist())
        else:
            try:
                text = data.decode("utf-8-sig")
            except UnicodeDecodeError:
                text = data.decode("cp1252", errors="replace")
            sheets.append(list(csv.reader(io.StringIO(text))))
    except Exception as e:
        return None, f"Could not read the file: {e}"

    for rows in sheets:
        h = _find_header(rows)
        if h is not None:
            return {"df": _frame(rows, h), "header_row": h + 1}, None
    return None, ("Could not find a header row with `EE Code` and `Deduction Desc` "
                  "in the first 20 rows.")


def column_map(df):
    """{expected name: actual header} and the required names that are missing."""
    found = {str(c).strip().lower(): c for c in df.columns}
    cols = {r: found[r.lower()] for r in REQUIRED_COLUMNS + OPTIONAL_COLUMNS
            if r.lower() in found}
    return cols, [r for r in REQUIRED_COLUMNS if r not in cols]


# ---------------------------------------------------------------- percent

def detect_fraction(df, cols):
    """True when the file writes percents as fractions (0.04 = 4%).

    Decided per file: any non-zero Percent strictly between -1 and 1 means
    fractions. Chief Delivery and Spelman export whole percents (5 = 5%), and
    multiplying those would make 5% into 500%.
    Returns (is_fraction, fractions_seen, whole_seen).
    """
    vals = [_num(v) for v in df[cols[COL_PCT]]]
    vals = [v for v in vals if v is not None and v != 0]
    frac = sum(1 for v in vals if abs(v) < 1)
    return frac > 0, frac, len(vals) - frac


def scale_percent(v):
    """'0.04' -> '4', '0.00432' -> '0.432'. Blank and non-numbers come back as they are."""
    d = _num(v)
    if d is None:
        return v
    out = format((d * 100).normalize(), "f")
    return "0" if out in ("-0", "") else out


# ---------------------------------------------------------------- removal

def default_category(description):
    """Which pre-ticked group a Deduction Desc falls in, or None.

    `LOA ... Reimb` is NOT a reimbursement to the employee: it is the employee
    paying back benefit premiums during a leave of absence, a real deduction.
    """
    words = re.sub(r"[^A-Z]+", " ", _s(description).upper()).split()
    joined, compact = " ".join(words), "".join(words)
    if joined in ("CHECKING", "SAVINGS"):
        return CAT_DD
    if ("SUPPORT" in words or "LEVY" in words or "LIEN" in words or "SPT" in words
            or "WAGEASSIGN" in compact or "WAGEAGREMNT" in compact
            or any(w.startswith(("GARN", "BANKRUPT")) for w in words)):
        return CAT_GARN
    if words and words[0] == "LOA":
        return None
    # "REIM" catches Paycom's truncated "Employer Health Insurance Reim".
    if (any(w.startswith("REIM") for w in words)
            or any(k in compact for k in ("PAYACTIV", "ZAYZOON", "TAPCHECK", "EARNEDWAGE"))):
        return CAT_EWA
    return None


def description_summary(df, cols):
    live = df[df[cols[COL_ID]].str.strip() != ""]
    out = []
    for desc, g in live.groupby(cols[COL_DESC], sort=False):
        codes = sorted({c.strip() for c in g[cols[COL_CODE]] if c.strip()})
        out.append({"description": desc, "codes": codes, "rows": len(g),
                    "employees": g[cols[COL_ID]].str.strip().nunique(),
                    "default": default_category(desc)})
    return out


def month_cutoff(today=None):
    """The 1st of the current month: a Stop Date before it has already passed."""
    today = today or date.today()
    return date(today.year, today.month, 1)


def _is_zero_or_blank(v):
    t = _s(v).strip()
    if not t:
        return True
    d = _num(t)
    return d is not None and d == 0


def removal_reasons(df, cols, remove_descriptions, cutoff):
    """{row index: [reasons]} for every row that goes."""
    reasons = {}
    for idx, r in df.iterrows():
        why = []
        if not r[cols[COL_ID]].strip():
            why.append(CAT_BLANK_ID)
        else:
            stop = parse_mdy(r[cols[COL_STOP]])
            if stop is not None and stop < cutoff:
                why.append(CAT_STOPPED)
            if _is_zero_or_blank(r[cols[COL_AMT]]) and _is_zero_or_blank(r[cols[COL_PCT]]):
                why.append(CAT_ZERO)
            if r[cols[COL_DESC]] in remove_descriptions:
                why.append(default_category(r[cols[COL_DESC]]) or CAT_USER)
        if why:
            reasons[idx] = why
    return reasons


# ---------------------------------------------------------------- flags

def find_flags(kept, cols, src, today):
    """Every API-breaking issue on the rows that stay (after the % change)."""
    issues = []

    def add(idx, issue):
        issues.append((idx, issue))

    def col(r, name):
        return r[cols[name]] if name in cols else ""

    for idx, r in kept.iterrows():
        eid, code, desc = r[cols[COL_ID]], r[cols[COL_CODE]], r[cols[COL_DESC]]
        amt, pct = r[cols[COL_AMT]].strip(), r[cols[COL_PCT]].strip()
        if not code.strip():
            add(idx, "Deduction Code is blank")
        if not desc.strip():
            add(idx, "Deduction Desc is blank")
        spaced = [n for n, v in (("EE Code", eid), ("Deduction Code", code),
                                 ("Deduction Desc", desc),
                                 ("Tax Treatment", col(r, COL_TAX)),
                                 ("Company Match", col(r, COL_MATCH)))
                  if v != v.strip()]
        if spaced:
            add(idx, "Extra spaces around " + " / ".join(spaced))
        if amt and pct:
            add(idx, "Amount and Percent are both filled (the API takes only one)")
        for label, v in (("Amount", amt), ("Percent", pct)):
            if v and _num(v) is None:
                add(idx, f"{label} is not a plain number")
        a, p = _num(amt), _num(pct)
        if a is not None and a < 0:
            add(idx, "Amount is negative")
        if p is not None and (p < 0 or p > 100):
            add(idx, "Percent is outside 0-100")
        tax = col(r, COL_TAX)
        if not tax.strip():
            add(idx, "Tax Treatment is blank")
        elif not tax.startswith(TAX_PREFIXES):
            add(idx, "Tax Treatment does not start with a code such as 'A -' or 'B -'")
        start_raw, stop_raw = col(r, COL_START), r[cols[COL_STOP]]
        start = parse_mdy(start_raw)
        for label, raw, parsed in (("Start Date", start_raw, start),
                                   ("Stop Date", stop_raw, parse_mdy(stop_raw))):
            if not _no_date(raw) and parsed is None:
                add(idx, f"{label} is not MM/DD/YYYY")
        stop = parse_mdy(stop_raw)
        if stop is not None and start_raw is not None:
            # A blank / 00/00/0000 Start Date is TODAY to the API.
            effective = start if start is not None else (today if _no_date(start_raw) else None)
            if effective is not None and stop < effective:
                add(idx, "Stop Date is before the Start Date (a blank Start Date "
                         "counts as today)")
        limit = col(r, COL_LIMIT).strip()
        if limit and limit.lower() != "no" and not _LIMIT.match(limit):
            add(idx, "Limit is not 'No' or 'Yes (amount)'")
        accum = col(r, COL_ACCUM).strip()
        if accum and _num(accum) is None:
            add(idx, "Limit Accum is not a number")

    key = kept[cols[COL_ID]].str.strip() + "\x1f" + kept[cols[COL_DESC]].str.strip()
    api = [cols[c] for c in (COL_ID, COL_CODE, COL_DESC, COL_AMT, COL_PCT)]
    for _, g in kept.groupby(key, sort=False):
        if len(g) < 2:
            continue
        identical = len(g[api].apply(lambda c: c.str.strip()).drop_duplicates()) == 1
        label = ("Duplicate - identical rows for the same employee and description"
                 if identical else
                 "Duplicate - same employee and description, different code or amount")
        for idx in g.index:
            add(idx, label)

    if not issues:
        return pd.DataFrame(columns=["Source Row", "Issue"] + list(kept.columns))
    order = {idx: n for n, idx in enumerate(kept.index)}
    issues.sort(key=lambda t: order[t[0]])
    return pd.DataFrame([{"Source Row": int(src[idx]), "Issue": issue,
                          **kept.loc[idx].to_dict()} for idx, issue in issues])


# ---------------------------------------------------------------- run

def run_sanity(file, remove_descriptions=None, multiply_percent=None, today=None):
    """Read, remove, convert %, flag.

    `remove_descriptions=None` means the default ticks; `multiply_percent=None`
    means the per-file detection; `today` is for tests.
    Returns (result, error); error is a str, or {"missing", "found"}.
    """
    today = today or date.today()
    info, err = read_deduction_file(file)
    if err:
        return None, err
    df, header_row = info["df"], info["header_row"]
    cols, missing = column_map(df)
    if missing:
        return None, {"missing": missing, "found": list(df.columns)}

    summary = description_summary(df, cols)
    if remove_descriptions is None:
        remove_descriptions = {d["description"] for d in summary if d["default"]}
    is_fraction, n_frac, n_whole = detect_fraction(df, cols)
    if multiply_percent is None:
        multiply_percent = is_fraction

    cutoff = month_cutoff(today)
    src = pd.Series(range(header_row + 1, header_row + 1 + len(df)), index=df.index)
    reasons = removal_reasons(df, cols, set(remove_descriptions), cutoff)

    removed = df.loc[list(reasons)].copy()
    removed.insert(0, "Reason", ["; ".join(reasons[i]) for i in removed.index])
    removed.insert(0, "Source Row", src[removed.index])
    kept = df.drop(index=list(reasons)).copy()

    pct_changes = []
    if multiply_percent:
        pc = cols[COL_PCT]
        for idx, v in kept[pc].items():
            new = scale_percent(v)
            if new != v:
                pct_changes.append((idx, v, new))
                kept.at[idx, pc] = new

    flags = find_flags(kept, cols, src, today)
    return {"df": df, "cols": cols, "header_row": header_row, "summary": summary,
            "kept": kept.reset_index(drop=True), "kept_src": src[kept.index].tolist(),
            "removed": removed.reset_index(drop=True), "flags": flags,
            "remove": set(remove_descriptions), "cutoff": cutoff,
            "is_fraction": is_fraction, "n_frac": n_frac, "n_whole": n_whole,
            "multiply_percent": multiply_percent,
            "pct_changes": [(int(src[i]), kept.at[i, cols[COL_ID]],
                             kept.at[i, cols[COL_DESC]], old, new)
                            for i, old, new in pct_changes]}, None


def corrected_csv_bytes(kept):
    """Every original column. Plain UTF-8, NO BOM."""
    return kept.to_csv(index=False).encode("utf-8")


def change_log(result):
    cols = result["cols"]
    name_col = cols.get(COL_NAME)
    rows = []
    if result["header_row"] > 1:
        rows.append({"Source Row": "(file)", "EE Code": "(All rows)", "EE Name": "",
                     "Deduction Code": "", "Deduction Desc": "",
                     "Action": "Rows above the header removed", "Field": "",
                     "Old Value": "", "New Value": "",
                     "Reason": f"{result['header_row'] - 1} preamble row(s); the API "
                               "expects the header on row 1"})
    for _, r in result["removed"].iterrows():
        rows.append({"Source Row": r["Source Row"], "EE Code": r[cols[COL_ID]],
                     "EE Name": r[name_col] if name_col else "",
                     "Deduction Code": r[cols[COL_CODE]], "Deduction Desc": r[cols[COL_DESC]],
                     "Action": "Row removed", "Field": "", "Old Value": "", "New Value": "",
                     "Reason": r["Reason"]})
    for src, eid, desc, old, new in result["pct_changes"]:
        rows.append({"Source Row": src, "EE Code": eid, "EE Name": "",
                     "Deduction Code": "", "Deduction Desc": desc,
                     "Action": "Value changed", "Field": "Percent",
                     "Old Value": old, "New Value": new,
                     "Reason": "Paycom writes percents as fractions; the API reads 4 as 4%"})
    return pd.DataFrame(rows, columns=["Source Row", "EE Code", "EE Name", "Deduction Code",
                                       "Deduction Desc", "Action", "Field", "Old Value",
                                       "New Value", "Reason"])


def report_xlsx_bytes(result):
    removed = result["removed"]
    garn = removed[removed["Reason"].str.contains(CAT_GARN, regex=False)] if len(removed) \
        else removed
    sheets = {
        "Corrected Data": result["kept"],
        "Removed Rows": removed if len(removed) else pd.DataFrame({"Message": ["No rows removed"]}),
        "Garnishments": (garn.drop(columns="Reason") if len(garn)
                         else pd.DataFrame({"Message": ["No garnishment rows removed"]})),
        "Flagged Rows": (result["flags"] if len(result["flags"])
                         else pd.DataFrame({"Message": ["No issues found"]})),
        "Change Log": change_log(result),
    }
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as xw:
        for name, frame in sheets.items():
            frame.to_excel(xw, sheet_name=name, index=False)
    return out.getvalue()


# ---------------------------------------------------------------- UI

def _missing_columns_error(info):
    st.markdown(_callout(
        "error", "Required column missing — no file was produced",
        "These columns could not be found: "
        + ", ".join(f"<b>{c}</b>" for c in info["missing"])
        + ". Re-export the Employee Scheduled Deductions report with them included."),
        unsafe_allow_html=True)
    st.caption("Columns found: " + ", ".join(str(c) for c in info["found"]))


def _render_percent_choice(first, tag):
    st.subheader("Percent format")
    if first["is_fraction"]:
        st.caption(f"This file writes percents as fractions ({first['n_frac']} value(s) "
                   "between 0 and 1, e.g. 0.04 = 4%), so they are multiplied by 100.")
    else:
        st.caption(f"Every non-zero percent in this file is already a whole number "
                   f"({first['n_whole']} value(s), e.g. 5 = 5%), so they are left as they are.")
    return st.checkbox("Multiply Percent by 100 (0.04 → 4)", value=first["is_fraction"],
                       key=f"pded_pct_{tag}")


def _render_removal_choices(summary, tag):
    st.subheader("Descriptions to remove")
    st.caption("Tick a description to drop every row that carries it. Garnishments "
               "(support, garnishment, levy, bankruptcy, wage assignment, lien), "
               "reimbursements and TapCheck / ZayZoon / Payactiv start ticked. "
               "`LOA … Reimb` is a real deduction (benefits paid back during a leave) "
               "and starts unticked.")
    chosen = set()
    with st.container(height=400, border=True):
        for i, d in enumerate(summary):
            codes = ", ".join(d["codes"]) or "no code"
            note = f" — *{d['default']}*" if d["default"] else ""
            label = (f"**{d['description'] or '(blank)'}** · code {codes} · "
                     f"{d['rows']} row(s) · {d['employees']} employee(s){note}")
            if st.checkbox(label, value=bool(d["default"]), key=f"pded_rm_{tag}_{i}"):
                chosen.add(d["description"])
    return chosen


def _render_results(result):
    removed, flags = result["removed"], result["flags"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Rows kept", len(result["kept"]))
    c2.metric("Rows removed", len(removed))
    c3.metric("Percent values × 100", len(result["pct_changes"]))
    c4.metric("Rows flagged", flags["Source Row"].nunique() if len(flags) else 0)

    if len(flags):
        id_col = result["cols"][COL_ID]
        lines = []
        for issue, g in flags.groupby("Issue", sort=False):
            ids = sorted({str(v).strip() for v in g[id_col]})
            shown = ", ".join(ids[:8]) + (f" … +{len(ids) - 8} more" if len(ids) > 8 else "")
            lines.append(f"<li><b>{issue}</b> — {len(g)} row(s): {shown}</li>")
        st.markdown(_callout("error", "Needs your attention — nothing below was changed",
                             "<ul style='margin:4px 0 0 0;'>" + "".join(lines) + "</ul>"),
                    unsafe_allow_html=True)
        with st.expander("View the full list of flagged rows"):
            with st.container(height=400, border=True):
                st.dataframe(flags, hide_index=True, width="stretch")
    else:
        st.markdown(_callout("ok", "No issues found",
                             "Every remaining row passes the checks the onboarding API runs."),
                    unsafe_allow_html=True)

    if result["pct_changes"]:
        st.markdown(_callout("ok", "Changed automatically",
                             f"<b>{len(result['pct_changes'])}</b> Percent value(s) multiplied "
                             "by 100 (e.g. 0.04 → 4). Each one is in the Change Log."),
                    unsafe_allow_html=True)

    if len(removed):
        counts = {}
        for reasons in removed["Reason"]:
            for reason in reasons.split("; "):
                counts[reason] = counts.get(reason, 0) + 1
        st.markdown(_callout(
            "warn", "Removed from the corrected file",
            "<br>".join(f"{reason}: <b>{n}</b> row(s)" for reason, n in counts.items())
            + f"<br>Stop Date cutoff: before <b>{result['cutoff']:%m/%d/%Y}</b>; "
              "00/00/0000 and blank Stop Dates stay. A row removed for more than one "
              "reason is counted under each. Every removed row is in the Removed Rows "
              "sheet and the Change Log."),
            unsafe_allow_html=True)
        with st.expander("View the removed rows"):
            with st.container(height=400, border=True):
                st.dataframe(removed, hide_index=True, width="stretch")


def render_ui():
    st.title(APP_TITLE)
    st.markdown("""
    Upload the **Paycom Employee Scheduled Deductions** report (.csv / .xlsx).

    **Rows removed:** Stop Date before the current month (00/00/0000 stays), rows with
    no non-zero Amount or Percent, and the descriptions you tick below.
    **One value changed:** Percent × 100 when the file writes fractions (0.04 → 4).
    Everything else the onboarding API would reject is **flagged** with the EE Codes.

    **You get:** `<Client>_Paycom_Deduction_Corrected.csv` (every original column, plain
    UTF-8) and an `.xlsx` report with Removed Rows, Garnishments, Flagged Rows and the
    Change Log.
    """)
    client_name = st.text_input("Client Name", value="Client", key="paycom_ded_sanity_client")
    f = st.file_uploader("Paycom Employee Scheduled Deductions", type=["csv", "xlsx", "xls"],
                         key="paycom_ded_sanity_upload")
    if f is None:
        return

    first, err = run_sanity(f)
    if err:
        if isinstance(err, dict):
            _missing_columns_error(err)
        else:
            st.error(err)
        return

    tag = zlib.crc32(repr((f.name, getattr(f, "size", None))).encode())
    multiply = _render_percent_choice(first, tag)
    chosen = _render_removal_choices(first["summary"], tag)
    result = (first if chosen == first["remove"] and multiply == first["multiply_percent"]
              else run_sanity(f, chosen, multiply)[0])

    _render_results(result)

    ts = pd.Timestamp.now().strftime("%d_%m_%Y_%H%M")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button("⬇️ Corrected Deduction File (.csv)",
                           data=corrected_csv_bytes(result["kept"]),
                           file_name=f"{client_name}_Paycom_Deduction_Corrected.csv",
                           mime="text/csv", type="primary", key="pded_dl_csv")
    with d2:
        st.download_button("⬇️ Sanity Report (.xlsx)",
                           data=report_xlsx_bytes(result),
                           file_name=f"{client_name}_Paycom_Deduction_Sanity_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="pded_dl_xlsx")

    # Keep the corrected file for the Push to Uzio section below.
    st.session_state["pc_ded_cached"] = {"csv": corrected_csv_bytes(result["kept"]),
                                  "filename": f"{client_name}_Paycom_Deduction_Corrected.csv"}

    # --- Push to Uzio (Production) Section ---
    from utils.onboarding_api_client import render_push_deductions_section
    render_push_deductions_section(vendor="PAYCOM", data_key="pc_ded_cached",
                                   key_prefix="pc_ded_push2uzio")


if __name__ == "__main__":
    render_ui()
