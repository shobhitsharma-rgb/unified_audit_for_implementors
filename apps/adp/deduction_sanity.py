import csv
import io
import re
import zlib

import openpyxl
import pandas as pd
import streamlit as st

from utils.ui_components import _callout

APP_TITLE = "ADP Deduction Sanity Check"

# =========================================================
# ADP Deduction Sanity Check
# - Input: one ADP Voluntary Deduction export (.xlsx / .csv)
# - Removes rows only: the "Report Totals" / blank-Associate-ID rows, and
#   every row whose DEDUCTION DESCRIPTION the user ticks (direct deposit,
#   garnishments and earned-wage / reimbursement lines are pre-ticked).
# - Never edits a value. Anything the onboarding API would reject is FLAGGED
#   with the affected Associate IDs, and left for the user to fix.
#
# The rules flagged here are the onboarding API's own
# (ADPConfig.toPaycomDeductionRecord -> EmployeeDeductionValidator,
#  EmployeeDeductionSetUpServiceImpl):
#   - it reads only the five REQUIRED_COLUMNS, by header name
#   - the mapping file is looked up by DEDUCTION DESCRIPTION, exactly
#   - exactly one of amount / percent; amount >= 0; 0 <= percent <= 100;
#     each must parse as a plain number (no $ or %)
#   - an Associate ID unknown to UZIO fails every row of that employee
# =========================================================

COL_ID = "ASSOCIATE ID"
COL_CODE = "DEDUCTION CODE"
COL_DESC = "DEDUCTION DESCRIPTION"
COL_AMT = "DEDUCTION AMOUNT"
COL_PCT = "DEDUCTION %"
REQUIRED_COLUMNS = [COL_ID, COL_CODE, COL_DESC, COL_AMT, COL_PCT]

CAT_DD = "Direct deposit"
CAT_GARN = "Garnishment"
CAT_EWA = "Earned-wage access / reimbursement"
CAT_USER = "Ticked by you"
CAT_TOTALS = "Report Totals row"
CAT_BLANK_ID = "Blank Associate ID"
CAT_DUP = "Identical duplicate"
DUP_OF = "Duplicate Of Row"

# A plain decimal, which is what BigDecimal (the API's parser) accepts once
# commas are dropped. "$25.00" and "3%" fail it, as they fail the API.
_NUMBER = re.compile(r"^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$")
_SSN = re.compile(r"^\d{3}-?\d{2}-?\d{4}$")


def _s(v):
    """Cell value as the text it holds; blank for None."""
    return "" if v is None else str(v)


# ---------------------------------------------------------------- reading

def _cell_text(v):
    """An openpyxl value as text, the way the sheet shows it.

    ADP writes these columns as text ("3.0000"), which passes through
    untouched. A genuine number cell loses only a meaningless ".0".
    """
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


def _find_header(rows, limit=20):
    """Index of the first row carrying both DEDUCTION CODE and DEDUCTION DESCRIPTION."""
    for i, row in enumerate(rows[:limit]):
        cells = {str(c).strip().upper() for c in row if str(c).strip()}
        if COL_CODE in cells and COL_DESC in cells:
            return i
    return None


def _frame(rows, header_idx):
    head = [str(h) for h in rows[header_idx]]
    body = []
    for r in rows[header_idx + 1:]:
        r = list(r) + [""] * (len(head) - len(r))
        if any(str(c).strip() for c in r):          # a wholly blank line carries nothing
            body.append([str(c) for c in r[:len(head)]])
    return pd.DataFrame(body, columns=head, dtype=str)


def read_deduction_file(file):
    """The ADP export as all-text rows, preamble skipped.

    Returns ({"df", "header_row", "sheet"}, error). `header_row` is the
    1-based row the header sits on, so a source row number can be shown.
    Values are kept exactly as text — nothing is parsed or trimmed — because
    the corrected file must carry them unchanged.
    """
    name = (getattr(file, "name", "") or "").lower()
    data = file.getvalue()
    sheets = []
    try:
        if data[:2] == b"PK":
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            for ws in wb.worksheets:
                sheets.append((ws.title, [[_cell_text(v) for v in row]
                                          for row in ws.iter_rows(values_only=True)]))
        elif data[:4] == bytes.fromhex("d0cf11e0") or name.endswith(".xls"):
            for title, df in pd.read_excel(io.BytesIO(data), sheet_name=None, header=None,
                                           dtype=str, keep_default_na=False).items():
                sheets.append((title, df.values.tolist()))
        else:
            try:
                text = data.decode("utf-8-sig")
            except UnicodeDecodeError:
                text = data.decode("cp1252", errors="replace")
            sheets.append(("CSV", list(csv.reader(io.StringIO(text)))))
    except Exception as e:
        return None, f"Could not read the file: {e}"

    for title, rows in sheets:
        h = _find_header(rows)
        if h is not None:
            return {"df": _frame(rows, h), "header_row": h + 1, "sheet": title}, None
    return None, ("Could not find a header row with `DEDUCTION CODE` and "
                  "`DEDUCTION DESCRIPTION` in the first 20 rows.")


def column_map(df):
    """{required name: actual header} and the required names that are missing."""
    found = {str(c).strip().upper(): c for c in df.columns}
    cols = {r: found[r] for r in REQUIRED_COLUMNS if r in found}
    return cols, [r for r in REQUIRED_COLUMNS if r not in cols]


# ---------------------------------------------------------------- removal

def default_category(description):
    """Which pre-ticked group a description falls in, or None.

    Direct deposit is matched on the WHOLE description, so "HSA SAVINGS" (a
    real deduction) is not caught by "SAVINGS". Case and "%" are ignored.

    Garnishments include BANKRUPTCY (ADP code 70) and SPT/WAGEAGREMNT (codes
    75/76, the same codes SUPPORT uses); reimbursements include MILEAGE REIMB,
    whose negative amount the API rejects anyway.
    """
    words = re.sub(r"[^A-Z]+", " ", _s(description).upper()).split()
    joined, compact = " ".join(words), "".join(words)
    if joined in ("CHECKING", "SAVINGS"):
        return CAT_DD
    if ("SUPPORT" in words or "LEVY" in words or "SPT" in words or "WAGEAGREMNT" in compact
            or any(w.startswith(("GARNISH", "BANKRUPT")) for w in words)):
        return CAT_GARN
    if any(k in compact for k in ("PAYACTIV", "ZAYZOON", "TAPCHECK", "REIMB")):
        return CAT_EWA
    return None


def description_summary(df, cols):
    """One entry per distinct DEDUCTION DESCRIPTION (exactly as written), in file order.

    Rows with a blank Associate ID are left out: they are removed regardless.
    """
    live = df[df[cols[COL_ID]].str.strip() != ""]
    out = []
    for desc, g in live.groupby(cols[COL_DESC], sort=False):
        codes = sorted({c.strip() for c in g[cols[COL_CODE]] if c.strip()})
        out.append({"description": desc, "codes": codes, "rows": len(g),
                    "employees": g[cols[COL_ID]].str.strip().nunique(),
                    "default": default_category(desc)})
    return out


def _source_rows(df, header_row):
    return pd.Series(range(header_row + 1, header_row + 1 + len(df)), index=df.index)


def _duplicate_groups(df, cols):
    """[(group index list, identical?)] for every employee + description seen twice or more.

    "Identical" means every row of the group carries the same Associate ID, code,
    description, amount and % (spaces ignored) — typically one employee holding two
    positions. A group with any difference is NOT identical as a whole, because
    which row is right is the client's call.
    """
    key = df[cols[COL_ID]].str.strip() + "\x1f" + df[cols[COL_DESC]].str.strip()
    api = [cols[c] for c in REQUIRED_COLUMNS]
    out = []
    for _, g in df.groupby(key, sort=False):
        if len(g) > 1:
            same = len(g[api].apply(lambda col: col.str.strip()).drop_duplicates()) == 1
            out.append((list(g.index), same))
    return out


def apply_removals(df, cols, header_row, remove_descriptions, remove_duplicates=False):
    """Split the file into kept rows and removed rows (with a reason each).

    `remove_descriptions` is the set of descriptions the user left ticked.
    `remove_duplicates` keeps the FIRST row of each identical-duplicate group and
    removes the rest; groups whose rows differ stay, and are flagged.
    """
    src = _source_rows(df, header_row)
    blank_id = df[cols[COL_ID]].str.strip() == ""
    totals = df.apply(lambda r: any(_s(v).strip().lower().startswith("report totals")
                                    for v in r), axis=1)
    reason = pd.Series("", index=df.index)
    reason[blank_id & totals] = CAT_TOTALS
    reason[blank_id & ~totals] = CAT_BLANK_ID
    ticked = ~blank_id & df[cols[COL_DESC]].isin(remove_descriptions)
    reason[ticked] = df.loc[ticked, cols[COL_DESC]].map(
        lambda d: default_category(d) or CAT_USER)

    dup_of = pd.Series("", index=df.index, dtype=object)
    if remove_duplicates:
        for idx, identical in _duplicate_groups(df[reason == ""], cols):
            if identical:
                for i in idx[1:]:
                    reason[i] = CAT_DUP
                    # text, like the blanks around it: a column mixing "" and ints
                    # cannot be shown by st.dataframe (pyarrow refuses it)
                    dup_of[i] = str(int(src[idx[0]]))

    gone = reason != ""
    removed = df[gone].copy()
    removed.insert(0, DUP_OF, dup_of[gone])
    removed.insert(0, "Reason", reason[gone])
    removed.insert(0, "Source Row", src[gone])
    kept = df[~gone].copy()
    return kept, removed.reset_index(drop=True)


# ---------------------------------------------------------------- flags

def find_flags(kept, cols, header_row, src_rows=None):
    """Every API-breaking issue on the rows that stay. Nothing is changed.

    Returns a DataFrame: Source Row, Issue, then the row's own columns.
    """
    src = src_rows if src_rows is not None else _source_rows(kept, header_row)
    issues = []

    def add(idx, issue):
        issues.append((idx, issue))

    for idx, r in kept.iterrows():
        aid, code, desc = r[cols[COL_ID]], r[cols[COL_CODE]], r[cols[COL_DESC]]
        amt, pct = r[cols[COL_AMT]].strip(), r[cols[COL_PCT]].strip()
        if not code.strip():
            add(idx, "Deduction code is blank")
        if not desc.strip():
            add(idx, "Deduction description is blank")
        spaced = [n for n, v in (("Associate ID", aid), ("code", code), ("description", desc))
                  if v != v.strip()]
        if spaced:
            add(idx, "Extra spaces around " + " / ".join(spaced)
                + " (the mapping is matched on the exact description)")
        if not amt and not pct:
            add(idx, "Amount and % are both blank")
        elif amt and pct:
            add(idx, "Amount and % are both filled (the API takes only one)")
        for label, v in (("Amount", amt), ("%", pct)):
            if v and not _NUMBER.match(v.replace(",", "")):
                add(idx, f"{label} is not a plain number (e.g. a $ or % sign)")
        if amt and _NUMBER.match(amt.replace(",", "")) and float(amt.replace(",", "")) < 0:
            add(idx, "Amount is negative")
        if pct and _NUMBER.match(pct.replace(",", "")):
            p = float(pct.replace(",", ""))
            if p < 0 or p > 100:
                add(idx, "% is outside 0-100")
        if _SSN.match(aid.strip()):
            add(idx, "Associate ID looks like an SSN")

    # Duplicates still present: identical ones when "Remove identical duplicates" is
    # off, and differing ones always (which row is right is the client's call).
    for idx_list, identical in _duplicate_groups(kept, cols):
        label = ("Duplicate - identical rows for the same employee and description"
                 if identical else
                 "Duplicate - same employee and description, different code or amount")
        for idx in idx_list:
            add(idx, label)

    if not issues:
        return pd.DataFrame(columns=["Source Row", "Issue"] + list(kept.columns))
    order = {idx: n for n, idx in enumerate(kept.index)}
    issues.sort(key=lambda t: order[t[0]])
    out = pd.DataFrame([{"Source Row": int(src[idx]), "Issue": issue,
                         **kept.loc[idx].to_dict()} for idx, issue in issues])
    return out


# ---------------------------------------------------------------- run

def run_sanity(file, remove_descriptions=None, remove_duplicates=True):
    """Read, remove, flag. `remove_descriptions=None` means the default ticks.

    `remove_duplicates` (on by default, Rohit's call) drops all but the first row
    of each identical-duplicate group.

    Returns (result, error). result = {"df", "cols", "header_row", "summary",
    "kept", "removed", "flags", "remove", "remove_duplicates"}.
    """
    info, err = read_deduction_file(file)
    if err:
        return None, err
    df = info["df"]
    cols, missing = column_map(df)
    if missing:
        return None, {"missing": missing, "found": list(df.columns)}
    summary = description_summary(df, cols)
    if remove_descriptions is None:
        remove_descriptions = {d["description"] for d in summary if d["default"]}
    kept, removed = apply_removals(df, cols, info["header_row"], set(remove_descriptions),
                                   remove_duplicates)
    src = _source_rows(df, info["header_row"])
    flags = find_flags(kept, cols, info["header_row"], src[kept.index])
    return {"df": df, "cols": cols, "header_row": info["header_row"], "summary": summary,
            "kept": kept.reset_index(drop=True), "removed": removed, "flags": flags,
            "remove": set(remove_descriptions),
            "remove_duplicates": bool(remove_duplicates)}, None


def corrected_csv_bytes(kept):
    """The kept rows, every original column, values untouched. Plain UTF-8, NO BOM."""
    return kept.to_csv(index=False).encode("utf-8")


def change_log(result):
    cols = result["cols"]
    name_col = next((c for c in result["df"].columns if str(c).strip().upper() == "NAME"), None)
    rows = []
    if result["header_row"] > 1:
        rows.append({"Source Row": "(file)", "Associate ID": "(All rows)", "Name": "",
                     "Deduction Code": "", "Deduction Description": "",
                     "Action": "Rows above the header removed",
                     "Reason": f"{result['header_row'] - 1} preamble row(s); the API "
                               "expects the header on row 1"})
    for _, r in result["removed"].iterrows():
        rows.append({"Source Row": r["Source Row"], "Associate ID": r[cols[COL_ID]],
                     "Name": r[name_col] if name_col else "",
                     "Deduction Code": r[cols[COL_CODE]],
                     "Deduction Description": r[cols[COL_DESC]],
                     "Action": "Row removed",
                     "Reason": (f"{CAT_DUP} - row {r[DUP_OF]} kept" if r["Reason"] == CAT_DUP
                                else r["Reason"])})
    return pd.DataFrame(rows, columns=["Source Row", "Associate ID", "Name", "Deduction Code",
                                       "Deduction Description", "Action", "Reason"])


def report_xlsx_bytes(result):
    removed = result["removed"]
    garn = removed[removed["Reason"] == CAT_GARN]
    sheets = {
        "Corrected Data": result["kept"],
        "Removed Rows": removed if len(removed) else pd.DataFrame({"Message": ["No rows removed"]}),
        "Garnishments": (garn.drop(columns=["Reason", DUP_OF]) if len(garn)
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
        "The onboarding API reads these columns by name: "
        + ", ".join(f"<b>{c}</b>" for c in info["missing"])
        + " could not be found. Re-export the Voluntary Deduction report with them included."),
        unsafe_allow_html=True)
    st.caption("Columns found: " + ", ".join(str(c) for c in info["found"]))


def _render_removal_choices(summary, tag):
    """One checkbox per description; returns the set left ticked."""
    st.subheader("Rows to remove")
    st.caption("Tick a description to drop every row that carries it. Direct deposit "
               "(CHECKING / SAVINGS), garnishments (SUPPORT / GARNISHMENT / TAX LEVY) and "
               "Payactiv / ZayZoon / Tapcheck / reimbursement lines start ticked. "
               "Rows with no Associate ID (the Report Totals line) are always removed.")
    chosen = set()
    with st.container(height=400, border=True):
        for i, d in enumerate(summary):
            codes = ", ".join(d["codes"]) or "no code"
            note = f" — *{d['default']}*" if d["default"] else ""
            label = (f"**{d['description'] or '(blank)'}** · code {codes} · "
                     f"{d['rows']} row(s) · {d['employees']} employee(s){note}")
            if st.checkbox(label, value=bool(d["default"]), key=f"ded_rm_{tag}_{i}"):
                chosen.add(d["description"])
    return chosen


def _render_results(result):
    removed, flags = result["removed"], result["flags"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Rows kept", len(result["kept"]))
    c2.metric("Rows removed", len(removed))
    c3.metric("Rows flagged", flags["Source Row"].nunique() if len(flags) else 0)

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

    if len(removed):
        counts = removed["Reason"].value_counts()
        st.markdown(_callout("warn", "Removed from the corrected file",
                             "<br>".join(f"{reason}: <b>{n}</b> row(s)"
                                         for reason, n in counts.items())
                             + "<br>Every removed row is in the Removed Rows sheet and the "
                               "Change Log; garnishments also get their own sheet, to be set "
                               "up in UZIO separately."),
                    unsafe_allow_html=True)
        with st.expander("View the removed rows"):
            with st.container(height=400, border=True):
                st.dataframe(removed, hide_index=True, width="stretch")


def render_ui():
    st.title(APP_TITLE)
    st.markdown("""
    Upload the **ADP Voluntary Deduction** report (.xlsx / .csv).

    - **Rows are removed, values are never changed.** Pick the descriptions to drop
      below; the Report Totals line always goes, and identical duplicate rows are cut
      to one unless you untick that box.
    - Anything the onboarding API would reject — blank fields, amount *and* %,
      non-numeric values, extra spaces, duplicates, SSN-like IDs — is **flagged** with
      the Associate IDs, for you to fix.

    **You get:** `<Client>_ADP_Deduction_Corrected.csv` (every original column, plain
    UTF-8, ready for the API) and an `.xlsx` report with Removed Rows, Garnishments,
    Flagged Rows and the Change Log.
    """)
    client_name = st.text_input("Client Name", value="Client", key="adp_ded_sanity_client")
    f = st.file_uploader("ADP Voluntary Deduction report", type=["xlsx", "xls", "csv"],
                         key="adp_ded_sanity_upload")
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
    chosen = _render_removal_choices(first["summary"], tag)
    dedupe = st.checkbox(
        "Remove identical duplicates (keep the first row)", value=True, key=f"ded_dup_{tag}",
        help="When one employee has the same deduction twice with the same code, "
             "description, amount and % (usually two positions), only the first row is "
             "kept. Rows of the same deduction with a DIFFERENT code or amount are never "
             "removed — which one is right is for the client to say — and stay flagged.")
    result = (first if chosen == first["remove"] and dedupe == first["remove_duplicates"]
              else run_sanity(f, chosen, dedupe)[0])

    _render_results(result)

    ts = pd.Timestamp.now().strftime("%d_%m_%Y_%H%M")
    d1, d2 = st.columns(2)
    with d1:
        st.download_button("⬇️ Corrected Deduction File (.csv)",
                           data=corrected_csv_bytes(result["kept"]),
                           file_name=f"{client_name}_ADP_Deduction_Corrected.csv",
                           mime="text/csv", type="primary", key="ded_sanity_dl_csv")
    with d2:
        st.download_button("⬇️ Sanity Report (.xlsx)",
                           data=report_xlsx_bytes(result),
                           file_name=f"{client_name}_ADP_Deduction_Sanity_{ts}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           key="ded_sanity_dl_xlsx")

    # Keep the corrected file for the Push to Uzio section below.
    st.session_state["adp_ded_cached"] = {"csv": corrected_csv_bytes(result["kept"]),
                                  "filename": f"{client_name}_ADP_Deduction_Corrected.csv"}

    # --- Push to Uzio (Production) Section ---
    from utils.onboarding_api_client import render_push_deductions_section
    render_push_deductions_section(vendor="ADP", data_key="adp_ded_cached",
                                   key_prefix="adp_ded_push2uzio")


if __name__ == "__main__":
    render_ui()
