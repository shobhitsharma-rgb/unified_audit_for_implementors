"""ADP - Prior Payroll Setup Helper Tool.

Discovers what to configure in Uzio for an ADP Prior Payroll migration.
Given a sanitized ADP Prior Payroll file plus the State Tax Code master CSV,
emits an Excel workbook with:

  - Earnings_Codes      - REGULAR / OVERTIME + every ADDITIONAL EARNINGS code
                          with $ total, employee count, paired hours, avg rate.
  - Contributions       - 401k / 403b / 457 / Roth / HSA / FSA voluntary codes.
  - Deductions          - all other voluntary deductions, with pre-tax vs
                          post-tax verdict per code.
  - Taxes_Discovered    - every '* - EMPLOYEE TAX' / '* - EMPLOYER TAX' column.
  - Tax_Mapping         - one row per (tax_type, state) in the
                          'Payroll_Mappings_Tax_Mapping_CORRECTED' format.
  - Bonus_Classification- FLSA test (discretionary vs non-discretionary).

Pre/post-tax algorithm:
  gap_FIT = TOTAL EARNINGS - FEDERAL INCOME - EMPLOYEE TAXABLE.
  Find any subset of a row's non-zero deductions summing to gap_FIT (within
  $0.02). Every member of any passing subset is pre-tax for FIT. ONE positive
  proof anywhere in the file = pre-tax for everyone (the rule never varies
  per employee). Same logic for FICA / MEDI / SIT taxables to derive the
  flavor: section_125 (pre-FIT/FICA/MEDI/SIT) vs 401k_traditional
  (pre-FIT/SIT only, NOT pre-FICA/MEDI).
"""

import io
import re
from itertools import combinations

import pandas as pd
import streamlit as st

from apps.adp.prior_payroll_sanity import read_input_file, _find_col
from utils.audit_utils import clean_money_val


# ---------- helpers ----------

def _num(v):
    try:
        return clean_money_val(v)
    except Exception:
        return 0.0


def _strip_prefix(col, prefixes):
    s = str(col).strip()
    for p in prefixes:
        if s.upper().startswith(p.upper()):
            rest = s[len(p):].lstrip(" :").strip()
            return rest
    return s


# ---------- column categorization ----------

EARN_PREFIXES = ["ADDITIONAL EARNINGS"]
HOUR_PREFIXES = ["ADDITIONAL HOURS"]
DED_PREFIX = "VOLUNTARY DEDUCTION"
MEMO_PREFIX = "MEMO"

CONTRIB_PATTERN = re.compile(
    r"\b(401[Kk]?|403[Bb]?|457|ROTH|HSA|FSA|RETIRE|RETIREMENT)\b"
)


def categorize_columns(df):
    earn_cols, hour_cols, tax_cols, taxable_cols = [], [], [], []
    ded_cols, memo_cols = [], []
    for c in df.columns:
        s = str(c).strip(); u = s.upper()
        if u in ("REGULAR EARNINGS", "OVERTIME EARNINGS"):
            earn_cols.append(c)
        elif u.startswith("ADDITIONAL EARNINGS"):
            earn_cols.append(c)
        elif u in ("REGULAR HOURS", "OVERTIME HOURS"):
            hour_cols.append(c)
        elif u.startswith("ADDITIONAL HOURS"):
            hour_cols.append(c)
        elif u.startswith(DED_PREFIX):
            ded_cols.append(c)
        elif u.startswith(MEMO_PREFIX):
            memo_cols.append(c)
        elif u.endswith("TAXABLE"):
            taxable_cols.append(c)
        elif u.endswith("EMPLOYEE TAX") or u.endswith("EMPLOYER TAX"):
            if u.startswith("TOTAL "):
                continue  # aggregate, not a real tax row
            tax_cols.append(c)
    return {
        "earnings": earn_cols, "hours": hour_cols,
        "taxes": tax_cols, "taxables": taxable_cols,
        "deductions": ded_cols, "memos": memo_cols,
    }


# ---------- catalog builders ----------

def build_earnings_catalog(df, earn_cols, hour_cols):
    hour_lookup = {}
    for h in hour_cols:
        u = str(h).strip().upper()
        if u == "REGULAR HOURS":
            hour_lookup["REGULAR EARNINGS"] = h
        elif u == "OVERTIME HOURS":
            hour_lookup["OVERTIME EARNINGS"] = h
        else:
            code = _strip_prefix(h, HOUR_PREFIXES)
            hour_lookup[f"ADDITIONAL EARNINGS  : {code}"] = h
            hour_lookup[code] = h

    rows = []
    for c in earn_cols:
        amounts = df[c].apply(_num)
        total = float(amounts.sum())
        emp_count = int((amounts != 0).sum())
        u = str(c).strip().upper()
        if u == "REGULAR EARNINGS":
            code = "REGULAR"; kind = "Regular Wage"
        elif u == "OVERTIME EARNINGS":
            code = "OVERTIME"; kind = "Overtime"
        else:
            code = _strip_prefix(c, EARN_PREFIXES); kind = "Additional Earning"

        h_col = hour_lookup.get(str(c).strip()) or hour_lookup.get(code)
        if h_col is not None and h_col in df.columns:
            hours_total = float(df[h_col].apply(_num).sum())
            avg_rate = total / hours_total if hours_total > 0 else None
        else:
            hours_total = None; avg_rate = None

        rows.append({
            "Source Column": str(c).strip(), "Code": code, "Kind": kind,
            "Total $": round(total, 2), "Employees": emp_count,
            "Total Hours": round(hours_total, 2) if hours_total is not None else None,
            "Avg Rate ($/hr)": round(avg_rate, 4) if avg_rate is not None else None,
        })
    return rows


# ---------- pre/post-tax classifier ----------

def _row_gap(row, total_earn_col, taxable_col):
    return _num(row.get(total_earn_col)) - _num(row.get(taxable_col))


def _subset_sum_match(amounts, target, tol=0.02):
    n = len(amounts)
    if n == 0:
        return []
    matches = []
    for r in range(1, n + 1):
        for combo in combinations(range(n), r):
            s = sum(amounts[i] for i in combo)
            if abs(s - target) <= tol:
                matches.append(combo)
    return matches


def _name_heuristic(col):
    u = str(col).upper()
    if any(t in u for t in ("SUPPORT", "GARN", "GARNISH", "LEVY", "LIEN", "CHILD")):
        return "post_tax", "garnishment", [], "name_heuristic"
    if any(t in u for t in ("ADVANCE", "ADV-", "LOAN", "REPAY", "TAPCHECK", "DAILY")):
        return "post_tax", "advance_or_loan", [], "name_heuristic"
    if any(t in u for t in ("REVERSE", "REV-", "REISSU")):
        return "post_tax", "corrective", [], "name_heuristic"
    if any(t in u for t in ("ROTH",)):
        return "post_tax", "roth", [], "name_heuristic"
    if any(t in u for t in ("MEDICAL", "MED-", "DENTAL", "DEN-", "VISION", "VIS-",
                            "HSA", "FSA")):
        return "pre_tax", "section_125", ["FIT", "FICA", "MEDI", "SIT"], "name_heuristic"
    if CONTRIB_PATTERN.search(u):
        return "pre_tax", "401k_traditional", ["FIT", "SIT"], "name_heuristic"
    return "post_tax", "default_unknown", [], "name_heuristic"


def classify_deductions_pretax(
    df, ded_cols, total_earn_col, fit_taxable_col, fica_taxable_col,
    medi_taxable_col, sit_taxable_col, tol=0.02, max_subset=8,
):
    proven = {c: {"FIT": False, "FICA": False, "MEDI": False, "SIT": False}
              for c in ded_cols}
    sample = {c: [] for c in ded_cols}

    def _try_axis(taxable_col, key):
        if taxable_col is None:
            return
        for _, row in df.iterrows():
            gap = _row_gap(row, total_earn_col, taxable_col)
            if gap <= tol:
                continue
            present = [(c, _num(row.get(c))) for c in ded_cols if _num(row.get(c)) > 0]
            if not present or len(present) > max_subset:
                continue
            cols = [c for c, _ in present]
            amts = [a for _, a in present]
            for combo in _subset_sum_match(amts, gap, tol):
                for i in combo:
                    proven[cols[i]][key] = True
                if key == "FIT":
                    eid = row.get("ASSOCIATE ID") or row.get("Associate ID")
                    for i in combo:
                        if len(sample[cols[i]]) < 3:
                            sample[cols[i]].append({
                                "associate": str(eid) if eid is not None else "",
                                "gap_fit": round(gap, 2),
                                "subset": [cols[j] for j in combo],
                                "subset_sum": round(sum(amts[j] for j in combo), 2),
                            })

    _try_axis(fit_taxable_col, "FIT")
    _try_axis(fica_taxable_col, "FICA")
    _try_axis(medi_taxable_col, "MEDI")
    _try_axis(sit_taxable_col, "SIT")

    rows = []
    for c in ded_cols:
        amounts = df[c].apply(_num)
        total = float(amounts.sum())
        emp_count = int((amounts != 0).sum())
        p = proven[c]
        if p["FIT"] and p["FICA"] and p["MEDI"]:
            verdict = "pre_tax"; flavor = "section_125"
            pre_taxes = ["FIT", "FICA", "MEDI"] + (["SIT"] if p["SIT"] else [])
        elif p["FIT"] and p["SIT"] and not p["FICA"]:
            verdict = "pre_tax"; flavor = "401k_traditional"; pre_taxes = ["FIT", "SIT"]
        elif p["FIT"] and not (p["FICA"] or p["MEDI"]):
            verdict = "pre_tax"; flavor = "pretax_unknown"; pre_taxes = ["FIT"]
        elif p["FIT"] or p["FICA"] or p["MEDI"] or p["SIT"]:
            verdict = "pre_tax"; flavor = "mixed_unusual"
            pre_taxes = [k for k in ("FIT", "FICA", "MEDI", "SIT") if p[k]]
        else:
            verdict = "post_tax"; flavor = ""; pre_taxes = []

        if emp_count == 0:
            verdict, flavor, pre_taxes, confidence = _name_heuristic(c)
        else:
            confidence = "empirical_subset_sum"

        code = _strip_prefix(c, [DED_PREFIX])
        is_contrib = bool(CONTRIB_PATTERN.search(code.upper()))
        rows.append({
            "Source Column": str(c).strip(), "Code": code,
            "Total $": round(total, 2), "Employees": emp_count,
            "Verdict": verdict, "Pre-Tax Of": ",".join(pre_taxes),
            "Pre-Tax Flavor": flavor, "Confidence": confidence,
            "Sample": "; ".join(
                f"{s['associate']}: gap={s['gap_fit']}, subset_sum={s['subset_sum']}"
                for s in sample[c][:2]
            ),
            "_is_contribution": is_contrib,
        })
    return rows


# ---------- bonus classifier (FLSA) ----------

def classify_bonus(df, earn_cols):
    reg_e = _find_col(df, ["REGULAR EARNINGS"])
    reg_h = _find_col(df, ["REGULAR HOURS"])
    ot_e = _find_col(df, ["OVERTIME EARNINGS"])
    ot_h = _find_col(df, ["OVERTIME HOURS"])

    bonus_cols = []
    for c in earn_cols:
        u = str(c).upper()
        code = _strip_prefix(c, EARN_PREFIXES).upper()
        if "BONUS" in u or re.search(r"\bBN[A-Z0-9]*\b", code) or code.startswith("BN"):
            if "BACKUP" in u or code.startswith("BCK"):
                continue
            bonus_cols.append(c)

    if not bonus_cols or not (reg_e and reg_h and ot_e and ot_h):
        return {
            "verdict": "indeterminate",
            "reason": "Missing bonus / overtime columns to test",
            "bonus_columns_found": [str(c) for c in bonus_cols],
            "rows_tested": 0, "discretionary_rows": 0, "non_discretionary_rows": 0,
            "samples": [],
        }

    rows_tested = 0; discretionary_rows = 0; non_disc_rows = 0
    samples = []
    rate_tol_pct = 0.005

    for _, r in df.iterrows():
        bonus_amt = sum(_num(r.get(c)) for c in bonus_cols)
        re_v = _num(r.get(reg_e)); rh_v = _num(r.get(reg_h))
        oe_v = _num(r.get(ot_e)); oh_v = _num(r.get(ot_h))
        if bonus_amt <= 0 or oh_v <= 0 or rh_v <= 0 or re_v <= 0:
            continue
        rows_tested += 1
        regular_rate = re_v / rh_v
        expected_ot_rate = 1.5 * regular_rate
        actual_ot_rate = oe_v / oh_v
        diff_pct = (actual_ot_rate - expected_ot_rate) / expected_ot_rate

        verdict_row = "discretionary"
        if diff_pct > rate_tol_pct:
            verdict_row = "non_discretionary"; non_disc_rows += 1
        else:
            discretionary_rows += 1

        if len(samples) < 5:
            eid = r.get("ASSOCIATE ID") or r.get("Associate ID")
            samples.append({
                "associate": str(eid) if eid is not None else "",
                "regular_earnings": round(re_v, 2), "regular_hours": round(rh_v, 4),
                "regular_rate": round(regular_rate, 4),
                "expected_ot_rate_1.5x": round(expected_ot_rate, 4),
                "actual_ot_rate": round(actual_ot_rate, 4),
                "diff_pct": round(diff_pct * 100, 3),
                "bonus_amt": round(bonus_amt, 2), "verdict_row": verdict_row,
            })

    if rows_tested == 0:
        verdict = "indeterminate"
        reason = "No row had both bonus and overtime hours"
    elif non_disc_rows > 0:
        verdict = "non_discretionary"
        reason = (
            f"{non_disc_rows} of {rows_tested} rows show actual OT rate "
            f"materially above 1.5 x regular rate => bonus inflated regular rate => "
            f"non-discretionary (any positive proof is conclusive under FLSA)."
        )
    else:
        verdict = "discretionary"
        reason = (
            f"All {rows_tested} rows show actual OT rate ~ 1.5 x regular rate => "
            f"bonus did not inflate the regular rate basis => discretionary."
        )

    return {
        "verdict": verdict, "reason": reason,
        "bonus_columns_found": [str(c) for c in bonus_cols],
        "rows_tested": rows_tested,
        "discretionary_rows": discretionary_rows,
        "non_discretionary_rows": non_disc_rows,
        "samples": samples,
    }


# ---------- tax mapping ----------

TAX_TOKEN_MAP = {
    "FEDERAL INCOME - EMPLOYEE TAX":          ("FED", "FIT"),
    "MEDICARE - EMPLOYEE TAX":                ("FED", "MEDI"),
    "MEDICARE - EMPLOYER TAX":                ("FED", "ER_MEDI"),
    "SOCIAL SECURITY - EMPLOYEE TAX":         ("FED", "FICA"),
    "SOCIAL SECURITY - EMPLOYER TAX":         ("FED", "ER_FICA"),
    "FUTA - EMPLOYER TAX":                    ("FED", "ER_FUTA"),
    "WORKED IN STATE - EMPLOYEE TAX":         ("STATE", "SIT"),
    "SUI/SDI - EMPLOYEE TAX":                 ("STATE", "SDI"),
    "SUI/SDI - EMPLOYER TAX":                 ("STATE", "ER_SUTA"),
    "FAMILY LEAVE INSURANCE - EMPLOYEE TAX":  ("STATE", "FLI"),
}


def lookup_canonical_tax(master_df, state_abbr, type_code):
    if master_df is None:
        return None
    pat = re.compile(rf"^\d{{2}}-000-0000-{re.escape(type_code)}-000$")
    sub = master_df[master_df["state_abbreviation"].astype(str).str.upper()
                    == state_abbr.upper()]
    if sub.empty:
        return None
    sub2 = sub[sub["unique_tax_id"].astype(str).apply(lambda s: bool(pat.match(s)))]
    if sub2.empty:
        broad = master_df[
            (master_df["state_abbreviation"].astype(str).str.upper() == state_abbr.upper())
            & master_df["unique_tax_id"].astype(str).str.contains(f"-{type_code}-", regex=False)
        ]
        if broad.empty:
            return None
        primary = broad[broad["sub_tax_desc"].fillna("").astype(str).str.strip() == ""]
        return primary.iloc[0] if not primary.empty else broad.iloc[0]
    primary = sub2[sub2["sub_tax_desc"].fillna("").astype(str).str.strip() == ""]
    return primary.iloc[0] if not primary.empty else sub2.iloc[0]


def build_tax_mapping(df, tax_cols, master_df):
    state_col = _find_col(df, ["WORKED IN STATE", "Worked In State", "State"])
    states = []
    if state_col:
        for v in df[state_col].dropna().astype(str):
            s = v.strip().upper()
            if s and s not in states and len(s) == 2:
                states.append(s)
    if not states:
        states = ["NY"]

    out_rows = []; not_found = []
    for tcol in tax_cols:
        key = str(tcol).strip().upper()
        scope_type = TAX_TOKEN_MAP.get(key)
        if not scope_type:
            not_found.append({"tax_column": str(tcol), "reason": "no rule defined"})
            continue
        scope, type_code = scope_type
        targets = ["FED"] if scope == "FED" else states
        for st_code in targets:
            rec = lookup_canonical_tax(master_df, st_code, type_code)
            if rec is None:
                not_found.append({"tax_column": str(tcol),
                                  "reason": f"{st_code} {type_code} not in master"})
                continue
            out_rows.append({
                "Source Tax Code": "",
                "Source Tax Code Name": str(tcol),
                "Source Tax Code Description": "",
                "Uzio Tax Code": rec.get("tax_code", ""),
                "Unique Tax ID": rec.get("unique_tax_id", ""),
                "Uzio Tax Code Description": rec.get("tax_name", ""),
                "Uzio Sub-Tax Description": rec.get("sub_tax_desc", "") or "",
            })
    return out_rows, states, not_found


def tax_mapping_to_csv_bytes(rows):
    cols = ["Source Tax Code", "Source Tax Code Name", "Source Tax Code Description",
            "Uzio Tax Code", "Unique Tax ID", "Uzio Tax Code Description",
            "Uzio Sub-Tax Description"]
    return pd.DataFrame(rows, columns=cols).to_csv(index=False).encode("utf-8")


# ---------- orchestrator ----------

def run_setup_helper(adp_file, master_csv_file):
    """Run the analysis. Returns (results_dict_of_lists, tax_csv_bytes, df)."""
    df, _, _ = read_input_file(adp_file)
    df = df.reset_index(drop=True)

    cats = categorize_columns(df)
    earn_rows = build_earnings_catalog(df, cats["earnings"], cats["hours"])

    total_earn_col = _find_col(df, ["TOTAL EARNINGS"]) or _find_col(df, ["GROSS PAY"])
    fit_taxable = _find_col(df, ["FEDERAL INCOME - EMPLOYEE TAXABLE"])
    fica_taxable = _find_col(df, ["SOCIAL SECURITY - EMPLOYEE TAXABLE"])
    medi_taxable = _find_col(df, ["MEDICARE - EMPLOYEE TAXABLE"])
    sit_taxable = _find_col(df, ["WORKED IN STATE - EMPLOYEE TAXABLE"])

    ded_rows = classify_deductions_pretax(
        df, cats["deductions"], total_earn_col,
        fit_taxable, fica_taxable, medi_taxable, sit_taxable,
    )
    contributions = [r for r in ded_rows if r.pop("_is_contribution", False)]
    deductions = [r for r in ded_rows if r not in contributions]
    for r in deductions:
        r.pop("_is_contribution", None)

    tax_rows = [{
        "Source Column": str(c).strip(),
        "Total $": round(float(df[c].apply(_num).sum()), 2),
        "Employees": int((df[c].apply(_num) != 0).sum()),
    } for c in cats["taxes"]]

    master_df = None
    if master_csv_file is not None:
        master_csv_file.seek(0)
        master_df = pd.read_csv(master_csv_file, dtype=str)
    tax_mapping_rows, states, missing = build_tax_mapping(df, cats["taxes"], master_df)

    bonus_info = classify_bonus(df, cats["earnings"])

    summary = [
        {"Metric": "Rows in file", "Value": len(df)},
        {"Metric": "Distinct earnings codes", "Value": len(earn_rows)},
        {"Metric": "Distinct contribution codes", "Value": len(contributions)},
        {"Metric": "Distinct deduction codes", "Value": len(deductions)},
        {"Metric": "Distinct tax columns", "Value": len(tax_rows)},
        {"Metric": "States detected", "Value": ", ".join(states) if states else "(none)"},
        {"Metric": "Tax mapping rows produced", "Value": len(tax_mapping_rows)},
        {"Metric": "Tax mapping rows missing from master", "Value": len(missing)},
        {"Metric": "Bonus classification verdict", "Value": bonus_info["verdict"]},
        {"Metric": "Bonus rows tested", "Value": bonus_info["rows_tested"]},
        {"Metric": "Bonus columns detected",
         "Value": ", ".join(bonus_info["bonus_columns_found"]) or "(none)"},
    ]

    bonus_rows = [{
        "Verdict": bonus_info["verdict"], "Reason": bonus_info["reason"],
        "Rows Tested": bonus_info["rows_tested"],
        "Discretionary Rows": bonus_info["discretionary_rows"],
        "Non-Discretionary Rows": bonus_info["non_discretionary_rows"],
        "Bonus Columns": ", ".join(bonus_info["bonus_columns_found"]),
    }]

    results = {
        "Summary": summary,
        "Earnings_Codes": earn_rows,
        "Contributions": contributions,
        "Deductions": deductions,
        "Taxes_Discovered": tax_rows,
        "Tax_Mapping": tax_mapping_rows,
        "Tax_Mapping_Missing": missing,
        "States_Detected": [{"State": s} for s in states],
        "Bonus_Classification": bonus_rows,
        "Bonus_Sample_Rows": bonus_info["samples"],
    }
    csv_bytes = tax_mapping_to_csv_bytes(tax_mapping_rows)
    return results, csv_bytes


def _results_to_xlsx_bytes(results):
    """Three-tab simplified xlsx that answers exactly:
      Tab 1 - What to set up in Uzio (Earnings | Contributions | Deductions)
      Tab 2 - Bonus discretionary or non-discretionary (verdict + one example)
      Tab 3 - Each deduction: pre-tax or post-tax + plain-English why
    Nothing else. Tax mapping is offered as a separate CSV download.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#1F4E78", "font_color": "white",
            "border": 1, "align": "left", "valign": "vcenter",
        })
        wrap_fmt = wb.add_format({"valign": "top", "text_wrap": True})
        verdict_pre = wb.add_format({
            "bold": True, "bg_color": "#C6EFCE", "font_color": "#006100",
            "align": "center", "valign": "vcenter",
        })
        verdict_post = wb.add_format({
            "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006",
            "align": "center", "valign": "vcenter",
        })
        verdict_nondisc = wb.add_format({
            "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006",
            "align": "left", "valign": "vcenter", "font_size": 14,
        })
        verdict_disc = wb.add_format({
            "bold": True, "bg_color": "#C6EFCE", "font_color": "#006100",
            "align": "left", "valign": "vcenter", "font_size": 14,
        })

        # ---- Tab 1: What to Set Up ----
        earn_codes = [r["Code"] for r in results["Earnings_Codes"]]
        contrib_codes = [r["Code"] for r in results["Contributions"]]
        ded_codes = [r["Code"] for r in results["Deductions"]]
        max_n = max(len(earn_codes), len(contrib_codes), len(ded_codes), 1)
        rows1 = []
        for i in range(max_n):
            rows1.append({
                "Earnings": earn_codes[i] if i < len(earn_codes) else "",
                "Contributions": contrib_codes[i] if i < len(contrib_codes) else "",
                "Deductions": ded_codes[i] if i < len(ded_codes) else "",
            })
        df1 = pd.DataFrame(rows1)
        df1.to_excel(writer, sheet_name="1. What to Set Up", index=False)
        ws1 = writer.sheets["1. What to Set Up"]
        ws1.set_column("A:A", 32); ws1.set_column("B:B", 24); ws1.set_column("C:C", 32)
        for i, col in enumerate(df1.columns):
            ws1.write(0, i, col, header_fmt)
        ws1.set_row(0, 24)

        # ---- Tab 2: Bonus ----
        bonus = results["Bonus_Classification"][0]
        sample = _pick_bonus_example(results["Bonus_Sample_Rows"], bonus["Verdict"])
        verdict_label = bonus["Verdict"].upper().replace("_", "-")
        rows2 = [
            ("Verdict", verdict_label),
            ("Reason", bonus["Reason"]),
            ("Bonus codes detected in file", bonus["Bonus Columns"]),
            ("Rows that had both bonus AND overtime", bonus["Rows Tested"]),
            ("    of which discretionary", bonus["Discretionary Rows"]),
            ("    of which non-discretionary", bonus["Non-Discretionary Rows"]),
        ]
        if sample:
            rows2 += [
                ("", ""),
                ("---- Example row that proves the verdict ----", ""),
                ("Associate ID", sample["associate"]),
                ("Regular earnings", f"${sample['regular_earnings']:,}"),
                ("Regular hours", sample["regular_hours"]),
                ("Regular rate ($/hr)", f"${sample['regular_rate']}"),
                ("Expected overtime rate  (1.5 x regular)", f"${sample['expected_ot_rate_1.5x']}"),
                ("Actual overtime rate from this row", f"${sample['actual_ot_rate']}"),
                ("Difference (%)", f"{sample['diff_pct']}%"),
                ("Bonus paid in this row", f"${sample['bonus_amt']:,}"),
                ("", ""),
                ("Plain-English explanation",
                    "Actual OT rate is HIGHER than 1.5 x regular rate => the bonus was rolled into "
                    "the regular rate before computing OT => bonus is NON-DISCRETIONARY (FLSA rule)."
                    if bonus["Verdict"] == "non_discretionary" else
                    "Actual OT rate matches 1.5 x regular rate exactly => the bonus did NOT inflate "
                    "the regular rate basis => bonus is DISCRETIONARY."
                    if bonus["Verdict"] == "discretionary" else
                    bonus["Reason"]),
            ]
        df2 = pd.DataFrame(rows2, columns=["Field", "Value"])
        df2.to_excel(writer, sheet_name="2. Bonus Verdict", index=False)
        ws2 = writer.sheets["2. Bonus Verdict"]
        ws2.set_column("A:A", 44); ws2.set_column("B:B", 80, wrap_fmt)
        for i, col in enumerate(df2.columns):
            ws2.write(0, i, col, header_fmt)
        ws2.set_row(0, 24)
        # Highlight the verdict cell (row 1, col B)
        if bonus["Verdict"] == "non_discretionary":
            ws2.write(1, 1, verdict_label, verdict_nondisc)
        elif bonus["Verdict"] == "discretionary":
            ws2.write(1, 1, verdict_label, verdict_disc)
        ws2.set_row(1, 28)

        # ---- Tab 3: Deductions Pre/Post-Tax ----
        rows3 = []
        for r in results["Contributions"] + results["Deductions"]:
            rows3.append({
                "Code": r["Code"],
                "Verdict": "PRE-TAX" if r["Verdict"] == "pre_tax" else "POST-TAX",
                "Why": _deduction_reason(r),
            })
        if not rows3:
            rows3 = [{"Code": "(none)", "Verdict": "",
                      "Why": "No voluntary deductions or contributions found in this file."}]
        df3 = pd.DataFrame(rows3)
        df3.to_excel(writer, sheet_name="3. Pre-Tax vs Post-Tax", index=False)
        ws3 = writer.sheets["3. Pre-Tax vs Post-Tax"]
        ws3.set_column("A:A", 26); ws3.set_column("B:B", 14)
        ws3.set_column("C:C", 110, wrap_fmt)
        for i, col in enumerate(df3.columns):
            ws3.write(0, i, col, header_fmt)
        ws3.set_row(0, 24)
        # Color the verdict cells
        for ri, r in enumerate(rows3, start=1):
            if r["Verdict"] == "PRE-TAX":
                ws3.write(ri, 1, "PRE-TAX", verdict_pre)
            elif r["Verdict"] == "POST-TAX":
                ws3.write(ri, 1, "POST-TAX", verdict_post)
            ws3.set_row(ri, 30)

    return buf.getvalue()


def _pick_bonus_example(samples, verdict):
    """Pick the single most illustrative row for the chosen verdict."""
    if not samples:
        return None
    if verdict == "non_discretionary":
        candidates = [s for s in samples if s["verdict_row"] == "non_discretionary"]
        return max(candidates, key=lambda s: s["diff_pct"]) if candidates else samples[0]
    if verdict == "discretionary":
        candidates = [s for s in samples if s["verdict_row"] == "discretionary"]
        return min(candidates, key=lambda s: abs(s["diff_pct"])) if candidates else samples[0]
    return samples[0]


def _deduction_reason(row):
    """Plain-English reason for a deduction's pre/post-tax verdict."""
    verdict = row["Verdict"]
    flavor = row.get("Pre-Tax Flavor", "")
    sample = row.get("Sample", "")
    if verdict == "post_tax":
        return "No row in the file showed taxable wages being reduced by this deduction's amount, so it does NOT shrink the tax base."
    if flavor == "section_125":
        first = sample.split(";")[0].strip() if sample else ""
        return ("Reduces FIT, FICA, Medicare, and state-income taxable wages by the deduction amount — Section 125 cafeteria plan." +
                (f" Example row: {first}" if first else ""))
    if flavor == "401k_traditional":
        first = sample.split(";")[0].strip() if sample else ""
        return ("Reduces FIT and state-income taxable wages but NOT FICA/Medicare — traditional 401(k)/403(b) pattern." +
                (f" Example row: {first}" if first else ""))
    if flavor == "pretax_unknown":
        return "Reduces FIT taxable wages only (no FICA/Medicare reduction observed)."
    if flavor == "mixed_unusual":
        return f"Pre-tax for: {row.get('Pre-Tax Of', '')} (unusual mix — review)."
    return "Pre-tax (see sample column for the matching row)."


# ---------- Streamlit UI ----------

def render_ui():
    st.title("ADP - Prior Payroll Setup Helper")
    st.caption(
        "Three answers from one ADP file: what to set up in Uzio, "
        "is the bonus discretionary, and which deductions are pre-tax vs post-tax."
    )

    col1, col2 = st.columns(2)
    with col1:
        adp_file = st.file_uploader(
            "ADP Prior Payroll File (sanitized)",
            type=["xlsx", "xls", "csv"],
            key="pps_helper_adp",
        )
    with col2:
        master_file = st.file_uploader(
            "State Tax Code Master CSV (optional, for tax mapping)",
            type=["csv"],
            key="pps_helper_master",
        )

    if not adp_file:
        st.info("Upload an ADP Prior Payroll file to begin.")
        return

    if not st.button("Run", type="primary"):
        return

    with st.spinner("Analyzing..."):
        try:
            results, csv_bytes = run_setup_helper(adp_file, master_file)
        except Exception as e:
            st.error(f"Failed to run analysis: {e}")
            raise

    # ------------------------------------------------------------------
    # ANSWER 1 — What to set up
    # ------------------------------------------------------------------
    st.markdown("## 1. What to set up in Uzio")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.markdown("**Earnings**")
        for r in results["Earnings_Codes"]:
            st.markdown(f"- {r['Code']}")
        if not results["Earnings_Codes"]:
            st.caption("(none)")
    with c2:
        st.markdown("**Contributions**")
        for r in results["Contributions"]:
            st.markdown(f"- {r['Code']}")
        if not results["Contributions"]:
            st.caption("(none)")
    with c3:
        st.markdown("**Deductions**")
        for r in results["Deductions"]:
            st.markdown(f"- {r['Code']}")
        if not results["Deductions"]:
            st.caption("(none)")

    # ------------------------------------------------------------------
    # ANSWER 2 — Bonus discretionary or non-discretionary
    # ------------------------------------------------------------------
    st.markdown("## 2. Bonus: discretionary or non-discretionary?")
    bonus = results["Bonus_Classification"][0]
    verdict = bonus["Verdict"]
    sample = _pick_bonus_example(results["Bonus_Sample_Rows"], verdict)

    if verdict == "non_discretionary":
        st.error("**NON-DISCRETIONARY**")
    elif verdict == "discretionary":
        st.success("**DISCRETIONARY**")
    else:
        st.warning(f"**{verdict.upper()}** — {bonus['Reason']}")

    if sample:
        st.markdown(
            f"""
**Example: Employee `{sample['associate']}`**

- Regular earnings: **${sample['regular_earnings']:,}** over **{sample['regular_hours']} hrs** → regular rate = **${sample['regular_rate']}/hr**
- Expected overtime rate (1.5 × regular rate) = **${sample['expected_ot_rate_1.5x']}/hr**
- Actual overtime rate from the file = **${sample['actual_ot_rate']}/hr**
- Bonus paid in this row: **${sample['bonus_amt']:,}**
"""
        )
        if verdict == "discretionary":
            st.markdown(
                "→ Actual OT rate matches 1.5 × regular rate. The bonus did **not** "
                "inflate the regular rate basis, so it's **discretionary**."
            )
        elif verdict == "non_discretionary":
            st.markdown(
                f"→ Actual OT rate is **higher** than 1.5 × regular rate "
                f"(diff: {sample['diff_pct']}%). The bonus was rolled into the regular "
                f"rate before computing OT, so it's **non-discretionary** under FLSA."
            )

    # ------------------------------------------------------------------
    # ANSWER 3 — Which deductions are pre-tax vs post-tax
    # ------------------------------------------------------------------
    st.markdown("## 3. Pre-tax vs post-tax (per deduction)")
    if not results["Deductions"] and not results["Contributions"]:
        st.caption("No voluntary deductions or contributions found in this file.")
    else:
        rows = []
        for r in results["Contributions"] + results["Deductions"]:
            rows.append({
                "Code": r["Code"],
                "Verdict": "PRE-TAX" if r["Verdict"] == "pre_tax" else "POST-TAX",
                "Why": _deduction_reason(r),
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    # ------------------------------------------------------------------
    # Downloads (full report stays available, just out of the way)
    # ------------------------------------------------------------------
    st.markdown("---")
    base = (adp_file.name or "ADP_Prior_Payroll").rsplit(".", 1)[0]
    dc1, dc2 = st.columns(2)
    with dc1:
        st.download_button(
            "Download Tax Mapping CSV",
            data=csv_bytes,
            file_name=f"{base}_Tax_Mapping.csv",
            mime="text/csv",
            disabled=not master_file,
            help=("Upload the State Tax Code master CSV to enable this download."
                  if not master_file else None),
        )
    with dc2:
        xlsx_bytes = _results_to_xlsx_bytes(results)
        st.download_button(
            "Download Full Detailed Report (xlsx)",
            data=xlsx_bytes,
            file_name=f"{base}_Setup_Helper.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Full multi-sheet workbook with $ totals, hours, sample rows, etc.",
        )
