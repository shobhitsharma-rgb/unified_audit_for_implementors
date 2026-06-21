import io
import re
import math
import pandas as pd
import streamlit as st

APP_TITLE = "ADP Payment Method Sanity Check"

# =========================================================
# ADP Payment Method Sanity Check
# - Input: single ADP payment method export (.xlsx / .csv)
# - Validates Uzio-compatible distribution rules:
#   R1. Multiple rows per EE = multiple accounts.
#   R2. Percent distribution: exactly one Full (% may be blank) + rest
#       Partial %. Sum of percents must equal 100%.
#   R3. Amount distribution: exactly one Full + any number of Partial
#       (amount) accounts.
#   R4. Mixed percent + amount is NOT supported by Uzio. Auto-fix:
#       keep percent rows as-is, split remaining % equally across the
#       non-percent accounts.
#   R5. Single row with Partial / Partial % is invalid -> fix to Full
#       with blank Amount and Percent.
# =========================================================

DEPOSIT_FULL = "Full"
DEPOSIT_PARTIAL_PCT = "Partial %"
DEPOSIT_PARTIAL_AMT = "Partial"


def _norm_text(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    return str(x).strip()


def _norm_id(x) -> str:
    s = _norm_text(x)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _norm_money(x):
    """Parse to float; return None when blank/unparseable."""
    if x is None:
        return None
    if isinstance(x, (int, float)) and not isinstance(x, bool):
        if pd.isna(x):
            return None
        return float(x)
    s = str(x).strip().replace(",", "").replace("$", "").replace("%", "")
    if not s:
        return None
    try:
        return float(s)
    except ValueError:
        return None


def _classify_deposit_type(raw_type: str) -> str:
    """Map ADP Deposit Type text to one of Full / Partial % / Partial / Unknown."""
    s = _norm_text(raw_type).lower()
    if not s:
        return "Unknown"
    if "balance" in s or s == "full" or "full" in s:
        return DEPOSIT_FULL
    if "partial %" in s or "partial%" in s or "percent" in s or "%" in s:
        return DEPOSIT_PARTIAL_PCT
    if "partial" in s or "amount" in s or "flat" in s:
        return DEPOSIT_PARTIAL_AMT
    return "Unknown"


def _find_col(df: pd.DataFrame, candidates) -> str:
    """Return the first matching column from `candidates` (case-insensitive, partial match)."""
    cols_upper = {c.upper(): c for c in df.columns}
    for cand in candidates:
        if cand.upper() in cols_upper:
            return cols_upper[cand.upper()]
    for cand in candidates:
        for col in df.columns:
            if cand.upper() in col.upper():
                return col
    return ""


def _read_adp_file(adp_file) -> pd.DataFrame:
    name = (adp_file.name or "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(adp_file, dtype=str)
    else:
        df = pd.read_excel(adp_file, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _round2(v):
    return round(float(v) + 1e-9, 2)


# ---------------------------------------------------------
# Rule A — de-duplicate identical accounts (same routing + account).
# ADP exports the SAME bank account multiple times across effective dates
# (account history), which the old logic mistook for multiple accounts and
# demoted to "Partial % 0.00" (rejected by the API). We collapse same-account
# rows to one, keeping the latest effective date, BEFORE the distribution logic
# runs. This alone makes most multi-row employees valid again; genuine
# multi-account employees (truly different routing/account) are left for the
# existing distribution rules.
# ---------------------------------------------------------
def _norm_acct(v) -> str:
    """Normalize a routing/account value for duplicate comparison — preserve
    leading zeros, just trim whitespace and a trailing float '.0'."""
    s = _norm_text(v)
    if s.endswith(".0"):
        s = s[:-2]
    return s


def _parse_eff(s) -> tuple:
    """Parse an Effective Date (MM/DD/YYYY) to a sortable (Y, M, D) tuple.
    Blank / unparseable -> (0, 0, 0) so it sorts oldest."""
    parts = str(s or "").strip().split("/")
    if len(parts) != 3:
        return (0, 0, 0)
    nums = []
    for p in parts:
        digits = "".join(ch for ch in p if ch.isdigit())
        if not digits:
            return (0, 0, 0)
        nums.append(int(digits))
    mm, dd, yy = nums
    if yy < 100:
        yy += 2000
    return (yy, mm, dd)


def _dedup_accounts(rows: list):
    """Collapse rows that are the SAME bank account (same routing + account)
    into one, keeping the row with the latest Effective Date. Rows with a blank
    routing or account are never merged. Returns (kept_rows, dropped_rows),
    preserving original row order in `kept`."""
    best: dict = {}
    for r in rows:
        routing, account = _norm_acct(r.get("routing")), _norm_acct(r.get("account"))
        if not routing or not account:
            continue
        key = (routing, account)
        if key not in best or _parse_eff(r.get("eff")) > _parse_eff(best[key].get("eff")):
            best[key] = r
    kept, dropped = [], []
    for r in rows:
        routing, account = _norm_acct(r.get("routing")), _norm_acct(r.get("account"))
        if not routing or not account:
            kept.append(r)
        elif best[(routing, account)] is r:
            kept.append(r)
        else:
            dropped.append(r)
    return kept, dropped


def _fix_employee(rows: list, issues: list, emp_id: str, emp_name: str):
    """Mutate `rows` in place to a Uzio-compatible distribution and append issues.

    Each row is a dict carrying: idx, raw_type, type, percent (float|None),
    amount (float|None), fixed_type, fixed_percent, fixed_amount.
    """
    if not rows:
        return

    n = len(rows)

    # ----- Rule 5: single-row Partial / Partial % -----
    if n == 1:
        r = rows[0]
        if r["type"] in (DEPOSIT_PARTIAL_PCT, DEPOSIT_PARTIAL_AMT):
            issues.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name,
                "Rule": "R5",
                "Severity": "Auto-Fix",
                "Issue": (
                    f"Single account but Deposit Type is '{r['raw_type']}'. "
                    "Uzio requires Full for a sole account."
                ),
                "Action": "Set Deposit Type=Full, clear Deposit Amount and Deposit Percent.",
            })
            r["fixed_type"] = DEPOSIT_FULL
            r["fixed_percent"] = None
            r["fixed_amount"] = None
        else:
            # Single Full row -> leave amount/percent blank (Uzio convention)
            r["fixed_type"] = DEPOSIT_FULL
            r["fixed_percent"] = None
            r["fixed_amount"] = None
        return

    # ----- Multi-row: classify the distribution mode -----
    full_rows = [r for r in rows if r["type"] == DEPOSIT_FULL]
    pct_rows = [r for r in rows if r["type"] == DEPOSIT_PARTIAL_PCT]
    amt_rows = [r for r in rows if r["type"] == DEPOSIT_PARTIAL_AMT]
    unknown_rows = [r for r in rows if r["type"] == "Unknown"]

    if unknown_rows:
        for r in unknown_rows:
            issues.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name,
                "Rule": "R0",
                "Severity": "Warning",
                "Issue": f"Unrecognized Deposit Type '{r['raw_type']}'.",
                "Action": "Review manually; treated as percent-based partial for fix.",
            })
            r["type"] = DEPOSIT_PARTIAL_PCT
            pct_rows.append(r)

    if len(full_rows) > 1:
        issues.append({
            "Employee ID": emp_id,
            "Employee Name": emp_name,
            "Rule": "R2/R3",
            "Severity": "Error",
            "Issue": f"{len(full_rows)} 'Full' rows found; Uzio allows only one.",
            "Action": "Keeping first Full as the balance account, converting the rest.",
        })
        # Demote extras to Partial % (will be handled below by mixed/percent logic)
        for extra in full_rows[1:]:
            extra["type"] = DEPOSIT_PARTIAL_PCT
            pct_rows.append(extra)
        full_rows = [full_rows[0]]

    has_pct = bool(pct_rows)
    has_amt = bool(amt_rows)

    # ----- Pure amount distribution (R3) -----
    if has_amt and not has_pct:
        if not full_rows:
            # Promote the last amount row to Full (consistent with R4 tiebreaker).
            issues.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name,
                "Rule": "R3",
                "Severity": "Error",
                "Issue": "Amount-based distribution has no 'Full' (remainder) account.",
                "Action": "Promoting last amount row to Full; its Amount cleared.",
            })
            promoted = amt_rows[-1]
            promoted["type"] = DEPOSIT_FULL
            full_rows = [promoted]
            amt_rows = [r for r in rows if r is not promoted]

        for r in rows:
            if r is full_rows[0]:
                r["fixed_type"] = DEPOSIT_FULL
                r["fixed_percent"] = None
                r["fixed_amount"] = None
            else:
                r["fixed_type"] = DEPOSIT_PARTIAL_AMT
                r["fixed_percent"] = None
                # Keep amount as-is; flag if missing
                amt = r["amount"]
                if amt is None or amt <= 0:
                    issues.append({
                        "Employee ID": emp_id,
                        "Employee Name": emp_name,
                        "Rule": "R3",
                        "Severity": "Error",
                        "Issue": "Partial (amount) row has blank/zero Deposit Amount.",
                        "Action": "Manual fix required: provide a Deposit Amount.",
                    })
                r["fixed_amount"] = amt
        return

    # ----- Pure percent distribution (R2) -----
    if has_pct and not has_amt:
        partial_sum = sum((r["percent"] or 0.0) for r in pct_rows)

        if not full_rows:
            # No Full -> the sum itself must equal 100; otherwise it's invalid.
            issues.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name,
                "Rule": "R2",
                "Severity": "Error",
                "Issue": "Percent-based distribution has no 'Full' account.",
                "Action": "Promoting the last Partial % row to Full as the remainder.",
            })
            promoted = pct_rows[-1]
            promoted["type"] = DEPOSIT_FULL
            full_rows = [promoted]
            pct_rows = [r for r in pct_rows if r is not promoted]
            partial_sum = sum((r["percent"] or 0.0) for r in pct_rows)

        full_row = full_rows[0]
        full_explicit = full_row["percent"]
        # If Full carries an explicit % too, include it in the check.
        total_if_full_explicit = (
            partial_sum + (full_explicit if full_explicit is not None else 0.0)
        )

        if full_explicit is not None and not math.isclose(total_if_full_explicit, 100.0, abs_tol=0.01):
            issues.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name,
                "Rule": "R2",
                "Severity": "Error",
                "Issue": (
                    f"Percent total = {total_if_full_explicit:.2f}% (Full row carries "
                    f"{full_explicit:.2f}%, Partial % sum = {partial_sum:.2f}%). "
                    "Total must equal 100%."
                ),
                "Action": (
                    f"Clearing Full %, leaving Partial % rows = {partial_sum:.2f}% "
                    "(remainder routes to Full)."
                ),
            })
        elif partial_sum > 100.0 + 0.01:
            issues.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name,
                "Rule": "R2",
                "Severity": "Error",
                "Issue": f"Partial % rows sum to {partial_sum:.2f}% (exceeds 100%).",
                "Action": "Manual fix required: reduce one or more Partial % values.",
            })
        elif partial_sum < 100.0 - 0.01:
            # Acceptable: Full absorbs the remainder. Informational only.
            pass

        for r in rows:
            if r is full_row:
                r["fixed_type"] = DEPOSIT_FULL
                r["fixed_percent"] = None
                r["fixed_amount"] = None
            else:
                r["fixed_type"] = DEPOSIT_PARTIAL_PCT
                r["fixed_percent"] = _round2(r["percent"] or 0.0)
                r["fixed_amount"] = None
        return

    # ----- Mixed percent + amount (R4) -----
    if has_pct and has_amt:
        kept_pct = sum((r["percent"] or 0.0) for r in pct_rows)
        non_pct_rows = amt_rows + full_rows  # rows to be re-percented
        remaining_pct = 100.0 - kept_pct

        if remaining_pct <= 0 or not non_pct_rows:
            issues.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name,
                "Rule": "R4",
                "Severity": "Error",
                "Issue": (
                    f"Mixed Percent + Amount; Partial % rows already total "
                    f"{kept_pct:.2f}%, no room to redistribute."
                ),
                "Action": "Manual fix required.",
            })
            # Best effort: keep percents, blank everything else as Partial %
            for r in rows:
                if r in pct_rows:
                    r["fixed_type"] = DEPOSIT_PARTIAL_PCT
                    r["fixed_percent"] = _round2(r["percent"] or 0.0)
                    r["fixed_amount"] = None
                else:
                    r["fixed_type"] = DEPOSIT_PARTIAL_PCT
                    r["fixed_percent"] = 0.0
                    r["fixed_amount"] = None
            return

        equal_share = _round2(remaining_pct / len(non_pct_rows))
        issues.append({
            "Employee ID": emp_id,
            "Employee Name": emp_name,
            "Rule": "R4",
            "Severity": "Auto-Fix",
            "Issue": (
                "Mixed Percent + Amount distribution (Uzio unsupported). "
                f"Percent rows kept ({kept_pct:.2f}% total); "
                f"remaining {remaining_pct:.2f}% split equally across "
                f"{len(non_pct_rows)} non-percent account(s)."
            ),
            "Action": (
                f"Each non-percent account set to {equal_share:.2f}%. "
                "Designated Full account (existing Full row, or last amount "
                "row when none) keeps Deposit Type = Full."
            ),
        })

        # Pick the Full row: prefer an existing Full; otherwise the last amount row.
        if full_rows:
            full_row = full_rows[0]
        else:
            full_row = amt_rows[-1]

        # Compensate rounding drift on the Full row so total hits 100.00 exactly.
        non_full_non_pct = [r for r in non_pct_rows if r is not full_row]
        running_total = kept_pct + equal_share * len(non_full_non_pct)
        full_share = _round2(100.0 - running_total)

        for r in rows:
            if r in pct_rows:
                r["fixed_type"] = DEPOSIT_PARTIAL_PCT
                r["fixed_percent"] = _round2(r["percent"] or 0.0)
                r["fixed_amount"] = None
            elif r is full_row:
                r["fixed_type"] = DEPOSIT_FULL
                # Full row's percent is informational; Uzio treats Full as remainder
                # so we leave it blank for cleanliness.
                r["fixed_percent"] = None
                r["fixed_amount"] = None
                # (full_share computed above is implicit — Full = 100 - others)
                _ = full_share
            else:
                r["fixed_type"] = DEPOSIT_PARTIAL_PCT
                r["fixed_percent"] = equal_share
                r["fixed_amount"] = None
        return

    # ----- Only Full rows (already handled if >1 above) -----
    if full_rows and not has_pct and not has_amt:
        if len(full_rows) == 1 and n == 1:
            return  # handled in single-row branch
        # Multiple rows all marked Full but no partials -> ambiguous
        issues.append({
            "Employee ID": emp_id,
            "Employee Name": emp_name,
            "Rule": "R2/R3",
            "Severity": "Error",
            "Issue": f"{n} rows all 'Full' with no Partial rows.",
            "Action": "Manual fix required: only one Full allowed.",
        })
        for i, r in enumerate(rows):
            r["fixed_type"] = DEPOSIT_FULL if i == 0 else DEPOSIT_PARTIAL_PCT
            r["fixed_percent"] = None
            r["fixed_amount"] = None
        return


def run_sanity(adp_file):
    df = _read_adp_file(adp_file)

    col_emp = _find_col(df, ["ASSOCIATE ID", "Employee ID", "Emp Code", "EmpID"])
    col_name = _find_col(df, ["NAME", "Employee Name", "Full Name"])
    col_routing = _find_col(df, ["ROUTING NUMBER", "Routing"])
    col_account = _find_col(df, ["ACCOUNT NUMBER", "Account"])
    col_dep_type = _find_col(df, ["DEPOSIT TYPE", "Deposit Type"])
    col_percent = _find_col(df, ["DEPOSIT PERCENT", "Percent"])
    col_amount = _find_col(df, ["DEPOSIT AMOUNT", "Amount"])
    col_eff = _find_col(df, ["EFFECTIVE DATE", "Effective"])

    required = {
        "Associate ID": col_emp,
        "Deposit Type": col_dep_type,
    }
    missing_required = [label for label, c in required.items() if not c]
    if missing_required:
        raise ValueError(
            "Could not find required column(s) in the ADP file: "
            + ", ".join(missing_required)
        )

    # Group rows per employee
    per_emp: dict[str, list[dict]] = {}
    emp_name_map: dict[str, str] = {}
    order: list[str] = []

    for idx, raw in df.iterrows():
        emp_id = _norm_id(raw.get(col_emp))
        if not emp_id:
            continue
        if emp_id not in per_emp:
            per_emp[emp_id] = []
            order.append(emp_id)
        if col_name:
            nm = _norm_text(raw.get(col_name))
            if nm and emp_id not in emp_name_map:
                emp_name_map[emp_id] = nm

        raw_type = _norm_text(raw.get(col_dep_type)) if col_dep_type else ""
        per_emp[emp_id].append({
            "idx": idx,
            "raw_type": raw_type,
            "type": _classify_deposit_type(raw_type),
            "routing": _norm_text(raw.get(col_routing)) if col_routing else "",
            "account": _norm_text(raw.get(col_account)) if col_account else "",
            "percent": _norm_money(raw.get(col_percent)) if col_percent else None,
            "amount": _norm_money(raw.get(col_amount)) if col_amount else None,
            "eff": _norm_text(raw.get(col_eff)) if col_eff else "",
            "fixed_type": None,
            "fixed_percent": None,
            "fixed_amount": None,
        })

    # Rule A — collapse duplicate same-account rows (keep latest effective date),
    # then analyze + fix each employee on the de-duplicated rows.
    issues: list[dict] = []
    dropped_idx: set = set()
    for emp_id in order:
        kept, dropped = _dedup_accounts(per_emp[emp_id])
        if dropped:
            for d in dropped:
                dropped_idx.add(d["idx"])
            issues.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name_map.get(emp_id, ""),
                "Rule": "R1",
                "Severity": "Auto-Fix",
                "Issue": (
                    f"{len(dropped)} duplicate account row(s): the same bank account "
                    "(identical Routing + Account Number) appears more than once, only "
                    "differing by effective date."
                ),
                "Action": "Kept the latest effective date and removed the older duplicate row(s).",
            })
            per_emp[emp_id] = kept
        _fix_employee(per_emp[emp_id], issues, emp_id, emp_name_map.get(emp_id, ""))

    # Build the corrected dataframe (preserve column order; rewrite the 3 fields)
    df_fixed = df.copy()
    if col_dep_type:
        for emp_id in order:
            for r in per_emp[emp_id]:
                if r["fixed_type"] is not None:
                    df_fixed.at[r["idx"], col_dep_type] = r["fixed_type"]
                if col_percent:
                    val = r["fixed_percent"]
                    df_fixed.at[r["idx"], col_percent] = "" if val is None else f"{val:.2f}"
                if col_amount:
                    val = r["fixed_amount"]
                    df_fixed.at[r["idx"], col_amount] = "" if val is None else f"{val:.2f}"

    # Drop the removed duplicate rows (Rule A) from the corrected output.
    if dropped_idx:
        df_fixed = df_fixed.drop(index=[i for i in dropped_idx if i in df_fixed.index])

    # Issues frame (always has the columns even when empty)
    issues_df = pd.DataFrame(
        issues,
        columns=["Employee ID", "Employee Name", "Rule", "Severity", "Issue", "Action"],
    )

    # Before / After preview per employee (only employees with issues or changes)
    preview_rows: list[dict] = []
    changed_ids = {i["Employee ID"] for i in issues}
    for emp_id in order:
        if emp_id not in changed_ids:
            continue
        for r in per_emp[emp_id]:
            preview_rows.append({
                "Employee ID": emp_id,
                "Employee Name": emp_name_map.get(emp_id, ""),
                "Routing": r["routing"],
                "Account": r["account"],
                "Deposit Type (Before)": r["raw_type"],
                "Deposit Percent (Before)": "" if r["percent"] is None else f"{r['percent']:.2f}",
                "Deposit Amount (Before)": "" if r["amount"] is None else f"{r['amount']:.2f}",
                "Deposit Type (After)": r["fixed_type"] or r["raw_type"],
                "Deposit Percent (After)": "" if r["fixed_percent"] is None else f"{r['fixed_percent']:.2f}",
                "Deposit Amount (After)": "" if r["fixed_amount"] is None else f"{r['fixed_amount']:.2f}",
            })
    preview_df = pd.DataFrame(
        preview_rows,
        columns=[
            "Employee ID", "Employee Name", "Routing", "Account",
            "Deposit Type (Before)", "Deposit Percent (Before)", "Deposit Amount (Before)",
            "Deposit Type (After)", "Deposit Percent (After)", "Deposit Amount (After)",
        ],
    )

    summary_df = pd.DataFrame({
        "Metric": [
            "Total rows",
            "Total employees",
            "Employees with issues",
            "Auto-Fix actions",
            "Errors requiring manual review",
        ],
        "Value": [
            len(df),
            len(order),
            issues_df["Employee ID"].nunique() if not issues_df.empty else 0,
            int((issues_df["Severity"] == "Auto-Fix").sum()) if not issues_df.empty else 0,
            int((issues_df["Severity"] == "Error").sum()) if not issues_df.empty else 0,
        ],
    })

    # Stringify everything to keep long account/routing numbers from being
    # rendered in exponential notation downstream (Excel display + pandas write).
    df_fixed_clean = df_fixed.fillna("").astype(str)
    df_fixed_clean = df_fixed_clean.replace({"nan": "", "NaN": "", "None": ""})

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        issues_df.to_excel(writer, sheet_name="Issues", index=False)
        preview_df.to_excel(writer, sheet_name="Before_After", index=False)
        df_fixed_clean.to_excel(writer, sheet_name="Corrected_Source", index=False)

    # CSV is intended for API ingestion: leave values plain and write as bare
    # UTF-8 (NO BOM). Downstream APIs match the first header literally, so a
    # utf-8-sig BOM would smuggle U+FEFF in front of e.g. "Associate ID" and the
    # column lookup silently misses. Excel users should open the XLSX export
    # instead — openpyxl writes long account / routing numbers as text cells
    # which Excel renders literally without scientific notation.
    csv_bytes = df_fixed_clean.to_csv(index=False).encode("utf-8")

    return out.getvalue(), csv_bytes, summary_df, issues_df, preview_df


def render_ui():
    st.title(APP_TITLE)
    st.markdown(
        """
**Purpose**: Validate an ADP payment-method export against Uzio's distribution rules
and auto-correct unsupported configurations.

**Rules enforced**:
1. Multiple rows per employee = multiple direct-deposit accounts.
2. Percent distribution must contain **exactly one Full** (% may be blank) and the
   rest **Partial %**. Sum of percents must equal **100%**.
3. Amount distribution must contain **exactly one Full**; the rest are Partial
   (amount) accounts.
4. Mixed Percent + Amount is **unsupported by Uzio**. Auto-fix keeps any
   Partial % rows as-is and splits the remaining percentage equally across the
   non-percent accounts.
5. A single row with `Partial` / `Partial %` is invalid — auto-fixed to `Full`
   with Deposit Amount and Deposit Percent cleared.
"""
    )

    client_name = st.text_input("Client Name", value="Client", key="adp_pm_sanity_client")
    adp_file = st.file_uploader(
        "Upload ADP Payment Method Export (.xlsx / .csv)",
        type=["xlsx", "xls", "csv"],
        key="adp_pm_sanity_upload",
    )

    if not adp_file:
        return

    if st.button("Run Sanity Check", type="primary", key="adp_pm_sanity_run"):
        try:
            with st.spinner("Analyzing payment methods..."):
                report_bytes, csv_bytes, summary_df, issues_df, preview_df = run_sanity(adp_file)
        except Exception as e:
            st.error(f"Failed: {e}")
            st.exception(e)
            return

        st.success("Sanity check complete.")

        st.subheader("Summary")
        st.dataframe(summary_df, hide_index=True, use_container_width=True)

        if issues_df.empty:
            st.info("No issues found — the file already conforms to Uzio rules.")
        else:
            st.subheader("Issues")
            st.dataframe(issues_df, hide_index=True, use_container_width=True)

            st.subheader("Before / After")
            st.dataframe(preview_df, hide_index=True, use_container_width=True)

        timestamp = pd.Timestamp.now().strftime("%d_%m_%Y_%H%M")
        xlsx_name = f"{client_name}_ADP_Payment_Method_Sanity_{timestamp}.xlsx"
        csv_name = f"{client_name}_ADP_Payment_Method_Corrected_{timestamp}.csv"

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="📊 Download Full Report (.xlsx)",
                data=report_bytes,
                file_name=xlsx_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                key="adp_pm_sanity_dl_xlsx",
            )
        with col2:
            st.download_button(
                label="📄 Download Corrected Source (.csv)",
                data=csv_bytes,
                file_name=csv_name,
                mime="text/csv",
                key="adp_pm_sanity_dl_csv",
            )


if __name__ == "__main__":
    st.set_page_config(page_title=APP_TITLE, layout="centered", initial_sidebar_state="collapsed")
    render_ui()
