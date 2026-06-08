"""ADP Consolidated Audit.

Runs the four standalone ADP audits — Census, Direct Deposit (Payment), Emergency
Contact, and License — in a single pass against ONE Uzio HR Report + three ADP
exports, and produces one consolidated workbook.

Architecture: the standalone tools own all audit logic, including every anomaly
check (FLSA Compliance, Salaried Driver Exception, mixed-mode R4 payment, dup SSN,
active/terminated missing in Uzio, hourly-rate anomalies, hourly-zero-hours, ...).
This module just reshapes the Uzio HR Report into the file shape each standalone
expects, calls compute_audit_dataframes() on each, and stitches the resulting
sheets into the consolidated workbook. Standalone tools stay completely untouched.

UI: 1 Uzio uploader (HR Report — same comprehensive CSV the Paycom consolidator
uses) + 3 ADP uploaders (Census, Direct Deposit, combined Emergency+License).
ADP file #3 drives both the Emergency audit and the License audit.
"""

import importlib
import io
from datetime import datetime

import pandas as pd
import streamlit as st

# Reuse the Uzio HR Report reader from the Paycom consolidator — exact same file
# shape, no point re-implementing.
from apps.common.paycom_combined_audit import (
    read_uzio_master,
    company_name_from_uzio_master,
)

APP_TITLE = "ADP - Consolidated Audit"

# ---------------------------------------------------------------------------
# Column maps: Uzio HR Report "Section|Field" → the column name the matching
# standalone ADP audit's reader expects to find. Each map is applied as a simple
# DataFrame rename + projection; the standalone then runs its normal pipeline.
# ---------------------------------------------------------------------------

# Census standalone uses utils.audit_utils.UZIO_RAW_MAPPING keys (with * / **
# suffixes that the Uzio Census Template actually carries). Match those exactly
# so the standalone's normalization recognizes each field.
_CENSUS_COL_MAP = {
    "Job|Employee ID": "Employee ID*",
    "Personal|First Name": "Employee First Name*",
    "Personal|Last Name": "Employee Last Name*",
    "Personal|Middle Name": "Employee Middle Initial",
    "Personal|Suffix": "Employee Suffix",
    "Job|Status": "Employment Status*",
    "Job|Date of Hire": "Date of Hire*",
    "Job|Original DOH": "Original DOH",
    "Job|Termination Date": "Termination Date",
    "Job|Termination Reason": "Termination Reason",
    "Job|Employment Type": "Employment Type*",
    "Job|Pay Type": "Pay Type*",
    "Job|Annual Salary": "Annual Salary(Digits)**",
    "Job|Hourly Rate": "Hourly Pay Rate**",
    "Job|Working Hours per Week": "Working Hours per Week(Digits)**",
    "Job|Job Title": "Job Title",
    "Job|Department": "Department",
    "Personal|Work Email": "Official Email*",
    "Home Address|Personal Email": "Personal Email",
    "Home Address|Phone": "Phone Number(Digits)",
    "Personal|SSN": "Employee SSN",
    "Personal|Date Of Birth": "Employee Date of Birth*",
    "Personal|Gender": "Employee Gender*",
    "Personal|Tobacco Usage": "Employee Tobacco usage in last 12 months",
    "Job|FLSA Classification": "FLSA Classification",
    "Home Address|Address Line 1": "Employee Address Line 1",
    "Home Address|Address Line 2": "Employee Address Line 2",
    "Home Address|City": "City*",
    "Home Address|Zip": "Zipcode*",
    "Home Address|State": "State(Abbreviation)*",
    "Mailing Address|Address Line 1": "Mailing Address Line 1",
    "Mailing Address|Address Line 2": "Mailing Address Line 2",
    "Mailing Address|City": "Mailing City",
    "Mailing Address|Zip": "Mailing Zipcode",
    "Mailing Address|State": "Mailing State(Abbreviation)",
    "Job|Reporting Manager": "Reporting Manager ID",
    "Job|Work Location": "Work Location",
    "Additional Information|License Number": "License Number*",
    "Additional Information|License Expiration Date": "License Expiration Date",
}

# Payment standalone uses its own get_col() lookup; keys here match its hits.
_PAYMENT_COL_MAP = {
    "Job|Employee ID": "Employee ID",
    "Payment Method|Routing Number": "Routing Number",
    "Payment Method|Account Number": "Account Number",
    "Payment Method|Account Type": "Account Type",
    "Payment Method|Paycheck Percentage": "Paycheck Percentage",
    "Payment Method|Paycheck Amount": "Paycheck Amount",
}

# Emergency standalone looks for these literal column names (with substring match).
_EMERGENCY_COL_MAP = {
    "Job|Employee ID": "Employee ID",
    "Emergency Contact|Name": "Name",
    "Emergency Contact|Relationship": "Relationship",
    "Emergency Contact|Phone": "Phone",
}

# License standalone scans for 'Employee ID' header in first 20 rows.
_LICENSE_COL_MAP = {
    "Job|Employee ID": "Employee ID",
    "Additional Information|License Number": "License Number",
    "Additional Information|License Expiration Date": "License Expiration Date",
}


def _full_name_series(df_master: pd.DataFrame) -> pd.Series:
    """Helper — Uzio HR Report stores name in Personal|First Name + Personal|Last Name."""
    fn = df_master.get("Personal|First Name", pd.Series([""] * len(df_master))).fillna("").astype(str)
    ln = df_master.get("Personal|Last Name", pd.Series([""] * len(df_master))).fillna("").astype(str)
    return (fn + " " + ln).str.strip()


def _project_and_rename(df_master: pd.DataFrame, col_map: dict) -> pd.DataFrame:
    """Project the HR Report DataFrame down to the columns in col_map and rename
    them to the standalone's expected column names. Columns missing in df_master
    are filled with empty strings so downstream readers don't KeyError."""
    out = pd.DataFrame()
    for src, dst in col_map.items():
        if src in df_master.columns:
            out[dst] = df_master[src]
        else:
            out[dst] = ""
    return out


def _adapt_census(df_master: pd.DataFrame) -> io.BytesIO:
    """Build an .xlsx with sheet 'Employee Details' and header on row 4 — the shape
    utils.audit_utils.read_uzio_raw_file expects. Dedup to one row per employee
    (census is per-employee data; HR Report multi-rows carry banking/emergency
    duplication which we don't want here)."""
    df = _project_and_rename(df_master, _CENSUS_COL_MAP)
    df = df.drop_duplicates(subset=["Employee ID*"], keep="first").reset_index(drop=True)
    df["Full Name"] = _full_name_series(df_master.drop_duplicates(
        subset=["Job|Employee ID"], keep="first").reset_index(drop=True))
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Employee Details", index=False, startrow=3, header=True)
    buf.seek(0)
    return buf


def _adapt_payment(df_master: pd.DataFrame) -> io.BytesIO:
    """Build an .xlsx with header on row 2 — what the standalone payment audit's
    _read_payment_file(header=1) expects. Preserve multi-row so every bank account
    per employee gets compared. Drop rows with no banking data."""
    df = _project_and_rename(df_master, _PAYMENT_COL_MAP)
    df["Full Name"] = _full_name_series(df_master)
    routing = df["Routing Number"].fillna("").astype(str).str.strip().ne("")
    account = df["Account Number"].fillna("").astype(str).str.strip().ne("")
    df = df[routing | account].reset_index(drop=True)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=1, header=True)
    buf.seek(0)
    return buf


def _adapt_emergency(df_master: pd.DataFrame) -> io.BytesIO:
    """Build an .xlsx with header on row 2 — what emergency_audit.compute_audit_dataframes
    expects via pd.read_excel(header=1). Preserve multi-row so every emergency
    contact per employee gets compared. Drop rows with no contact data."""
    df = _project_and_rename(df_master, _EMERGENCY_COL_MAP)
    name_filled = df["Name"].fillna("").astype(str).str.strip().ne("")
    phone_filled = df["Phone"].fillna("").astype(str).str.strip().ne("")
    df = df[name_filled | phone_filled].reset_index(drop=True)
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Sheet1", index=False, startrow=1, header=True)
    buf.seek(0)
    return buf


def _adapt_license(df_master: pd.DataFrame) -> pd.DataFrame:
    """License standalone takes a DataFrame directly (via read_uzio_license, but we
    can skip that and feed the DataFrame straight to run_license_audit). Preserve
    multi-row — same employee could have multiple licenses. Drop rows with no
    license number."""
    df = _project_and_rename(df_master, _LICENSE_COL_MAP)
    df["Full Name"] = _full_name_series(df_master)
    has_lic = df["License Number"].fillna("").astype(str).str.strip().ne("")
    return df[has_lic].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Per-audit runners — each one wraps the standalone tool's compute entry point
# in a try/except so a failure in one audit doesn't kill the others.
# ---------------------------------------------------------------------------

def _safe(callable_, *args, **kwargs):
    try:
        return callable_(*args, **kwargs), None
    except Exception as exc:
        return None, f"{type(exc).__name__}: {exc}"


def _run_census(df_master, adp_file):
    import apps.adp.census_audit as census_audit
    importlib.reload(census_audit)
    uzio_buf = _adapt_census(df_master)
    return census_audit.compute_audit_dataframes(uzio_buf, adp_file)


def _run_payment(df_master, adp_file):
    import apps.adp.payment_audit as payment_audit
    importlib.reload(payment_audit)
    uzio_buf = _adapt_payment(df_master)
    return payment_audit.compute_audit_dataframes(uzio_buf, adp_file)


def _run_emergency(df_master, adp_file):
    import apps.adp.emergency_audit as emergency_audit
    importlib.reload(emergency_audit)
    uzio_buf = _adapt_emergency(df_master)
    if hasattr(adp_file, "seek"):
        adp_file.seek(0)
    return emergency_audit.compute_audit_dataframes(uzio_buf, adp_file)


def _run_license(df_master, adp_file):
    import apps.adp.license_audit as license_audit
    importlib.reload(license_audit)
    uzio_df = _adapt_license(df_master)
    if hasattr(adp_file, "seek"):
        adp_file.seek(0)
    adp_df = license_audit.read_adp_license(adp_file)
    if adp_df is None:
        raise ValueError("Could not parse ADP license file (header row not found).")
    return license_audit.run_license_audit(uzio_df, adp_df)


# ---------------------------------------------------------------------------
# Workbook stitching
# ---------------------------------------------------------------------------

def _write_workbook(census_dfs, payment_dfs, emergency_dfs,
                    license_df) -> bytes:
    # Census summary/roll-up sheets are intentionally excluded from the
    # consolidated workbook — only the per-field detail and anomaly sheets ship.
    _CENSUS_SKIP = {"Summary", "Field_Summary_By_Status"}
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as writer:
        for name, df in (census_dfs or {}).items():
            if name in _CENSUS_SKIP:
                continue
            if df is None or df.empty:
                continue
            sheet = f"CEN_{name}"[:31]
            df.to_excel(writer, sheet_name=sheet, index=False)

        df_cmp = (payment_dfs or {}).get("Comparison_Detail", pd.DataFrame())
        df_exc = (payment_dfs or {}).get("Exception_Mixed_Mode", pd.DataFrame())
        if not df_cmp.empty:
            df_cmp.to_excel(writer, sheet_name="DD_Comparison_Detail", index=False)
        if not df_exc.empty:
            df_exc.to_excel(writer, sheet_name="DD_Exception_Mixed_Mode", index=False)
            workbook = writer.book
            exc_sheet = writer.sheets["DD_Exception_Mixed_Mode"]
            green_fmt = workbook.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
            red_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
            last_row = max(len(df_exc), 1) + 1
            exc_sheet.conditional_format(f"G2:G{last_row}", {
                "type": "text", "criteria": "containing",
                "value": "Corrected Setup", "format": green_fmt,
            })
            exc_sheet.conditional_format(f"G2:G{last_row}", {
                "type": "text", "criteria": "containing",
                "value": "Mismatch (Mixed Mode)", "format": red_fmt,
            })

        df_em = (emergency_dfs or {}).get("Emergency_Contact_Audit", pd.DataFrame())
        if not df_em.empty:
            df_em.to_excel(writer, sheet_name="EC_Emergency_Contact_Audit", index=False)

        if license_df is not None and not license_df.empty:
            license_df.to_excel(writer, sheet_name="LIC_License_Audit", index=False)

    return out.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

def render_ui():
    st.title(APP_TITLE)
    st.caption(
        "Runs the ADP Census, Direct Deposit, Emergency Contact, and License audits "
        "in a single pass against one Uzio HR Report and up to three ADP exports."
    )

    with st.expander("How this tool works"):
        st.markdown(
            """
            **What you upload**
            - **Uzio HR Report** — required. One comprehensive CSV that covers
              census, banking, emergency contacts, and license info for every
              employee. Same file the Paycom Consolidated Audit uses.
            - **ADP files (any combination)** — at least one is required:
              - *ADP Census Export* → drives the **Census audit**
              - *ADP Direct Deposit Export* → drives the **Direct Deposit audit**
              - *ADP Emergency + License Details Report* → drives **both** the
                Emergency Contact audit and the License audit

            **Partial runs are supported.** If you only have (e.g.) Census + DD
            ready, run the tool with just those two — the missing audits simply
            contribute no sheets to the workbook.

            **What you get**
            One consolidated workbook with one tab per sheet each standalone
            audit would have produced (census comparison + every anomaly check,
            payment mixed-mode R4 with highlighting, emergency contact
            comparison, license comparison).

            **Multi-row data is fully audited** — every bank account, every
            emergency contact, every license. The Uzio HR Report's multi-row
            representation is preserved through to each audit; census is
            deduplicated to one row per employee (correct, since census data
            is per-employee).
            """
        )


    st.markdown("### Uzio")
    uzio_file = st.file_uploader(
        "Uzio HR Report (.csv) — comprehensive export with Personal / Job / "
        "Payment Method / Emergency Contact / Additional Information sections",
        type=["csv"], key="adp_cons_uzio_hr",
        help=(
            "Download from Uzio → Reports → HR Reports → run the Master HR "
            "Report and export as CSV. Headers are on row 2 with section "
            "categories on row 1 (e.g. 'Job | Employee ID'). Required."
        ),
    )

    st.markdown("### ADP")
    a1, a2, a3 = st.columns(3)
    with a1:
        adp_census = st.file_uploader(
            "ADP Census Export (.xlsx/.csv)",
            type=["xlsx", "csv"], key="adp_cons_adp_cen",
            help=(
                "ADP Workforce Now → Reports → Custom Reports → Employee "
                "Census / Personnel Roster (the export that includes "
                "Associate ID, Legal First/Last Name, FLSA Description, "
                "Position Status, Hire Date, etc.). Drives the Census audit "
                "(field-by-field comparison + FLSA Compliance, Salaried "
                "Driver Exceptions, Active/Term Missing in Uzio, Duplicate "
                "SSN, Hourly Rate Anomalies, Hourly Zero Hours, Data "
                "Quality). Optional — leave blank to skip the Census audit."
            ),
        )
    with a2:
        adp_dd = st.file_uploader(
            "ADP Direct Deposit Export (.xlsx/.csv)",
            type=["xlsx", "csv"], key="adp_cons_adp_dd",
            help=(
                "ADP Workforce Now → Reports → Direct Deposit Information "
                "(or your equivalent banking export). One row per bank "
                "account (an employee with multiple accounts shows multiple "
                "rows). Drives the Direct Deposit audit including the "
                "mixed-mode 'Partial $ + Partial %' R4 verdict on the "
                "Exception sheet. Optional."
            ),
        )
    with a3:
        adp_em = st.file_uploader(
            "ADP Emergency + License Details Report (.xlsx/.csv) — drives both audits",
            type=["xlsx", "csv"], key="adp_cons_adp_em",
            help=(
                "ADP Workforce Now → Reports → Employee License Details + "
                "Emergency Contact (a single export that carries both "
                "Contact Name / Relationship / Mobile Phone columns AND "
                "License/Certification Code / Expiration Date columns). "
                "Drives BOTH the Emergency Contact audit and the License "
                "audit. Optional — without this file, both audits are "
                "skipped."
            ),
        )

    if st.button("Run Consolidated Audit", type="primary"):
        # Uzio HR Report is required (every audit needs the Uzio side). Each ADP
        # file is OPTIONAL — the matching audit runs only if its file is uploaded,
        # otherwise it simply contributes no sheets to the workbook.
        if not uzio_file:
            st.error("Please upload the Uzio HR Report — it is required for any audit to run.")
            return
        if not (adp_census or adp_dd or adp_em):
            st.error("Please upload at least one ADP file (Census, Direct Deposit, or Emergency + License).")
            return

        # Parse Uzio HR Report once.
        try:
            uzio_file.seek(0)
            df_master = read_uzio_master(uzio_file)
        except Exception as exc:
            st.error(f"Failed to parse Uzio HR Report: {exc}")
            return

        # Client name comes straight from the Uzio report's
        # Company Information > Company Name column.
        client_name = company_name_from_uzio_master(df_master) or "Client"

        SKIPPED = "SKIPPED"
        errs = {}
        census_dfs = payment_dfs = emergency_dfs = None
        license_df = None

        if adp_census:
            with st.spinner("Running Census audit..."):
                census_dfs, errs["census"] = _safe(_run_census, df_master, adp_census)
        else:
            errs["census"] = SKIPPED

        if adp_dd:
            with st.spinner("Running Direct Deposit (Payment) audit..."):
                payment_dfs, errs["payment"] = _safe(_run_payment, df_master, adp_dd)
        else:
            errs["payment"] = SKIPPED

        if adp_em:
            with st.spinner("Running Emergency Contact audit..."):
                emergency_dfs, errs["emergency"] = _safe(_run_emergency, df_master, adp_em)
            with st.spinner("Running License audit..."):
                license_df, errs["license"] = _safe(_run_license, df_master, adp_em)
        else:
            errs["emergency"] = SKIPPED
            errs["license"] = SKIPPED

        # Surface only genuine audit failures (a skipped audit — no file
        # uploaded — stays silent; the user asked for a clean, button-only UI).
        for label, key in [("Census", "census"), ("Direct Deposit", "payment"),
                            ("Emergency Contact", "emergency"), ("License", "license")]:
            err = errs.get(key)
            if err and err != SKIPPED:
                st.error(f"{label} audit failed: {err}")

        report_bytes = _write_workbook(
            census_dfs or {}, payment_dfs or {},
            emergency_dfs or {}, license_df,
        )
        ts = datetime.now().strftime("%d_%m_%Y_%H%M")
        display_client = client_name.strip() or "Client"
        safe_client = "".join(ch for ch in client_name
                              if ch.isalnum() or ch in ("_", "-")) or "Client"
        st.success("Report generated.")
        st.download_button(
            label=f"Download {display_client} Consolidated Audit Report",
            data=report_bytes,
            file_name=f"ADP_{safe_client}_Consolidated_Audit_Report_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )
