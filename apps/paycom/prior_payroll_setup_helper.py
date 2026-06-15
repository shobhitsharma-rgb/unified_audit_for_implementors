"""Paycom - Prior Payroll Setup Helper Tool.

Replaces the deprecated 'Paycom - Deduction Analyzer'. Mirrors the ADP
'Prior Payroll Setup Helper' shape: three answers, one Excel output with
exactly three tabs.

Inputs (both Paycom files required):
  1. Paycom Prior Payroll Register (long format with Type Code / Type
     Description / Amount / Code Description columns).
  2. Paycom Scheduled Deductions Report (with Deduction Code / Deduction
     Desc / Tax Treatment columns).

Outputs:
  Tab 1  What to Set Up (Earnings | Contributions | Deductions, codes only)
  Tab 2  Pre-tax vs Post Tax (read straight from Tax Treatment column;
         no algorithm needed -- Paycom labels each deduction directly)
  Tab 3  Bonus Verdict (FLSA: Strategy A+C using Paycom's WOT vs plain
         OT differential when present, otherwise indeterminate)
"""

from __future__ import annotations
import base64
import difflib
import io
import json
import os
import re
import zipfile

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# ---------- Pure-Python analysis (mirrors core/paycom/prior_payroll_setup_helper.py) ----------

def _num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)) and not pd.isna(v):
        return float(v)
    s = str(v).strip().replace(",", "").replace("$", "")
    if s in ("", "-", "nan", "NaT", "None"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def _read_either(file) -> pd.DataFrame:
    """Streamlit UploadedFile -> DataFrame."""
    file.seek(0)
    name = (file.name or "").lower()
    if name.endswith(".csv"):
        return pd.read_csv(file)
    return pd.read_excel(file)


CONTRIB_PATTERN = re.compile(r"\b(401[Kk]?|403[Bb]?|457|ROTH|HSA|FSA|RETIREMENT)\b")
BONUS_RE = re.compile(r"\b(BONUS|BNS|BND|BNH|BN[0-9]?|NA[0-9])\b", re.IGNORECASE)

# Deductions UZIO creates automatically when a client is set up.
# The Tampermonkey automation script must NOT try to re-create these;
# we filter them out of the extracted deductions list before display + export.
# Match is case-insensitive on the (already-stripped) Type Description.
# To add another default: append the lowercased description string here.
DEFAULT_UZIO_DEDUCTIONS_TO_SKIP = {
    "earned wage access",
}

# New Prior-Payroll bifurcation rule (replaces the lexical-prefix CONTRIB_PATTERN
# for the deductions extracted from the Prior Payroll file):
#   - If Type Code or Type Description contains the word "Match" or "Memo"
#     anywhere (case-insensitive, word-boundary), the row represents an
#     EMPLOYER-side entry (matching contribution, or a memo/informational
#     line tracked for reporting). Route to Contributions.
#   - Everything else is an employee-paid Deduction.
# Rationale: an employee's 401K Roth deferral IS a deduction from their gross
# pay; only the employer's matching contribution is a "Contribution". The old
# regex bucketed both rows as Contributions because both contain "401K", which
# is wrong.
MATCH_MEMO_RE = re.compile(r"\b(match|memo)\b", re.IGNORECASE)


def _is_ignored_paycom_item(type_code, type_description):
    """Items the tool drops EVERYWHERE — earnings, deductions, contributions, and
    taxes. Currently: Worker's Compensation (code WKC, or a name containing
    'worker(s) comp...'). UZIO handles WC separately, so we never emit it."""
    code = (type_code or "").strip().upper()
    desc = (type_description or "").strip().lower()
    if code == "WKC":
        return True
    if "worker" in desc and "comp" in desc:
        return True
    return False


def _extract_unique_pairs_by_code_description(prior_df, code_description_value):
    """Generic helper: filter Prior Payroll file(s) to rows where the
    `Code Description` column equals a specific value, then return unique
    (Type Code, Type Description) pairs sorted by code.

    Used by both extract_unique_deductions_from_prior and
    extract_unique_earnings_from_prior so the filter logic stays in one place.
    """
    required = {"Code Description", "Type Code", "Type Description"}
    missing = required - set(prior_df.columns)
    if missing:
        return []

    mask = prior_df["Code Description"].astype(str).str.strip() == code_description_value
    sub = prior_df[mask]
    if sub.empty:
        return []

    pairs = sub[["Type Code", "Type Description"]].copy()
    pairs["Type Code"] = pairs["Type Code"].astype(str).str.strip()
    pairs["Type Description"] = pairs["Type Description"].astype(str).str.strip()
    # Treat literal "nan" / "None" from string-coerced NaNs as empty.
    pairs = pairs[~pairs["Type Code"].str.lower().isin(["", "nan", "none"])]
    unique = pairs.drop_duplicates().sort_values(["Type Code", "Type Description"]).reset_index(drop=True)

    return [{"Type Code": r["Type Code"], "Type Description": r["Type Description"]}
            for _, r in unique.iterrows()
            if not _is_ignored_paycom_item(r["Type Code"], r["Type Description"])]


def extract_unique_deductions_from_prior(prior_df):
    """Extract unique (Type Code, Type Description) pairs from the concatenated
    Prior Payroll file(s) where `Code Description == "Deductions"`.

    Source of truth is the `Code Description` column in the Paycom Prior Payroll
    file — Paycom labels every line as one of: Earnings / W/H Taxes / Deductions /
    Net Pay Distribution / Employee Benefits / Client Side Liabilities. We filter
    to "Deductions" and dedupe on the (Type Code, Type Description) pair, because
    one Type Code can appear with multiple descriptions (e.g. R4P with both
    "Roth 401K %" and "Roth 401K % Match").

    Returns a list of {"Type Code", "Type Description"} dicts, sorted by code.
    """
    return _extract_unique_pairs_by_code_description(prior_df, "Deductions")


def extract_unique_earnings_from_prior(prior_df):
    """Extract unique (Type Code, Type Description) pairs from the concatenated
    Prior Payroll file(s) where `Code Description == "Earnings"`.

    Same shape as extract_unique_deductions_from_prior; different filter value.

    Returns a list of {"Type Code", "Type Description"} dicts, sorted by code.
    """
    return _extract_unique_pairs_by_code_description(prior_df, "Earnings")


def _normalize_calc_desc(s):
    """Lowercase + strip for case/whitespace-insensitive Calc Description match."""
    if not isinstance(s, str):
        return ""
    return s.strip().lower()


# Map a Paycom Calc Description (normalized: lowercased + stripped) to the
# UZIO tax-treatment label the implementor needs. Extend as new values are
# added by adding one entry per row.
#
# The "FICA/FUTA/SUTA Taxable Only (...)" family is Paycom's label for
# traditional pre-tax retirement deferrals: the deferral reduces FIT and SIT
# taxable wages but FICA/FUTA/SUTA still apply. All three flavors below
# (401k, 403b, 457b) are PRE-TAX from UZIO's setup perspective.
CALC_DESC_TO_TAX_TREATMENT = {
    "s125 pre-tax":                          "Pre-tax",
    "after tax deduction":                   "Post Tax",
    "fica/futa/suta taxable only (401k)":    "Pre-tax",
    "fica/futa/suta taxable only (403b)":    "Pre-tax",
    "fica/futa/suta taxable only (457b)":    "Pre-tax",
}

# ─────────────────────────────────────────────────────────────────────────────
# UZIO "Add Deduction" form mapping config
#
# The Tampermonkey automation needs each deduction expressed in UZIO's terms:
#   Master Deductions List -> the dropdown that drives the whole form
#   Method                 -> Fixed $ / % of Gross Pay / % of Disposable Net Pay
#   Auto-Sync              -> only for benefit-type deductions (toggle in UI)
#
# These tables are intentionally explicit (not fuzzy) so the output is
# predictable and reviewable. When a Paycom code isn't found, we fall back to
# keyword inference, then to a "<NEEDS REVIEW>" sentinel that surfaces loudly
# in the UI and the xlsx so an implementor never ships a wrong guess.
# ─────────────────────────────────────────────────────────────────────────────

# Sentinel written into the file when we can't confidently map a value.
NEEDS_REVIEW = "<NEEDS REVIEW>"

# UZIO's generic catch-all master. When "Other" is selected, UZIO locks the
# Deduction Type to "Post Tax" and the implementor types the real name into the
# Deduction Name field (the master itself is just "Other"). So for these rows we
# keep the Paycom description as the Deduction Name and force the type Post Tax.
# (Confirmed against the manual ITR "Tuition Reimbursement" setup.)
MASTER_OTHER = "Other"

# Full UZIO "Master Deductions List" dropdown options (verbatim from the live
# UI). Used to populate the manual-mapping dropdown for deductions the tool
# can't confidently map (NEEDS_REVIEW). Keep in exact UZIO spelling.
UZIO_MASTER_DEDUCTIONS = [
    "401(k) Loan", "401k", "Accident Insurance After-tax", "Advance",
    "Basic Life and AD&D", "Cancer Insurance After-tax", "Cancer Insurance Pre-tax",
    "Child Support", "Child Support 2", "Creditor Garnishment",
    "Critical Illness After-tax", "Critical Illness Pre-tax", "Dental After-tax",
    "Dental Pre-tax", "Earned Wage Access", "Federal Tax Lien",
    "Gap Medical Pre-tax", "Group Term Life", "Health Cues Claim",
    "Health Cues Premium", "Health Reimbursement Arrangement (HRA) Pre-tax",
    "Health Savings Account(HSA) Pre-tax", "Hearing Insurance Pre-tax",
    "Hospital Indemnity After-tax", "Hospital Indemnity Pre-tax", "Loan",
    "Med Claim Reimbursement", "Med Plus Premium", "Medical After-tax",
    "Medical Pre-tax", "Overpayment", "Pet Insurance After-tax", "Reverse / Reissue",
    "Roth 401k", "Roth IRA", "Spousal Support Order", "State Tax Lien",
    "Student Loan", "Supplemental Life", "Supplemental Medical After-tax",
    "Supplemental Medical Pre-tax", "Vision After-tax", "Vision Pre-tax",
    "Voluntary AD&D After-tax", "Voluntary AD&D Pre-tax",
    "Voluntary Life Child After-tax", "Voluntary Life Child Pre-tax",
    "Voluntary Life Employee After-tax", "Voluntary Life Employee Pre-tax",
    "Voluntary Life Spouse After-tax", "Voluntary Life Spouse Pre-tax",
    "Voluntary LTD After-tax", "Voluntary STD After-tax", "Wellness Pre-tax",
    "Whole Life Insurance After-tax", "Other",
]

# Full UZIO "Earning Type" dropdown options (verbatim from the live UI). Used to
# populate the manual-mapping dropdown for earnings that fall to "Other".
UZIO_EARNING_TYPES = [
    "Bonus", "Commission", "Vacation", "Reimbursements",
    "Group Term Life Insurance", "Cash Tip", "Pay Check Tips",
    "Expense reimbursement", "Mileage reimbursement", "Stock Options",
    "Severance", "3rd Party Sick Pay - Taxable", "3rd Party Sick Pay - Nontaxable",
    "Dividend", "Moving Expenses", "Clothing Allowance", "Tool Allowance",
    "Tuition Assistance", "Non Tax Tuition Assistance", "Allocated Tips",
    "COVID 100 Sick", "COVID 2/3 Sick", "COVID Family Leave", "Sick", "Other",
    "Owner's Draw", "Unpaid Time Off", "OT Adjustment", "Station Closure",
    "DA Recognition - TWA",
]

# Method options EXACTLY as they appear in UZIO's Method dropdown. These must
# match character-for-character so the Tampermonkey script can select them.
# NOTE: for garnishment-style deductions UZIO's option is "% of Disposable Net
# Pay" (confirmed in the live UI) — NOT "% of Disposable Income".
METHOD_FIXED = "Fixed $"
METHOD_PCT_GROSS = "% of Gross Pay"
METHOD_PCT_DISPOSABLE = "% of Disposable Net Pay"

# Exact (Paycom Type Code) -> UZIO Master Deductions List entry.
# Built from the Chief Delivery review. Extend per-client as needed.
# NOTE: the UZIO "Master Deductions List" dropdown is dynamic — entries already
# created for a company disappear from it. The values below are still valid
# UZIO master names; the Tampermonkey script handles the "already exists" case
# at runtime by reading the company's existing Company Deductions list.
PAYCOM_CODE_TO_UZIO_MASTER = {
    "ACC": "Voluntary AD&D After-tax",
    "CIL": "Critical Illness After-tax",
    "CPC": "Med Claim Reimbursement",
    "CPP": "Med Plus Premium",
    "CS1": "Child Support",
    "CS2": "Child Support 2",
    "DEN": "Dental Pre-tax",
    "GP1": "Creditor Garnishment",
    "ITR": "Other",
    "K4P": "401k",
    "LN1": "401(k) Loan",
    "MDC": "Medical Pre-tax",
    "MLP": "Medical After-tax",
    "R4P": "Roth 401k",
    "STD": "Voluntary STD After-tax",
    "VEE": "Voluntary Life Employee After-tax",
    "VIS": "Vision Pre-tax",
}

# Master Deductions List entries that always use % of Disposable Net Pay
# (CCPA-limited wage garnishments). Lowercased for matching.
DISPOSABLE_INCOME_MASTERS = {
    "creditor garnishment",
    "federal tax lien",
    "state tax lien",
}

# Master Deductions List entries that always use Fixed $ regardless of the
# Paycom description (court-ordered fixed amounts, loans, advances, premiums).
FIXED_DOLLAR_MASTERS = {
    "child support",
    "child support 2",
    "spousal support order",
    "401(k) loan",
    "loan",
    "advance",
    "overpayment",
    "reverse / reissue",
}

# Benefit-type deductions show the "Auto-Sync from Uzio Benefits" radio on the
# UZIO form. Detected by keyword on the mapped Master Deductions List value.
# Per the implementor's spec: dental, medical, vision, voluntary life (child/
# spouse/employee), critical illness, accident insurance, cancer insurance,
# hospital indemnity, STD, AD&D.
# NOTE: Med Claim Reimbursement and Med Plus Premium are NOT benefit types
# (confirmed by the implementor) — they're custom reimbursement deductions. They
# still belong to ASSIGN_ALL_LOCKED_MASTERS (Assign-to-all forced Yes), but they
# do not Auto-Sync and do not track arrears.
BENEFIT_TYPE_KEYWORDS = (
    "dental", "medical", "vision", "voluntary life", "critical illness",
    "accident insurance", "cancer insurance", "hospital indemnity",
    "voluntary std", "ad&d",
)

# Masters whose "Assign to all employees" field is FORCED to "Yes" by UZIO and
# disabled (the script cannot change it). Lowercased for matching. These are the
# custom medical-reimbursement style deductions the implementor flagged.
ASSIGN_ALL_LOCKED_MASTERS = {
    "med claim reimbursement",
    "med plus premium",
    "health cues",
    "health cues premium",
}

# Static UZIO Add-Deduction field defaults (constant for every row right now;
# promoted to named constants so the one place to change them is obvious).
DEFAULT_DEDUCTION_SCHEDULE = "Every Paycheck"
DEFAULT_ASSIGN_TO_ALL = "No"
ARREARS_PROCESSING_TOTAL = "Total Amount"
AUTOSYNC_NA = "N/A"
# W-2 Box behavior depends on the master:
#   - A real master entry AUTO-FILLS the W-2 Box and DISABLES the field, so the
#     automation must not set it. We emit a locked marker so the Tampermonkey
#     script knows to skip it.
#   - "Other"/unmapped masters leave W-2 Box enabled; the implementor enters
#     "Not Required" (the normal value for these).
DEFAULT_W2_BOX = "Not Required"
W2_BOX_LOCKED = "(Auto-filled by UZIO - do not set)"

# Canonical column order for the Deductions tab + the UI dataframe, so the
# spreadsheet, the on-screen table, and the Tampermonkey consumer all agree.
DEDUCTION_OUTPUT_COLUMNS = [
    "Type Code", "Type Description", "Calc Description", "Pre/Post Tax",
    "UZIO Master Deductions List", "UZIO Deduction Type", "UZIO Deduction Name",
    "UZIO Method", "Amount per pay", "Auto-Sync from Uzio Benefits",
    "Assign to all employees", "Deduction Schedule", "Track arrears",
    "Arrears Processing Method", "W-2 Box",
]

# ─────────────────────────────────────────────────────────────────────────────
# UZIO "Add Earning" form mapping config
#
# The Earning form fields:
#   Earning Type (dropdown, driver)  Earning Name (text)  Display Order (text)
#   Paid Earning (Yes/No)  Hourly Based Earning (Yes/No)
#   Subject to garnishment disposable income? (Yes/No)
#   Subject to Workers' Compensation (Yes/No)  Taxability Type (dropdown)
#   W-2 Box (dropdown)
#
# Fields vary per earning, so the tool emits rule-based DEFAULTS as columns that
# the implementor can edit in the Excel before feeding the Tampermonkey script.
# The Earning Type dropdown is the driver (like Master Deductions List was).
# ─────────────────────────────────────────────────────────────────────────────

# All values below are DATA-DRIVEN from the 52-client / 1090-earning DSP
# database export ("Earning with earning_type.csv"). The earning_type column
# there gave the exact UZIO Earning Type per earning; time_bounded gave Hourly;
# taxability_type gave Taxable/Non-Taxable; is_default + the known auto-created
# set told us which earnings UZIO seeds (and we must NOT re-create).

EARNING_TYPE_OTHER = "Other"
EARNING_TAXABILITY_TAXABLE = "Taxable"
EARNING_TAXABILITY_NONTAX = "Non-Taxable"

# ── (1) UZIO auto-created / default earnings — SKIP, never re-create ──────────
# These are seeded by UZIO on every company (confirmed in the DSP DB: each
# appears ~once per company; matches the implementor's UZIO earnings screenshot).
# Detected from the Paycom (code, description) via DEFAULT_EARNING_RULES below,
# because Paycom names differ from UZIO's (e.g. Paycom "Regular" = UZIO
# "Regular Wage"; Paycom WOT "Overtime Hours (Weighted)" = UZIO "Overtime").
# Each rule: (predicate over normalized desc) -> UZIO default name. Ordered;
# first match wins. Guards prevent false hits (e.g. "Retro Regular Pay" must NOT
# match Regular Wage).
def _edesc(type_description):
    return " ".join((type_description or "").lower().split())

DEFAULT_EARNING_RULES = [
    # (matcher(desc) -> bool, UZIO default earning name)
    (lambda d: "look back" in d or "lookback" in d,                 "Lookback bonus"),
    (lambda d: "realtime" in d or "real time" in d,                 "Realtime bonus"),
    (lambda d: "ot adjustment" in d or d == "otadj",                "OT Adjustment"),
    (lambda d: "double" in d and "overtime" in d,                   "Double Overtime"),
    # Weighted overtime (Paycom WOT) is the UZIO default "Overtime".
    (lambda d: "overtime" in d and ("weighted" in d or "(weighted)" in d), "Overtime"),
    (lambda d: d in ("overtime", "overtime hours", "ot"),           "Overtime"),
    (lambda d: "holiday" in d and "premium" in d,                   "Holiday Premium"),
    (lambda d: d == "holiday",                                      "Holiday"),
    (lambda d: "meal break" in d,                                   "Meal Break Premium"),
    (lambda d: "rest break" in d,                                   "Rest Break Premium"),
    # Retro Overtime Pay is a system-created default (Earning Type OT Adjustment).
    (lambda d: "retro" in d and "overtime" in d and "pay" in d,     "OT Adjustment"),
    # NOTE: "Makeup Pay" is NOT in UZIO's seeded earnings list (confirmed against
    # the live UZIO Earnings table), so it is intentionally NOT skipped here — a
    # "Makeup Pay" earning must be CREATED (it falls through to Earning Type
    # "Other"). Do not re-add it to the skip rules.
    (lambda d: "pto" in d and ("balance" in d or "payout" in d),    "PTO Balance Payout"),
    # Plain "Reimbursements" only — NOT Expense/Mileage/Tuition reimbursement
    # (those are their own creatable Earning Types).
    (lambda d: d in ("reimbursements", "reimbursement"),            "Reimbursements"),
    # Regular Wage: starts with / equals "regular" and is NOT a retro variant.
    (lambda d: (d == "regular" or d == "regular wage" or d == "regular pay"
                or d.startswith("regular ")) and "retro" not in d,  "Regular Wage"),
]


def default_earning_name(type_code, type_description):
    """Return the UZIO default earning name this Paycom earning corresponds to
    (so it gets SKIPPED), or "" if it's a real earning to create."""
    d = _edesc(type_description)
    for pred, uzio_name in DEFAULT_EARNING_RULES:
        try:
            if pred(d):
                return uzio_name
        except Exception:
            pass
    return ""


# ── (2) Earning Type mapping (driver dropdown) ───────────────────────────────
# Keyword -> UZIO Earning Type dropdown label. Ordered; first match wins.
# Derived from the DB's name->earning_type pairs. Unmatched -> "Other" (which is
# the correct, common answer for DSP-specific custom earnings per the data).
EARNING_TYPE_KEYWORD_MAP = [
    ("unpaid time off",  "Unpaid Time Off"),
    ("unpaid leave",     "Unpaid Time Off"),
    ("vto",              "Unpaid Time Off"),
    ("paid time off",    "Vacation"),
    ("vacation",         "Vacation"),
    ("pto",              "Vacation"),
    # Tuition: reimbursement / non-tax -> Non Tax Tuition; plain assistance ->
    # Tuition Assistance.
    ("tuition reimbursement",     "Non Tax Tuition Assistance"),
    ("non tax tuition",           "Non Tax Tuition Assistance"),
    ("tuition",                   "Tuition Assistance"),
    ("expense reimburs", "Expense reimbursement"),
    ("mileage",          "Mileage reimbursement"),
    ("reimburs",         "Reimbursements"),
    ("commission",       "Commission"),
    ("severance",        "Severance"),
    ("group term life",  "Group Term Life Insurance"),
    ("paycheck tip",     "Pay Check Tips"),
    ("pay check tip",    "Pay Check Tips"),
    ("tip",              "Cash Tip"),
    ("sick",             "Sick"),
    ("bonus",            "Bonus"),
    # Any "station closure" (incl. "Station Closure Payment") -> Station Closure
    # earning type (per implementor).
    ("station closure",  "Station Closure"),
]
EARNING_TYPE_EXACT_MAP = {}

# ── (3) Hourly + Taxability are driven by Earning Type (from time_bounded /
# taxability_type in the DB). For a mapped type UZIO auto-fills+locks these, so
# they're informational in the Excel; for "Other" they're editable defaults. ──
EARNING_TYPE_HOURLY_NO = {
    "Bonus", "OT Adjustment", "Reimbursements", "Expense reimbursement",
    "Mileage reimbursement", "Tuition Assistance", "Non Tax Tuition Assistance",
    "Severance", "Commission", "Cash Tip", "Pay Check Tips",
}
EARNING_TYPE_NONTAX = {
    "Reimbursements", "Expense reimbursement", "Mileage reimbursement",
    "Non Tax Tuition Assistance",
}

# Defaults for "Other" earnings (per implementor: hourly Yes, Taxable).
EARNING_DEFAULT_PAID = "Yes"
EARNING_OTHER_HOURLY = "Yes"
EARNING_DEFAULT_DISPOSABLE = "Yes"
EARNING_DEFAULT_WORKERS_COMP = "Yes"
EARNING_DEFAULT_W2_BOX = "Not Required"

# Per-Earning-Type overrides for the "disposable income" / "Workers' Comp" radios.
# Those two are editable for a few named types (not just "Other"), so the tool
# sets the right default and the userscript applies it (it self-skips if the field
# turns out locked). Reimbursements -> both No. "DA Recognition - TWA" keeps the
# Yes/Yes default, so it needs no entry here.
EARNING_TYPE_FIELD_DEFAULTS = {
    "Reimbursements": {"disposable": "No", "workersComp": "No"},
}

# Rate Determination Factor only appears for "Other" earnings with Hourly=Yes.
# Then a "Rate" box appears where the value is 1. For everything else: "NA".
EARNING_RATE_FACTOR_MULTIPLES = "Multiples of Regular Wage Rate"
EARNING_RATE_DEFAULT_VALUE = "1"
EARNING_NA = "NA"

# "Include Bonus in Overtime Rate Calculation?" appears for bonus earnings
# (except Lookback / Realtime, which are system defaults and skipped anyway).
# Default No. This is ALSO the discretionary determination:
#   Yes => non-discretionary bonus (must be in OT rate)
#   No  => discretionary bonus
EARNING_INCLUDE_OT_COL = "Include Bonus in Overtime Calculation"
EARNING_INCLUDE_OT_DEFAULT = "No"

# "Time Off Policies" dropdown only appears for time-off Earning Types:
#   - Vacation (e.g. Amazon "Paid Time Off") -> "All" (all policies)
#   - Unpaid Time Off                         -> "All" (all policies)
#   - every other type                        -> NA (the field isn't shown)
# Editable in the Excel if a client needs a specific policy instead of All.
EARNING_TIMEOFF_COL = "Time Off Policy"
EARNING_TIMEOFF_VACATION = "All"
EARNING_TIMEOFF_ALL = "All"

# Canonical column order for the Earnings tab + UI dataframe.
EARNING_OUTPUT_COLUMNS = [
    "Type Code", "Type Description", "Earning Type", "Earning Name",
    "Display Order", "Paid Earning", "Hourly Based Earning",
    "Rate Determination Factor", "Rate",
    "Subject to garnishment disposable income", "Subject to Workers Compensation",
    "Taxability Type", EARNING_INCLUDE_OT_COL, EARNING_TIMEOFF_COL, "W-2 Box",
]


def is_bonus_earning(type_description):
    """True if this earning is a bonus that gets the 'Include in Overtime' /
    discretionary question — i.e. contains 'bonus' (case-insensitive) but is NOT
    a Lookback or Realtime bonus (those are system defaults, skipped). Also
    catches LK2/LKB-style lookback codes via the description."""
    d = _edesc(type_description)
    if "bonus" not in d:
        return False
    if "look back" in d or "lookback" in d or "realtime" in d or "real time" in d:
        return False
    return True


def map_paycom_to_earning_type(type_code, type_description):
    """Return the UZIO Earning Type for a Paycom earning. Forced-Other rules
    first, then exact map, then keyword map, then "Other"."""
    d = _edesc(type_description)
    # FORCE Other: "Bonus Hours" / "Bonus (Hours)" (case-insensitive). The UZIO
    # Bonus type locks Hourly Based = No (can't edit), but an hourly bonus must
    # be hourly — so we create it as "Other" with Hourly=Yes (which then gets
    # Rate Determination Factor = Multiples of Regular Wage Rate, Rate = 1).
    if "bonus" in d and "hour" in d:
        return EARNING_TYPE_OTHER
    if d in EARNING_TYPE_EXACT_MAP:
        return EARNING_TYPE_EXACT_MAP[d]
    blob = f"{type_code or ''} {type_description or ''}".lower()
    for kw, etype in EARNING_TYPE_KEYWORD_MAP:
        if kw in blob:
            return etype
    return EARNING_TYPE_OTHER


def filter_default_uzio_earnings(rows):
    """Split extracted earnings into (kept, skipped). UZIO auto-creates a set of
    earnings on every company (Regular Wage, Overtime, Holiday, etc.); the
    automation must not re-create them. Each skipped row is annotated with the
    matched UZIO default name. Returns (kept_rows, skipped_rows)."""
    if not rows:
        return [], []
    kept, skipped = [], []
    for r in rows:
        uzio_name = default_earning_name(r.get("Type Code"), r.get("Type Description"))
        if uzio_name:
            skipped.append({**r, "UZIO Default Earning": uzio_name})
        else:
            kept.append(r)
    return kept, skipped


def enrich_earnings_for_uzio(rows, start_display_order=20, include_in_ot_map=None,
                             earning_type_override_map=None):
    """Add the UZIO Add-Earning form fields to each (already default-filtered)
    earning row as columns.

    `start_display_order`: Display Order for the first created earning; each
    subsequent one increments by 1.

    `include_in_ot_map`: optional {autosync_row_key(code, desc) -> "Yes"/"No"}
    from the "Is the earning Non Discretionary?" UI section. Only consulted for
    bonus earnings (is_bonus_earning); defaults to No. Non-bonus rows get "NA".

    Field values: Earning Type drives Hourly + Taxability (auto-filled & locked
    by UZIO for mapped types; editable for "Other"). Disposable / Workers' Comp /
    Paid default to Yes; W-2 Not Required. All editable in the Excel.
    """
    include_in_ot_map = include_in_ot_map or {}
    earning_type_override_map = earning_type_override_map or {}
    out = []
    order = start_display_order
    for r in rows:
        etype = map_paycom_to_earning_type(r.get("Type Code"), r.get("Type Description"))
        # Manual override from the UI for earnings that fell to "Other".
        ovr = earning_type_override_map.get(autosync_row_key(r.get("Type Code"), r.get("Type Description")))
        if ovr:
            etype = ovr
        if etype == EARNING_TYPE_OTHER:
            hourly = EARNING_OTHER_HOURLY
            taxability = EARNING_TAXABILITY_TAXABLE
        else:
            hourly = "No" if etype in EARNING_TYPE_HOURLY_NO else "Yes"
            taxability = EARNING_TAXABILITY_NONTAX if etype in EARNING_TYPE_NONTAX else EARNING_TAXABILITY_TAXABLE
        # Rate Determination Factor + Rate:
        #   - "Other" + Hourly=Yes -> Multiples of Regular Wage Rate, Rate = 1
        #   - "Unpaid Time Off"    -> Multiples of Regular Wage Rate, Rate = 0
        #   - everything else      -> NA
        if etype == EARNING_TYPE_OTHER and hourly == "Yes":
            rate_factor = EARNING_RATE_FACTOR_MULTIPLES
            rate_value = EARNING_RATE_DEFAULT_VALUE   # "1"
        elif etype == "Unpaid Time Off":
            rate_factor = EARNING_RATE_FACTOR_MULTIPLES
            rate_value = "0"
        else:
            rate_factor = EARNING_NA
            rate_value = EARNING_NA
        # Include-in-overtime applies ONLY when the (possibly overridden) Earning
        # Type is exactly "Bonus" — that's the only UZIO type that shows the
        # "Include Bonus in Overtime Rate Calculation?" field. For every other type
        # (including "Other") the field doesn't exist, so emit "NA". Driving this
        # off the final etype (not the name) means changing Bonus→Other flips the
        # value to NA, and Other→Bonus flips it back to the No/Yes default.
        if etype == "Bonus":
            include_ot = include_in_ot_map.get(
                autosync_row_key(r.get("Type Code"), r.get("Type Description")),
                EARNING_INCLUDE_OT_DEFAULT,
            )
        else:
            include_ot = EARNING_NA
        # Disposable-income / Workers'-Comp defaults, with per-type overrides
        # (e.g. Reimbursements -> No/No).
        disposable = EARNING_DEFAULT_DISPOSABLE
        workers_comp = EARNING_DEFAULT_WORKERS_COMP
        _fov = EARNING_TYPE_FIELD_DEFAULTS.get(etype)
        if _fov:
            disposable = _fov.get("disposable", disposable)
            workers_comp = _fov.get("workersComp", workers_comp)
        # Time Off Policies dropdown: Vacation -> "Paid PTO", Unpaid Time Off ->
        # "All", everything else -> NA (field not shown for other types).
        if etype == "Vacation":
            timeoff = EARNING_TIMEOFF_VACATION
        elif etype == "Unpaid Time Off":
            timeoff = EARNING_TIMEOFF_ALL
        else:
            timeoff = EARNING_NA
        out.append({
            **r,
            "Earning Type": etype,
            "Earning Name": r.get("Type Description", ""),
            "Display Order": str(order),
            "Paid Earning": EARNING_DEFAULT_PAID,
            "Hourly Based Earning": hourly,
            "Rate Determination Factor": rate_factor,
            "Rate": rate_value,
            "Subject to garnishment disposable income": disposable,
            "Subject to Workers Compensation": workers_comp,
            "Taxability Type": taxability,
            EARNING_INCLUDE_OT_COL: include_ot,
            EARNING_TIMEOFF_COL: timeoff,
            "W-2 Box": EARNING_DEFAULT_W2_BOX,
        })
        order += 1
    return out


# ─────────────────────────────────────────────────────────────────────────────
# UZIO "Add Contribution" form mapping config
#
# The Contribution form is simpler than the deduction form: a free-text
# Contribution Name, an optional link to a company deduction, a Method, optional
# Monthly/Annual limits, a W-2 Box, and an assign-to-all question. There is NO
# Master Deductions List dropdown.
#
# Method = "Formula" is a tiered employer-match formula. For DSP/Paycom 401k &
# Roth matches the standard safe-harbor formula is:
#     Tier 1: 100% of the first 1%
#     Tier 2:  50% of the next  4%
# The Tampermonkey script selects "Formula", fills tier 1 (100 / 1), clicks
# "Add more", then fills tier 2 (50 / 4). Encoded as constants so the one place
# to change the formula is obvious.
# ─────────────────────────────────────────────────────────────────────────────
CONTRIB_METHOD_FORMULA = "Formula"
CONTRIB_METHOD_FIXED = "Fixed $"
# Each tier is (match_percent, up_to_percent). Order = the order rows are added.
# Used ONLY for 401k / Roth 401k matches; all other contributions use Fixed $.
CONTRIB_FORMULA_TIERS = [
    (100, 1),
    (50, 4),
]
CONTRIB_DEFAULT_ASSIGN_TO_ALL = "No"
CONTRIB_DEFAULT_W2_BOX = "Not Required"
# Sentinel shown in the UI dropdown for "do not link this contribution".
CONTRIB_LINK_NONE = "(none - do not link)"

# Canonical column order for the Contributions tab + UI dataframe.
CONTRIBUTION_OUTPUT_COLUMNS = [
    "Type Code", "Type Description", "Contribution Name",
    "Link to Company Deduction", "Linked Deduction", "Method",
    "Formula", "Monthly Limit", "Annual Limit", "W-2 Box",
    "Assign to all employees",
]


def _format_formula(tiers):
    """Human-readable formula string for the Excel/UI, e.g.
    '100% of first 1%; 50% of next 4%'. The Tampermonkey script reads the
    structured tiers from CONTRIB_FORMULA_TIERS, not this string."""
    parts = []
    for i, (match, upto) in enumerate(tiers):
        word = "first" if i == 0 else "next"
        parts.append(f"{match}% of {word} {upto}%")
    return "; ".join(parts)


def map_contribution_to_deduction(type_code, type_description, available_deduction_masters):
    """Pick the default UZIO deduction to link a contribution to.

    `available_deduction_masters` is the set/list of UZIO Master Deductions List
    names actually created for this client (so we never default-link to a
    deduction that doesn't exist).

    Defaults (only applied when the target deduction is available):
      - Roth 401k match  -> "Roth 401k"
      - 401k match       -> "401k"
      - Medical ER memo  -> "Medical Pre-tax" if present, else "Medical After-tax"
    Anything else -> "" (no default; the user links it manually in the UI).
    """
    avail = {(_m or "").strip().lower() for _m in (available_deduction_masters or [])}
    desc = (type_description or "").strip().lower()

    def have(name):
        return name.strip().lower() in avail

    is_roth = "roth" in desc
    is_401k = "401k" in desc or "401(k)" in desc
    is_medical = "medical" in desc or "med er" in desc or "med memo" in desc or "medical er" in desc

    if is_roth and have("Roth 401k"):
        return "Roth 401k"
    if is_401k and have("401k"):
        return "401k"
    if is_medical:
        if have("Medical Pre-tax"):
            return "Medical Pre-tax"
        if have("Medical After-tax"):
            return "Medical After-tax"
    return ""


def is_formula_contribution(type_description):
    """True if this contribution uses the tiered match Formula (401k / Roth 401k
    matches). Everything else uses Fixed $."""
    desc = (type_description or "").strip().lower()
    return "401k" in desc or "401(k)" in desc


def enrich_contributions_for_uzio(rows, link_map=None):
    """Add the UZIO Add-Contribution form fields to each contribution row.

    `link_map`: optional {autosync_row_key(code, desc) -> linked deduction name
    or "" }. "" / CONTRIB_LINK_NONE means do not link.

    Columns added (see CONTRIBUTION_OUTPUT_COLUMNS):
      - Contribution Name             = Type Description verbatim
      - Link to Company Deduction     = "Yes" if a linked deduction is chosen else "No"
      - Linked Deduction              = the deduction name (blank if not linked)
      - Method                        = "Formula" for 401k/Roth 401k, else "Fixed $"
      - Formula                       = tiers string (only for Formula rows)
      - Monthly Limit / Annual Limit  = "" (optional)
      - W-2 Box                       = "Not Required"
      - Assign to all employees       = "No"
    """
    link_map = link_map or {}
    out = []
    for r in rows:
        key = autosync_row_key(r.get("Type Code"), r.get("Type Description"))
        linked = (link_map.get(key) or "").strip()
        if linked in ("", CONTRIB_LINK_NONE):
            linked, link_yn = "", "No"
        else:
            link_yn = "Yes"
        use_formula = is_formula_contribution(r.get("Type Description"))
        out.append({
            **r,
            "Contribution Name": r.get("Type Description", ""),
            "Link to Company Deduction": link_yn,
            "Linked Deduction": linked,
            "Method": CONTRIB_METHOD_FORMULA if use_formula else CONTRIB_METHOD_FIXED,
            "Formula": _format_formula(CONTRIB_FORMULA_TIERS) if use_formula else "",
            "Monthly Limit": "",
            "Annual Limit": "",
            "W-2 Box": CONTRIB_DEFAULT_W2_BOX,
            "Assign to all employees": CONTRIB_DEFAULT_ASSIGN_TO_ALL,
        })
    return out


# Keyword families that have BOTH a Pre-tax and an After-tax master in UZIO.
# The variant chosen depends on the deduction's tax treatment (the Pre/Post Tax
# verdict), so e.g. "Dental" with After Tax Deduction -> "Dental After-tax".
# (kw, pre_tax_master, after_tax_master)
TAX_PAIRED_KEYWORD_MASTERS = [
    ("dental",            "Dental Pre-tax",            "Dental After-tax"),
    ("vision",            "Vision Pre-tax",            "Vision After-tax"),
    ("medical",           "Medical Pre-tax",           "Medical After-tax"),
    ("critical illness",  "Critical Illness Pre-tax",  "Critical Illness After-tax"),
    ("cancer",            "Cancer Insurance Pre-tax",  "Cancer Insurance After-tax"),
    ("hospital indemnity", "Hospital Indemnity Pre-tax", "Hospital Indemnity After-tax"),
]

# Keyword families with a SINGLE master regardless of tax treatment.
# ORDER MATTERS:
#  - "spousal" must precede the generic child-support / support-order rules so
#    "Spousal Support Order" -> Spousal Support Order (NOT Child Support).
#  - "support order" / "child support" both -> Child Support (per implementor).
#  - "401k loan" must precede "401k" so the loan isn't swallowed by 401k.
SINGLE_KEYWORD_MASTERS = [
    ("spousal",       "Spousal Support Order"),
    ("child support", "Child Support"),
    ("support order", "Child Support"),
    ("garnish",       "Creditor Garnishment"),
    ("roth",          "Roth 401k"),
    ("401k loan",     "401(k) Loan"),
    ("401(k) loan",   "401(k) Loan"),
    ("401k",          "401k"),
    ("hsa",           "Health Savings Account(HSA) Pre-tax"),
]


def map_paycom_to_uzio_master(type_code, type_description, tax_treatment=""):
    """Return the UZIO Master Deductions List value for a Paycom deduction.

    Order: explicit code table -> keyword inference on description (tax-aware for
    Pre/After-tax paired families) -> NEEDS_REVIEW sentinel.

    `tax_treatment` is the Pre/Post Tax verdict ("Pre-tax" / "Post Tax" /
    "Unknown"). For paired families it selects the After-tax variant when the
    verdict is "Post Tax"; otherwise the Pre-tax variant.
    """
    code = (type_code or "").strip().upper()
    if code in PAYCOM_CODE_TO_UZIO_MASTER:
        return PAYCOM_CODE_TO_UZIO_MASTER[code]

    # Light keyword inference for codes we haven't explicitly mapped.
    desc = (type_description or "").strip().lower()
    is_post = (tax_treatment or "").strip().lower() == "post tax"

    # Reverse / Reissue is recognized only when BOTH "reverse" AND "issue" appear
    # in the description (per the implementor). "reissue" satisfies "issue", so
    # "Reverse / Reissue", "Reverse Reissue", "Payroll Reverse - Reissued" all
    # match, while a lone "reverse ..." or unrelated "...issue..." does not.
    if "reverse" in desc and "issue" in desc:
        return "Reverse / Reissue"

    # Tax-paired families first: pick After-tax vs Pre-tax by the verdict.
    for kw, pre_master, post_master in TAX_PAIRED_KEYWORD_MASTERS:
        if kw in desc:
            return post_master if is_post else pre_master

    for kw, master in SINGLE_KEYWORD_MASTERS:
        if kw in desc:
            return master
    return NEEDS_REVIEW


def determine_method(uzio_master, type_description):
    """Return the UZIO Method value for a deduction.

    Rules (in priority order):
      1. Garnishment-style masters -> % of Disposable Net Pay
      2. Fixed-dollar masters (child support, loans, etc.) -> Fixed $
      3. Paycom description contains "%" -> % of Gross Pay
      4. Otherwise -> Fixed $
    """
    master_l = (uzio_master or "").strip().lower()
    if master_l in DISPOSABLE_INCOME_MASTERS:
        return METHOD_PCT_DISPOSABLE
    if master_l in FIXED_DOLLAR_MASTERS:
        return METHOD_FIXED
    if "%" in (type_description or ""):
        return METHOD_PCT_GROSS
    return METHOD_FIXED


def is_benefit_type(uzio_master):
    """True if this Master Deductions List value shows the Auto-Sync radio."""
    m = (uzio_master or "").strip().lower()
    return any(kw in m for kw in BENEFIT_TYPE_KEYWORDS)


def is_assign_all_locked(uzio_master):
    """True if UZIO forces 'Assign to all employees' = Yes (and disables it)
    for this Master Deductions List value."""
    return (uzio_master or "").strip().lower() in ASSIGN_ALL_LOCKED_MASTERS


def autosync_row_key(type_code, type_description):
    """Stable, unique key for a deduction row, shared by the Auto-Sync toggle
    map (built in the UI) and the lookup in enrich_deductions_for_uzio. Keying on
    Type Code alone is unsafe (a code can repeat with different descriptions) and
    Type Description alone is unsafe (two rows can share a description, e.g. two
    "Medical" lines). The (code, description) pair is unique because the upstream
    extractor dedupes on exactly that pair."""
    return f"{(type_code or '').strip()}||{(type_description or '').strip()}"


def classify_pre_post_from_calc_description(prior_df, rows):
    """For each (Type Code, Type Description) in `rows`, look up its
    `Calc Description` in `prior_df`'s Deductions rows and map to a UZIO
    tax-treatment verdict.

    Adds two fields to each row in-place-style (we return a fresh list):
      - "Calc Description": the most-common raw Paycom value seen
      - "Pre/Post Tax": one of "Pre-tax" / "Post Tax" / "Unknown"

    Mapping is driven by CALC_DESC_TO_TAX_TREATMENT (case-insensitive).
    Any Calc Description not in the map lands in "Unknown" with the raw
    text preserved so the user can see exactly what came through and tell
    us how to extend the mapping.
    """
    if not rows:
        return rows

    if "Calc Description" not in prior_df.columns or "Code Description" not in prior_df.columns:
        return [{**r, "Calc Description": "(column missing)", "Pre/Post Tax": "Unknown"} for r in rows]

    deds = prior_df[prior_df["Code Description"].astype(str).str.strip() == "Deductions"]

    out = []
    for r in rows:
        tc = r["Type Code"]
        td = r["Type Description"]
        sub = deds[
            (deds["Type Code"].astype(str).str.strip() == tc)
            & (deds["Type Description"].astype(str).str.strip() == td)
        ]
        calc_descs = sub["Calc Description"].dropna().astype(str).str.strip()
        if calc_descs.empty:
            raw, verdict = "", "Unknown"
        else:
            # Take the most common value when there's variance across rows
            # (extremely rare but defensive).
            raw = calc_descs.mode().iloc[0]
            verdict = CALC_DESC_TO_TAX_TREATMENT.get(_normalize_calc_desc(raw), "Unknown")
        out.append({**r, "Calc Description": raw, "Pre/Post Tax": verdict})
    return out


def build_3tab_setup_xlsx(earnings, deductions, contributions):
    """Build a single .xlsx with three tabs in this order:
        1. Earnings       (Type Code, Type Description)
        2. Deductions     (Type Code, Type Description, Calc Description, Pre/Post Tax)
        3. Contributions  (Type Code, Type Description)

    Each tab has a frozen header row, sensible column widths, and the Pre/Post Tax
    column on the Deductions tab is color-coded (green = Pre-tax, red = Post Tax,
    grey = Unknown). Returns the xlsx as bytes for st.download_button.
    """
    import xlsxwriter  # local import keeps the module importable even if missing

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book

        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#1F4E78", "font_color": "white",
            "border": 1, "align": "left", "valign": "vcenter",
        })
        pre_fmt = wb.add_format({
            "bold": True, "bg_color": "#C6EFCE", "font_color": "#006100",
            "align": "center", "valign": "vcenter",
        })
        post_fmt = wb.add_format({
            "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006",
            "align": "center", "valign": "vcenter",
        })
        unknown_fmt = wb.add_format({
            "bold": True, "bg_color": "#F0F0F0", "font_color": "#666666",
            "align": "center", "valign": "vcenter",
        })

        # ── Tab 1: Earnings ──────────────────────────────────────────────
        # Enriched rows carry the full UZIO Add-Earning column set; pin the
        # header so an un-enriched fallback still produces a stable sheet.
        earn_cols = EARNING_OUTPUT_COLUMNS
        df_e = (pd.DataFrame(earnings, columns=earn_cols)
                if earnings else pd.DataFrame(columns=earn_cols))
        df_e.to_excel(writer, sheet_name="Earnings", index=False)
        ws_e = writer.sheets["Earnings"]
        earn_widths = {
            "Type Code": 12, "Type Description": 28, "Earning Type": 22,
            "Earning Name": 28, "Display Order": 13, "Paid Earning": 13,
            "Hourly Based Earning": 18, "Rate Determination Factor": 28, "Rate": 8,
            "Subject to garnishment disposable income": 36,
            "Subject to Workers Compensation": 28,
            "Taxability Type": 16,
            "Include Bonus in Overtime Calculation": 32,
            "Time Off Policy": 16, "W-2 Box": 14,
        }
        for i, c in enumerate(earn_cols):
            ws_e.set_column(i, i, earn_widths.get(c, 18))
            ws_e.write(0, i, c, header_fmt)
        ws_e.set_row(0, 24)
        ws_e.freeze_panes(1, 0)

        # ── Tab 2: Deductions ────────────────────────────────────────────
        # Full UZIO Add-Deduction field set as columns (see
        # DEDUCTION_OUTPUT_COLUMNS). Rows may or may not be enriched; pinning
        # the column list guarantees a stable header even for un-enriched input.
        ded_cols = DEDUCTION_OUTPUT_COLUMNS
        df_d = (pd.DataFrame(deductions, columns=ded_cols)
                if deductions else pd.DataFrame(columns=ded_cols))
        df_d.to_excel(writer, sheet_name="Deductions", index=False)
        ws_d = writer.sheets["Deductions"]
        # Per-column widths keyed by header name (order-independent).
        ded_widths = {
            "Type Code": 12, "Type Description": 28, "Calc Description": 34,
            "Pre/Post Tax": 12, "UZIO Master Deductions List": 30,
            "UZIO Deduction Type": 16, "UZIO Deduction Name": 30,
            "UZIO Method": 22, "Amount per pay": 14,
            "Auto-Sync from Uzio Benefits": 24, "Assign to all employees": 20,
            "Deduction Schedule": 18, "Track arrears": 13,
            "Arrears Processing Method": 22, "W-2 Box": 10,
        }
        pre_post_idx = ded_cols.index("Pre/Post Tax")
        for i, c in enumerate(ded_cols):
            ws_d.set_column(i, i, ded_widths.get(c, 18))
            ws_d.write(0, i, c, header_fmt)
        ws_d.set_row(0, 24)
        ws_d.freeze_panes(1, 0)
        # Color-code the Pre/Post Tax column.
        for ri, r in enumerate(deductions or [], start=1):
            v = (r.get("Pre/Post Tax") or "").strip()
            if v == "Pre-tax":
                ws_d.write(ri, pre_post_idx, v, pre_fmt)
            elif v == "Post Tax":
                ws_d.write(ri, pre_post_idx, v, post_fmt)
            elif v:
                ws_d.write(ri, pre_post_idx, v, unknown_fmt)

        # ── Tab 3: Contributions ─────────────────────────────────────────
        # Enriched rows carry the full UZIO Add-Contribution column set; pin the
        # header so an un-enriched fallback still produces a stable sheet.
        contrib_cols = CONTRIBUTION_OUTPUT_COLUMNS
        df_c = (pd.DataFrame(contributions, columns=contrib_cols)
                if contributions else pd.DataFrame(columns=contrib_cols))
        df_c.to_excel(writer, sheet_name="Contributions", index=False)
        ws_c = writer.sheets["Contributions"]
        contrib_widths = {
            "Type Code": 12, "Type Description": 28, "Contribution Name": 28,
            "Link to Company Deduction": 22, "Linked Deduction": 24, "Method": 14,
            "Formula": 30, "Monthly Limit": 14, "Annual Limit": 14, "W-2 Box": 14,
            "Assign to all employees": 20,
        }
        for i, c in enumerate(contrib_cols):
            ws_c.set_column(i, i, contrib_widths.get(c, 18))
            ws_c.write(0, i, c, header_fmt)
        ws_c.set_row(0, 24)
        ws_c.freeze_panes(1, 0)

    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# API "Mapping files" (Source name -> UZIO name)
#
# The API run takes TWO inputs: (1) the original SOURCE file (the Paycom Prior
# Payroll), and (2) three MAPPING files (Earnings / Deductions / Contributions).
# Each mapping file is a 4-column CSV that translates the source name into the
# exact UZIO name. Columns are verbatim from the client templates — DO NOT change
# the spelling/spacing of the headers:
#     Source <X> Code, Source <X> Code Name, Uzio <X> Code, Uzio <X> Code Name
# Only the two *Name* columns are filled; the two *Code* columns stay blank
# (matching the provided templates; the API matches on name).
#
# CRITICAL: the "Source ... Code Name" must match the source file VERBATIM —
# including any leading/trailing spaces — or the API silently drops the row.
# So we emit the RAW (unstripped) Type Description captured from the source via
# _raw_name_map, NOT the stripped value used everywhere else in this tool.
#
# The "Uzio ... Code Name" is taken from the exact same field the Tampermonkey
# script types into UZIO (Earning Name / UZIO Deduction Name / Contribution
# Name), so the mapping can never drift from what actually gets created.
#
# The EARNINGS mapping intentionally also lists the UZIO *default* earnings
# (Regular Wage, Overtime, Holiday, Lookback bonus, ...) that are skipped in the
# setup Excel, because the API still needs to upload data against those existing
# (system-created) earnings. Those UZIO names come from DEFAULT_EARNING_RULES.
# ─────────────────────────────────────────────────────────────────────────────

MAPPING_EARNING_COLUMNS = [
    "Source Earning Code", "Source Earning Code Name",
    "Uzio Earning Code", "Uzio Earning Code Name",
]
MAPPING_DEDUCTION_COLUMNS = [
    "Source Deduction Code", "Source Deduction Code Name",
    "Uzio Deduction Code", "Uzio Deduction Code Name",
]
MAPPING_CONTRIBUTION_COLUMNS = [
    "Source Contribution Code", "Source Contribution Code Name",
    "Uzio Contribution Code", "Uzio Contribution Code Name",
]


def _raw_name_map(prior_df, code_description_value):
    """{autosync_row_key(stripped_code, stripped_desc) -> RAW (unstripped) Type
    Description as it literally appears in the source file}.

    When one key has several raw spellings (e.g. "Bonus" and "Bonus "), the most
    frequent raw form wins. This lets the mapping file's Source name match the
    source file character-for-character (the API rejects rows with stray
    whitespace), while the rest of the tool keeps using stripped names.
    """
    required = {"Code Description", "Type Code", "Type Description"}
    if required - set(prior_df.columns):
        return {}
    mask = prior_df["Code Description"].astype(str).str.strip() == code_description_value
    sub = prior_df[mask]
    counts = {}
    for _, r in sub.iterrows():
        raw_code = "" if pd.isna(r["Type Code"]) else str(r["Type Code"])
        raw_desc = "" if pd.isna(r["Type Description"]) else str(r["Type Description"])
        sc, sd = raw_code.strip(), raw_desc.strip()
        if sc.lower() in ("", "nan", "none"):
            continue
        key = autosync_row_key(sc, sd)
        counts.setdefault(key, {})
        counts[key][raw_desc] = counts[key].get(raw_desc, 0) + 1
    return {k: max(v.items(), key=lambda kv: kv[1])[0] for k, v in counts.items()}


def _mapping_csv_bytes(rows, columns):
    """rows: list of dicts keyed by `columns`. Returns UTF-8 CSV bytes with the
    exact header and no index. No value is trimmed or re-padded — what goes in is
    what comes out (the API rejects rows with stray whitespace)."""
    df = pd.DataFrame(rows, columns=columns) if rows else pd.DataFrame(columns=columns)
    return df.to_csv(index=False).encode("utf-8")


def build_earnings_mapping_rows(enriched_earnings, skipped_earnings, earn_raw=None):
    """One row per earning — BOTH the created ones (enriched) and the skipped
    UZIO defaults. Source name = RAW source name; Uzio name = the created
    Earning Name for kept rows, the UZIO default name for skipped rows."""
    earn_raw = earn_raw or {}
    rows = []
    for r in enriched_earnings or []:
        key = autosync_row_key(r.get("Type Code"), r.get("Type Description"))
        rows.append({
            "Source Earning Code": "",
            "Source Earning Code Name": earn_raw.get(key, r.get("Type Description", "")),
            "Uzio Earning Code": "",
            "Uzio Earning Code Name": r.get("Earning Name", r.get("Type Description", "")),
        })
    for r in skipped_earnings or []:
        key = autosync_row_key(r.get("Type Code"), r.get("Type Description"))
        rows.append({
            "Source Earning Code": "",
            "Source Earning Code Name": earn_raw.get(key, r.get("Type Description", "")),
            "Uzio Earning Code": "",
            "Uzio Earning Code Name": r.get("UZIO Default Earning", ""),
        })
    return rows


def build_deductions_mapping_rows(enriched_deds, ded_raw=None):
    """One row per created deduction. Source name = RAW source name; Uzio name =
    the exact name the script types into UZIO (UZIO Deduction Name)."""
    ded_raw = ded_raw or {}
    rows = []
    for r in enriched_deds or []:
        key = autosync_row_key(r.get("Type Code"), r.get("Type Description"))
        rows.append({
            "Source Deduction Code": "",
            "Source Deduction Code Name": ded_raw.get(key, r.get("Type Description", "")),
            "Uzio Deduction Code": "",
            "Uzio Deduction Code Name": r.get("UZIO Deduction Name", r.get("Type Description", "")),
        })
    return rows


def build_contributions_mapping_rows(enriched_contribs, ded_raw=None):
    """One row per created contribution. Source name = RAW source name (contribs
    are bifurcated from the Deductions rows, so they share ded_raw); Uzio name =
    the exact Contribution Name the script types into UZIO."""
    ded_raw = ded_raw or {}
    rows = []
    for r in enriched_contribs or []:
        key = autosync_row_key(r.get("Type Code"), r.get("Type Description"))
        rows.append({
            "Source Contribution Code": "",
            "Source Contribution Code Name": ded_raw.get(key, r.get("Type Description", "")),
            "Uzio Contribution Code": "",
            "Uzio Contribution Code Name": r.get("Contribution Name", r.get("Type Description", "")),
        })
    return rows


def build_mapping_files_zip(safe_name, earnings_rows, deductions_rows, contributions_rows):
    """Bundle the three mapping CSVs into a single .zip for one-click download.
    File names mirror the client templates: <Client>_Earnings_mapping.csv etc."""
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(f"{safe_name}_Earnings_mapping.csv",
                    _mapping_csv_bytes(earnings_rows, MAPPING_EARNING_COLUMNS))
        zf.writestr(f"{safe_name}_Deductions_mapping.csv",
                    _mapping_csv_bytes(deductions_rows, MAPPING_DEDUCTION_COLUMNS))
        zf.writestr(f"{safe_name}_Contributions_mapping.csv",
                    _mapping_csv_bytes(contributions_rows, MAPPING_CONTRIBUTION_COLUMNS))
    return buf.getvalue()


# HTML/JS for a SINGLE button that downloads all three mapping CSVs as separate
# files (st.download_button can only emit one file per click). Rendered via
# st.components.v1.html — its iframe sandbox includes `allow-downloads`, so the
# programmatic anchor clicks below actually download. `__FILES__` is replaced with
# a JSON array of [filename, base64-csv] pairs. The 400ms stagger lets the browser
# accept the multi-file download (Chrome prompts "allow multiple downloads" once).
MAPPING_DOWNLOAD_HTML = """
<style>
  .dlwrap { font-family: "Source Sans Pro", system-ui, -apple-system, sans-serif; }
  .dl-btn {
    background:#ff4b4b; color:#fff; border:1px solid #ff4b4b; border-radius:.5rem;
    padding:.55rem 1rem; font-size:1rem; font-weight:600; cursor:pointer;
  }
  .dl-btn:hover { background:#e53935; border-color:#e53935; }
  .dl-btn:active { transform:translateY(1px); }
  .dl-note { color:#808495; font-size:.8rem; margin-top:.5rem; }
</style>
<div class="dlwrap">
  <button class="dl-btn" id="dlAllMaps">&#128229; Download all mapping CSVs</button>
  <div class="dl-note">Saves the mapping files (earnings, deductions, contributions, taxes)
    as separate files. The first time, Chrome may ask to
    &ldquo;allow downloading multiple files&rdquo; &mdash; choose <b>Allow</b>.</div>
</div>
<script>
  (function () {
    var FILES = __FILES__;
    var btn = document.getElementById('dlAllMaps');
    btn.addEventListener('click', function () {
      FILES.forEach(function (f, i) {
        setTimeout(function () {
          var a = document.createElement('a');
          a.href = 'data:text/csv;charset=utf-8;base64,' + f[1];
          a.download = f[0];
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        }, i * 400);
      });
    });
  })();
</script>
"""


def filter_default_uzio_deductions(rows):
    """Split an extracted-deductions list into (kept, skipped) based on
    DEFAULT_UZIO_DEDUCTIONS_TO_SKIP. UZIO auto-creates a small set of
    deductions on every new client (currently just "Earned Wage Access");
    the automation must not re-create them.

    Match is case-insensitive on Type Description (whitespace stripped).
    Returns: (kept_rows, skipped_rows) — both same shape as input.
    """
    if not rows:
        return [], []
    kept, skipped = [], []
    for r in rows:
        td = (r.get("Type Description") or "").strip().lower()
        (skipped if td in DEFAULT_UZIO_DEDUCTIONS_TO_SKIP else kept).append(r)
    return kept, skipped


def enrich_deductions_for_uzio(rows, auto_sync_map=None, master_override_map=None):
    """Take the classified deduction rows (each with Type Code / Type Description /
    Calc Description / Pre/Post Tax) and add every UZIO "Add Deduction" form field
    the Tampermonkey automation needs, as columns.

    Columns added (see DEDUCTION_OUTPUT_COLUMNS for the canonical order):
      - UZIO Master Deductions List   the dropdown that drives the whole form
      - UZIO Deduction Type           mirrors Pre/Post Tax (UZIO auto-locks this)
      - UZIO Deduction Name           defaults to the Master value
      - UZIO Method                   Fixed $ / % of Gross Pay / % of Disposable Net Pay
      - Amount per pay                blank (per-employee; filled later / by automation)
      - Auto-Sync from Uzio Benefits  Yes/No for benefit types, else "N/A"
      - Assign to all employees       "Yes" (forced+locked Yes for the special masters)
      - Deduction Schedule            "Every Paycheck"
      - Track arrears                 "Yes" for benefit types, else "No"
      - Arrears Processing Method     "Total Amount" for benefit types, else ""
      - W-2 Box                       blank (optional in UZIO)
      - Is Benefit Type               bool — drives the Auto-Sync toggle UI (not exported)

    `auto_sync_map`: optional {autosync_row_key(code, desc) -> "Yes"/"No"} from the
    UI toggle bar. Only consulted for benefit-type rows; defaults to "No" if a
    benefit row isn't present in the map.

    Returns a fresh list of dicts (original keys preserved).
    """
    auto_sync_map = auto_sync_map or {}
    master_override_map = master_override_map or {}
    out = []
    for r in rows:
        master = map_paycom_to_uzio_master(
            r.get("Type Code"), r.get("Type Description"), r.get("Pre/Post Tax")
        )
        # Manual override from the UI for deductions the tool couldn't map.
        ovr = master_override_map.get(autosync_row_key(r.get("Type Code"), r.get("Type Description")))
        if ovr:
            master = ovr
        method = determine_method(master, r.get("Type Description"))
        benefit = is_benefit_type(master) and master != NEEDS_REVIEW

        if benefit:
            auto_sync = auto_sync_map.get(
                autosync_row_key(r.get("Type Code"), r.get("Type Description")), "No"
            )
            track_arrears = "Yes"
            arrears_method = ARREARS_PROCESSING_TOTAL
        else:
            auto_sync = AUTOSYNC_NA
            track_arrears = "No"
            arrears_method = ""

        assign_all = "Yes" if is_assign_all_locked(master) else DEFAULT_ASSIGN_TO_ALL

        # "Other" (and unmapped) rows: the master is generic, so the real name
        # lives in Deduction Name. UZIO also locks "Other" to Post Tax.
        if master in (MASTER_OTHER, NEEDS_REVIEW):
            ded_name = r.get("Type Description", "")
        else:
            ded_name = master
        ded_type = "Post Tax" if master == MASTER_OTHER else r.get("Pre/Post Tax", "")

        # W-2 Box: real masters auto-fill + lock it; Other/unmapped are editable.
        w2_box = DEFAULT_W2_BOX if master in (MASTER_OTHER, NEEDS_REVIEW) else W2_BOX_LOCKED

        out.append({
            **r,
            "UZIO Master Deductions List": master,
            "UZIO Deduction Type": ded_type,
            "UZIO Deduction Name": ded_name,
            "UZIO Method": method,
            "Amount per pay": "",
            "Auto-Sync from Uzio Benefits": auto_sync,
            "Assign to all employees": assign_all,
            "Deduction Schedule": DEFAULT_DEDUCTION_SCHEDULE,
            "Track arrears": track_arrears,
            "Arrears Processing Method": arrears_method,
            "W-2 Box": w2_box,
            "Is Benefit Type": benefit,
        })
    return out


def bifurcate_match_memo(rows):
    """Split a list of (Type Code, Type Description) dicts into
    (contributions, deductions) using the Match/Memo rule.

    - "Match" anywhere in Type Code or Type Description (whole-word,
      case-insensitive) -> Contribution (employer-side matching dollars).
    - "Memo" anywhere in Type Code or Type Description -> Contribution
      (informational/tracking-only line, employer-side).
    - Everything else -> Deduction (employee-paid).

    The word-boundary check stops false positives like "Rematch" or
    "Memorial" from being miscategorized.
    """
    contribs, deds = [], []
    for r in rows:
        combined = f"{r.get('Type Code', '')} {r.get('Type Description', '')}"
        if MATCH_MEMO_RE.search(combined):
            contribs.append(r)
        else:
            deds.append(r)
    return contribs, deds


def build_earnings_catalog(prior_df):
    if "Code Description" not in prior_df.columns:
        return []
    earn = prior_df[prior_df["Code Description"].astype(str).str.strip() == "Earnings"]
    rows, seen = [], set()
    for _, r in earn.iterrows():
        tc = str(r.get("Type Code", "")).strip()
        td = str(r.get("Type Description", "")).strip()
        key = (tc, td)
        if not tc or key in seen:
            continue
        seen.add(key)
        amt = earn[(earn["Type Code"] == tc) & (earn["Type Description"] == td)]["Amount"].apply(_num)
        rows.append({
            "Type Code": tc, "Type Description": td,
            "Total $": round(float(amt.sum()), 2),
            "Employees": int(len(amt[amt != 0])),
        })
    return rows


def build_taxes_discovered(prior_df):
    if "Code Description" not in prior_df.columns:
        return []
    tax = prior_df[prior_df["Code Description"].astype(str).str.strip() == "W/H Taxes"]
    rows, seen = [], set()
    for _, r in tax.iterrows():
        tc = str(r.get("Type Code", "")).strip()
        td = str(r.get("Type Description", "")).strip()
        key = (tc, td)
        if not tc or key in seen:
            continue
        seen.add(key)
        amt = tax[(tax["Type Code"] == tc) & (tax["Type Description"] == td)]["Amount"].apply(_num)
        rows.append({
            "Type Code": tc, "Type Description": td,
            "Total $": round(float(amt.sum()), 2),
            "Employees": int(len(amt[amt != 0])),
        })
    return rows


def split_contribs_deductions(scheduled_df):
    if "Deduction Code" not in scheduled_df.columns:
        return [], []
    rows, seen = [], set()
    for _, r in scheduled_df.iterrows():
        dc = str(r.get("Deduction Code", "")).strip()
        dd = str(r.get("Deduction Desc", "")).strip()
        key = (dc, dd)
        if not dc or key in seen:
            continue
        seen.add(key)
        rows.append({
            "Deduction Code": dc, "Deduction Desc": dd,
            "Setup Count": int(((scheduled_df["Deduction Code"] == dc)
                                & (scheduled_df["Deduction Desc"] == dd)).sum()),
        })
    contribs, deds = [], []
    for r in rows:
        u = (r["Deduction Code"] + " " + r["Deduction Desc"]).upper()
        (contribs if CONTRIB_PATTERN.search(u) else deds).append(r)
    return contribs, deds


def classify_pre_post_tax(scheduled_df):
    if "Deduction Code" not in scheduled_df.columns or "Tax Treatment" not in scheduled_df.columns:
        return []
    rows = []
    grouped = scheduled_df.groupby(["Deduction Code", "Deduction Desc"], dropna=False)
    for (dc, dd), grp in grouped:
        treatments = grp["Tax Treatment"].dropna().astype(str).str.strip().unique().tolist()
        if not treatments:
            verdict, flavor, why = "unknown", "", "Tax Treatment column was blank for every row of this deduction."
        else:
            primary = grp["Tax Treatment"].dropna().astype(str).str.strip().mode()
            tt = primary.iloc[0] if not primary.empty else treatments[0]
            tt_upper = tt.upper()
            if tt_upper.startswith("B"):
                verdict, flavor = "PRE-TAX", "Section 125"
                why = f"Tax Treatment '{tt}' = Section 125 cafeteria plan (reduces FIT, FICA, Medicare, and state-income taxable wages)."
            elif tt_upper.startswith("H"):
                verdict, flavor = "PRE-TAX", "401k traditional"
                why = f"Tax Treatment '{tt}' = traditional 401(k) (reduces FIT and SIT but NOT FICA/Medicare)."
            elif tt_upper.startswith("A"):
                verdict, flavor = "POST-TAX", ""
                why = f"Tax Treatment '{tt}' = post-tax deduction (does not reduce taxable wages)."
            else:
                verdict, flavor = "unknown", "review"
                why = f"Tax Treatment '{tt}' is not a recognized Paycom code -- please review manually."
            if len(treatments) > 1:
                why += f"  (Multiple distinct Tax Treatments seen: {treatments}; using the most common.)"
        rows.append({
            "Code": str(dc).strip(),
            "Description": str(dd).strip() if dd is not None else "",
            "Verdict": verdict, "Flavor": flavor, "Why": why,
        })
    return rows


def classify_bonus(prior_df):
    if "Code Description" not in prior_df.columns or "Type Code" not in prior_df.columns:
        return {"verdict": "indeterminate",
                "reason": "Prior Payroll Register is missing Code Description / Type Code columns.",
                "bonus_codes_found": [], "samples": []}
    earn = prior_df[prior_df["Code Description"].astype(str).str.strip() == "Earnings"]
    bonus_codes = sorted({
        str(r["Type Code"]).strip() for _, r in earn.iterrows()
        if BONUS_RE.search(f"{r.get('Type Code', '')} {r.get('Type Description', '')}".upper())
    })
    ot_codes = ["OT", "OVT", "OVR"]
    wot_codes = ["WOT"]
    has_ot = any(c in earn["Type Code"].astype(str).unique() for c in ot_codes)
    has_wot = any(c in earn["Type Code"].astype(str).unique() for c in wot_codes)

    if not bonus_codes:
        return {"verdict": "no_bonus_in_file",
                "reason": ("No bonus codes found in the Prior Payroll Register. "
                           "(Looked for Type Codes containing BONUS / BNS / BND / BNH / BN# / NA#.) "
                           "If a bonus exists outside this pay period, supply that file too."),
                "bonus_codes_found": [], "ot_present": has_ot, "wot_present": has_wot,
                "samples": []}

    if not (has_ot and has_wot):
        msg = []
        if has_wot and not has_ot:
            msg.append("File contains only Paycom's WOT (weighted overtime) lines; "
                       "the plain-OT comparison line is absent so the WOT-vs-OT differential "
                       "test cannot run.")
        elif has_ot and not has_wot:
            msg.append("File contains only plain-OT lines; the WOT (weighted overtime) "
                       "comparison is absent.")
        else:
            msg.append("File contains neither OT nor WOT lines.")
        msg.append("To classify the bonus, supply a Paycom Payroll Register Detail report "
                   "with hours, OR confirm the bonus type with the implementer directly.")
        return {"verdict": "indeterminate", "reason": " ".join(msg),
                "bonus_codes_found": bonus_codes, "ot_present": has_ot,
                "wot_present": has_wot, "samples": []}

    pivot = earn.pivot_table(
        index="EE Code", columns="Type Code", values="Amount",
        aggfunc=lambda s: float(sum(_num(v) for v in s)), fill_value=0.0,
    )
    samples = []
    differential_rows = matching_rows = 0
    rate_tol_pct = 0.005
    for eid, row in pivot.iterrows():
        ot_amt = sum(_num(row[c]) for c in ot_codes if c in row.index)
        wot_amt = sum(_num(row[c]) for c in wot_codes if c in row.index)
        bonus_amt = sum(_num(row[c]) for c in bonus_codes if c in row.index)
        if ot_amt <= 0 or wot_amt <= 0 or bonus_amt <= 0:
            continue
        diff_pct = (wot_amt - ot_amt) / ot_amt if ot_amt > 0 else 0.0
        if diff_pct > rate_tol_pct:
            differential_rows += 1
        else:
            matching_rows += 1
        if len(samples) < 5:
            samples.append({
                "employee": str(eid),
                "plain_ot_amount": round(ot_amt, 2),
                "weighted_ot_amount": round(wot_amt, 2),
                "differential_pct": round(diff_pct * 100, 3),
                "bonus_amount": round(bonus_amt, 2),
                "row_verdict": "non_discretionary" if diff_pct > rate_tol_pct else "discretionary",
            })
    rows_tested = differential_rows + matching_rows
    if rows_tested == 0:
        return {"verdict": "indeterminate",
                "reason": ("Bonus codes were found but no employee in this pay period had both "
                           "OT, WOT, and a bonus amount in the same row."),
                "bonus_codes_found": bonus_codes, "samples": []}
    if differential_rows > 0:
        return {"verdict": "non_discretionary",
                "reason": (f"{differential_rows} of {rows_tested} employees show Paycom's WOT "
                           f"materially higher than plain OT. Paycom rolls non-discretionary "
                           f"bonuses into the regular rate before computing weighted OT, so the "
                           f"gap means the bonus is non-discretionary under FLSA."),
                "bonus_codes_found": bonus_codes,
                "rows_tested": rows_tested, "differential_rows": differential_rows,
                "matching_rows": matching_rows, "samples": samples}
    return {"verdict": "discretionary",
            "reason": (f"All {rows_tested} tested employees show WOT == plain OT (no weighted "
                       f"adjustment). Paycom did NOT roll the bonus into the regular rate, so "
                       f"the bonus is discretionary."),
            "bonus_codes_found": bonus_codes,
            "rows_tested": rows_tested, "differential_rows": 0,
            "matching_rows": matching_rows, "samples": samples}


def _pick_bonus_example(bonus_info):
    samples = bonus_info.get("samples", [])
    if not samples:
        return None
    if bonus_info["verdict"] == "non_discretionary":
        cands = [s for s in samples if s["row_verdict"] == "non_discretionary"]
        return max(cands, key=lambda s: s["differential_pct"]) if cands else samples[0]
    if bonus_info["verdict"] == "discretionary":
        cands = [s for s in samples if s["row_verdict"] == "discretionary"]
        return min(cands, key=lambda s: abs(s["differential_pct"])) if cands else samples[0]
    return samples[0]


def build_simplified_xlsx_bytes(results):
    """Three-tab xlsx output matching the ADP setup helper format."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({"bold": True, "bg_color": "#1F4E78",
                                    "font_color": "white", "border": 1,
                                    "align": "left", "valign": "vcenter"})
        wrap_fmt = wb.add_format({"valign": "top", "text_wrap": True})
        v_pre = wb.add_format({"bold": True, "bg_color": "#C6EFCE",
                               "font_color": "#006100", "align": "center", "valign": "vcenter"})
        v_post = wb.add_format({"bold": True, "bg_color": "#FFC7CE",
                                "font_color": "#9C0006", "align": "center", "valign": "vcenter"})
        v_nondisc = wb.add_format({"bold": True, "bg_color": "#FFC7CE",
                                   "font_color": "#9C0006", "align": "left",
                                   "valign": "vcenter", "font_size": 14})
        v_disc = wb.add_format({"bold": True, "bg_color": "#C6EFCE",
                                "font_color": "#006100", "align": "left",
                                "valign": "vcenter", "font_size": 14})

        # Tab 1
        earn = [r["Type Code"] + " - " + r["Type Description"] for r in results["Earnings_Codes"]]
        contrib = [r["Deduction Code"] + " - " + r["Deduction Desc"] for r in results["Contributions"]]
        ded = [r["Deduction Code"] + " - " + r["Deduction Desc"] for r in results["Deductions"]]
        max_n = max(len(earn), len(contrib), len(ded), 1)
        df1 = pd.DataFrame([{
            "Earnings": earn[i] if i < len(earn) else "",
            "Contributions": contrib[i] if i < len(contrib) else "",
            "Deductions": ded[i] if i < len(ded) else "",
        } for i in range(max_n)])
        df1.to_excel(writer, sheet_name="1. What to Set Up", index=False)
        ws1 = writer.sheets["1. What to Set Up"]
        ws1.set_column("A:A", 38); ws1.set_column("B:B", 32); ws1.set_column("C:C", 38)
        for i, c in enumerate(df1.columns):
            ws1.write(0, i, c, header_fmt)
        ws1.set_row(0, 24)

        # Tab 2
        rows2 = [{"Code": r["Code"], "Description": r["Description"], "Verdict": r["Verdict"],
                  "Flavor": r["Flavor"], "Why": r["Why"]} for r in results["Pre_Post_Tax"]]
        if not rows2:
            rows2 = [{"Code": "(none)", "Description": "", "Verdict": "", "Flavor": "",
                      "Why": "Scheduled Deductions report had no rows."}]
        df2 = pd.DataFrame(rows2)
        df2.to_excel(writer, sheet_name="2. Pre-tax vs Post Tax", index=False)
        ws2 = writer.sheets["2. Pre-tax vs Post Tax"]
        ws2.set_column("A:A", 14); ws2.set_column("B:B", 30)
        ws2.set_column("C:C", 11); ws2.set_column("D:D", 20)
        ws2.set_column("E:E", 90, wrap_fmt)
        for i, c in enumerate(df2.columns):
            ws2.write(0, i, c, header_fmt)
        ws2.set_row(0, 24)
        for ri, r in enumerate(rows2, start=1):
            v = r["Verdict"]
            if v == "PRE-TAX":
                ws2.write(ri, 2, "PRE-TAX", v_pre)
            elif v == "POST-TAX":
                ws2.write(ri, 2, "POST-TAX", v_post)
            ws2.set_row(ri, 30)

        # Tab 3
        bonus = results["Bonus"]
        sample = _pick_bonus_example(bonus)
        verdict_label = bonus["verdict"].upper().replace("_", "-")
        rows3 = [
            ("Verdict", verdict_label),
            ("Reason", bonus["reason"]),
            ("Bonus codes detected", ", ".join(bonus.get("bonus_codes_found", [])) or "(none)"),
        ]
        if "rows_tested" in bonus:
            rows3 += [
                ("Employees tested", bonus.get("rows_tested", 0)),
                ("    of which non-discretionary (WOT > OT)", bonus.get("differential_rows", 0)),
                ("    of which discretionary (WOT == OT)", bonus.get("matching_rows", 0)),
            ]
        if sample:
            rows3 += [
                ("", ""),
                ("---- Example employee that proves the verdict ----", ""),
                ("Employee", sample["employee"]),
                ("Plain OT amount (Paycom 'OT')", f"${sample['plain_ot_amount']:,}"),
                ("Weighted OT amount (Paycom 'WOT', FLSA-corrected)", f"${sample['weighted_ot_amount']:,}"),
                ("Differential (%)", f"{sample['differential_pct']}%"),
                ("Bonus amount in this period", f"${sample['bonus_amount']:,}"),
                ("", ""),
                ("Plain-English explanation",
                    "WOT > OT => Paycom rolled the bonus into the regular rate before "
                    "calculating the weighted OT. Per FLSA, that means the bonus is "
                    "NON-DISCRETIONARY."
                    if bonus["verdict"] == "non_discretionary" else
                    "WOT matches plain OT exactly => Paycom did NOT roll the bonus into the "
                    "regular rate => bonus is DISCRETIONARY."
                    if bonus["verdict"] == "discretionary" else
                    bonus["reason"]),
            ]
        df3 = pd.DataFrame(rows3, columns=["Field", "Value"])
        df3.to_excel(writer, sheet_name="3. Bonus Verdict", index=False)
        ws3 = writer.sheets["3. Bonus Verdict"]
        ws3.set_column("A:A", 50); ws3.set_column("B:B", 80, wrap_fmt)
        for i, c in enumerate(df3.columns):
            ws3.write(0, i, c, header_fmt)
        ws3.set_row(0, 24)
        if bonus["verdict"] == "non_discretionary":
            ws3.write(1, 1, verdict_label, v_nondisc)
        elif bonus["verdict"] == "discretionary":
            ws3.write(1, 1, verdict_label, v_disc)
        ws3.set_row(1, 28)

    return buf.getvalue()


def _deduction_reason_short(verdict, flavor):
    if verdict == "PRE-TAX" and flavor == "Section 125":
        return "Reduces FIT, FICA, Medicare, and state-income taxable wages -- Section 125 cafeteria plan."
    if verdict == "PRE-TAX" and flavor == "401k traditional":
        return "Reduces FIT and state-income taxable wages but NOT FICA/Medicare -- traditional 401(k)/403(b)."
    if verdict == "POST-TAX":
        return "Does not reduce taxable wages."
    return "Tax Treatment value not recognized -- review manually."


# ─────────────────────────────────────────────────────────────────────────────
# Taxes: Paycom → UZIO tax-catalog mapping
#
# Paycom delivers taxes in two prior-payroll sections:
#   - "W/H Taxes"               → EMPLOYEE-side withholding (FIT, Medicare, SS,
#                                 State Income, local city/school)
#   - "Client Side Liabilities" → EMPLOYER-side taxes (FUTA, State Unemployment,
#                                 Employer Medicare/Social Security). This section
#                                 also holds employer benefit contributions, so we
#                                 keep only the tax rows (keyword filter).
# The SECTION is part of each tax's identity: the same code+name (e.g. MED
# "Medicare") appears in both sections and maps to a DIFFERENT UZIO tax
# (Medicare Tax vs Employer Medicare Tax).
#
# UZIO side = the bundled tax catalog (uzio_tax_catalog.csv): state_abbreviation,
# tax_code (TAX####), unique_tax_id, tax_name, sub_tax_desc. The 4th dash-segment
# of unique_tax_id is the tax-TYPE token (FIT / MEDI / FICA / ER_FUTA / ER_MEDI /
# ER_FICA / SIT / ER_SUTA / CITY / SCHL / ...), which is the precise match key.
#
# Mapping is NOT a creation step — UZIO already owns these taxes; we only produce
# the mapping file (Source tax → UZIO tax_code/unique_tax_id). Tiers:
#   federal  (deterministic)  state (deterministic)  local (best-guess + confirm).
# ─────────────────────────────────────────────────────────────────────────────

TAX_SECTION_WH = "W/H Taxes"
TAX_SECTION_EMPLOYER = "Client Side Liabilities"
TAX_SECTIONS = (TAX_SECTION_WH, TAX_SECTION_EMPLOYER)

# Output columns — verbatim from the client tax-mapping template.
MAPPING_TAX_COLUMNS = [
    "Source Tax Code", "Source Tax Code Name", "Source Tax Code Description",
    "Uzio Tax Code", "Unique Tax ID", "Uzio Tax Code Description",
    "Uzio Sub-Tax Description",
]

UZIO_TAX_CATALOG_PATH = os.path.join(os.path.dirname(__file__), "uzio_tax_catalog.csv")

STATE_NAME_TO_ABBREV = {
    "alabama": "AL", "alaska": "AK", "arizona": "AZ", "arkansas": "AR",
    "california": "CA", "colorado": "CO", "connecticut": "CT", "delaware": "DE",
    "district of columbia": "DC", "florida": "FL", "georgia": "GA", "hawaii": "HI",
    "idaho": "ID", "illinois": "IL", "indiana": "IN", "iowa": "IA", "kansas": "KS",
    "kentucky": "KY", "louisiana": "LA", "maine": "ME", "maryland": "MD",
    "massachusetts": "MA", "michigan": "MI", "minnesota": "MN", "mississippi": "MS",
    "missouri": "MO", "montana": "MT", "nebraska": "NE", "nevada": "NV",
    "new hampshire": "NH", "new jersey": "NJ", "new mexico": "NM", "new york": "NY",
    "north carolina": "NC", "north dakota": "ND", "ohio": "OH", "oklahoma": "OK",
    "oregon": "OR", "pennsylvania": "PA", "rhode island": "RI",
    "south carolina": "SC", "south dakota": "SD", "tennessee": "TN", "texas": "TX",
    "utah": "UT", "vermont": "VT", "virginia": "VA", "washington": "WA",
    "west virginia": "WV", "wisconsin": "WI", "wyoming": "WY",
}
US_STATE_ABBREVS = set(STATE_NAME_TO_ABBREV.values())

# Abbreviated / nickname state spellings Paycom uses, e.g. "N Carolina W/H" → NC,
# "Penn. State W/H" → PA. Checked (word-boundary) after full names, before the
# 2-letter fallback. Word boundaries keep short ones (ind/ill/del) from matching
# inside other words.
STATE_NAME_VARIANTS = {
    # directional
    "n carolina": "NC", "no carolina": "NC", "s carolina": "SC", "so carolina": "SC",
    "n dakota": "ND", "no dakota": "ND", "s dakota": "SD", "so dakota": "SD",
    "w virginia": "WV", "n hampshire": "NH", "n jersey": "NJ", "n mexico": "NM",
    "n york": "NY", "r island": "RI",
    # common Paycom abbreviations / nicknames
    "penn": "PA", "penna": "PA", "calif": "CA", "conn": "CT", "mass": "MA",
    "mich": "MI", "minn": "MN", "wisc": "WI", "wis": "WI", "tenn": "TN",
    "fla": "FL", "ariz": "AZ", "colo": "CO", "okla": "OK", "oreg": "OR",
    "wyo": "WY", "nev": "NV", "nebr": "NE", "neb": "NE", "mont": "MT",
    "kans": "KS", "ark": "AR", "ala": "AL", "miss": "MS", "ind": "IN",
    "ill": "IL", "del": "DE", "wash": "WA", "ga": "GA", "vt": "VT",
}


def load_uzio_tax_catalog(path=UZIO_TAX_CATALOG_PATH):
    """Read the bundled UZIO tax catalog into a list of dicts. Strings only,
    no NaN. Returns [] if the file is missing/unreadable."""
    try:
        df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except Exception:
        return []
    out = []
    for _, r in df.iterrows():
        out.append({
            "state_abbreviation": str(r.get("state_abbreviation", "")).strip(),
            "tax_code": str(r.get("tax_code", "")).strip(),
            "unique_tax_id": str(r.get("unique_tax_id", "")).strip(),
            "tax_name": str(r.get("tax_name", "")).strip(),
            "sub_tax_desc": str(r.get("sub_tax_desc", "")).strip(),
        })
    return out


def _uti_token(unique_tax_id):
    """The tax-TYPE token = 4th dash-segment of the unique_tax_id."""
    parts = (unique_tax_id or "").split("-")
    return parts[3].upper() if len(parts) >= 4 else ""


def _is_statewide(r):
    """True if a catalog row is the statewide entry (place segment == '0000') —
    i.e. the state-level tax, not a city/school within the state."""
    p = (r.get("unique_tax_id") or "").split("-")
    return len(p) >= 3 and p[2] == "0000"


# Client Side Liabilities mixes employer taxes with employer contributions; keep
# only rows that look like a tax.
_EMPLOYER_TAX_KW = (
    "futa", "suta", "sui", "unemploy", "medicare", "social security", "fica",
    "oasdi", "disability", "sdi", "state w/h", "state withhold", "income tax",
    "local", "lst", "eit",
)


def _is_employer_tax_row(name):
    d = (name or "").lower()
    return any(k in d for k in _EMPLOYER_TAX_KW)


def tax_row_key(section, code, desc):
    """Unique key per tax row — MUST include section (same code+name repeats across
    W/H Taxes vs Client Side Liabilities and maps to different UZIO taxes)."""
    return f"{(section or '').strip()}||{(code or '').strip()}||{(desc or '').strip()}"


def extract_unique_taxes_from_prior(prior_df):
    """Unique taxes from W/H Taxes (all) + Client Side Liabilities (tax rows only).
    Each dict: Type Code, Type Description, Section. Section order preserved
    (employee first, then employer)."""
    required = {"Code Description", "Type Code", "Type Description"}
    if required - set(prior_df.columns):
        return []
    out, seen = [], set()
    for section in TAX_SECTIONS:
        sub = prior_df[prior_df["Code Description"].astype(str).str.strip() == section]
        for _, r in sub.iterrows():
            tc = str(r.get("Type Code", "")).strip()
            td = str(r.get("Type Description", "")).strip()
            if tc.lower() in ("", "nan", "none"):
                continue
            if _is_ignored_paycom_item(tc, td):   # drop Worker's Comp (WKC) everywhere
                continue
            if section == TAX_SECTION_EMPLOYER and not _is_employer_tax_row(td):
                continue
            key = (section, tc, td)
            if key in seen:
                continue
            seen.add(key)
            out.append({"Type Code": tc, "Type Description": td, "Section": section})
    return out


def _detect_state_abbrev(name):
    """Pull a US state from a tax name: full name first, then a 2-letter token."""
    d = (name or "").lower()
    # Match the LONGEST full-name/variant phrase first, so "West Virginia" and
    # "W Virginia" win over the sub-word "Virginia" (and likewise for any future
    # containment like that).
    phrases = list(STATE_NAME_VARIANTS.items()) + list(STATE_NAME_TO_ABBREV.items())
    for phrase, ab in sorted(phrases, key=lambda kv: -len(kv[0])):
        if re.search(r"\b" + re.escape(phrase) + r"\b", d):
            return ab
    for tok in re.findall(r"\b[A-Za-z]{2}\b", name or ""):
        u = tok.upper()
        if u == "SD":  # "SD" in a tax name means School District, not South Dakota
            continue
        if u in US_STATE_ABBREVS:
            return u
    return None


def _is_employer_side(name, section):
    """Employee vs employer side. An explicit EE/ER (or Employee/Employer) in the
    NAME wins (e.g. 'New Jersey EE SUI' = employee); otherwise the Paycom section
    decides (Client Side Liabilities = employer)."""
    d = (name or "").lower()
    if re.search(r"\bemployer\b", d) or re.search(r"\ber\b", d):
        return True
    if re.search(r"\bemployee\b", d) or re.search(r"\bee\b", d):
        return False
    return (section or "").strip() == TAX_SECTION_EMPLOYER


def _federal_token(name, employer):
    """UZIO federal TYPE token for a Paycom federal tax (None if not federal)."""
    d = (name or "").lower()
    if "futa" in d or ("federal" in d and "unemploy" in d):
        return "ER_FUTA"                      # FUTA is employer-only
    if "additional medicare" in d:
        return "MEDI2"
    if "medicare" in d:
        return "ER_MEDI" if employer else "MEDI"
    if "social security" in d or "fica" in d or "oasdi" in d:
        return "ER_FICA" if employer else "FICA"
    if "federal" in d and ("withhold" in d or "income" in d or "fit" in d.split()):
        return "FIT"
    return None


def _state_token(name, employer):
    """UZIO state-level TYPE token, or None. Employer-aware so EE vs ER taxes map
    to the correct side. Catalog tokens: SIT, SUI (employee unemployment) /
    ER_SUTA (employer unemployment), SDI (employee) / ER_SDI (employer), FLI."""
    d = (name or "").lower()
    if "family leave" in d or "fli" in d.split():
        return "FLI"
    if "sdi" in d or "disability" in d:
        return "ER_SDI" if employer else "SDI"
    if "suta" in d or "sui" in d or "unemploy" in d:
        # Employee unemployment is the catalog token "SUI"; employer is "ER_SUTA".
        return "ER_SUTA" if employer else "SUI"
    # State withholding = state income tax. Paycom often omits the word "state"
    # (e.g. "N Carolina W/H"), so "w/h"/"withhold"/"income" is enough.
    if "w/h" in d or "withhold" in d or "income" in d or "sit" in d.split():
        return "SIT"
    return None


_LOCAL_SCHOOL_KW = ("lsd", "csd", "school", "district")


def _local_kind(name):
    """City vs School vs JEDD/JEDZ from a Paycom local tax name."""
    d = (name or "").lower()
    if "jedd" in d:
        return "JEDD"
    if "jedz" in d:
        return "JEDZ"
    if any(k in d for k in _LOCAL_SCHOOL_KW) or re.search(r"\bsd\b", d):
        return "SCHL"
    if "local" in d or "city" in d or "lst" in d or "eit" in d:
        return "CITY"
    return None


def _clean_place(name, state_ab):
    """Reduce a Paycom local name to its place tokens (drop state + kind words)."""
    s = (name or "").lower()
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    for full, ab in STATE_NAME_TO_ABBREV.items():
        if ab == state_ab:
            s = re.sub(r"\b" + re.escape(full) + r"\b", " ", s)
    drop = {
        (state_ab or "").lower(), "local", "city", "lsd", "csd", "sd", "school",
        "district", "jedd", "jedz", "tax", "lst", "eit", "township", "borough",
        "of", "the",
    }
    # Drop the leading numeric jurisdiction code (e.g. "510101") and de-duplicate
    # tokens (Paycom repeats names like "PHILADELPHIA CITY/PHILADELPHIA CITY").
    toks = [t for t in s.split() if t and t not in drop and not t.isdigit()]
    seen = set()
    return " ".join(t for t in toks if not (t in seen or seen.add(t)))


def _catalog_place(tax_name, state_ab):
    """Reduce a UZIO catalog tax_name to its place tokens for similarity scoring."""
    s = (tax_name or "").lower()
    s = re.sub(r"[^a-z0-9 ]", " ", s)
    if state_ab:
        s = re.sub(r"^\s*" + re.escape(state_ab.lower()) + r"\b", " ", s)
    s = re.sub(r"\b(city|lsd|csd|sd|school|district|tax|jedd|jedz|income|state|of|the)\b", " ", s)
    return " ".join(s.split())


def dominant_tax_state(taxes):
    """Most-common state among the non-federal taxes — used as the default state
    for locals whose name omits the state (e.g. 'Evergreen SD')."""
    from collections import Counter
    c = Counter()
    for t in taxes:
        if _federal_token(t.get("Type Description"), False):
            continue
        st = _detect_state_abbrev(t.get("Type Description"))
        if st:
            c[st] += 1
    return c.most_common(1)[0][0] if c else None


def map_tax_to_uzio(code, name, section, catalog, file_default_state=None):
    """Map one Paycom tax to the UZIO catalog.

    Returns {"tier", "match", "candidates", "state", "kind"} where:
      - match      = best catalog-row dict (or None)
      - candidates = ranked catalog rows (locals) / exact matches (fed/state)
      - tier       = "federal" | "state" | "local" | "unknown"
    """
    employer = _is_employer_side(name, section)

    # 1) Federal — deterministic.
    ftok = _federal_token(name, employer)
    if ftok:
        hits = [r for r in catalog
                if r["state_abbreviation"].upper() == "FED" and _uti_token(r["unique_tax_id"]) == ftok]
        return {"tier": "federal", "match": hits[0] if hits else None,
                "candidates": hits, "state": "FED", "kind": ftok, "detected": None}

    # 2) State-level (income / unemployment / disability) — but NOT when this is
    #    clearly a local (city/school) tax. Candidates span EVERY state so the
    #    dropdown is a real fallback when our state guess is off; we pre-select the
    #    detected state's statewide row (and only when the state was actually
    #    detected from the name — never a silent wrong guess).
    detected = _detect_state_abbrev(name)
    st = detected or file_default_state
    kind = _local_kind(name)
    stok = _state_token(name, employer)
    if stok and not kind:
        typed = [r for r in catalog if _uti_token(r["unique_tax_id"]) == stok]
        typed.sort(key=lambda r: (not _is_statewide(r),
                                  r["state_abbreviation"].upper() != (detected or ""),
                                  r["state_abbreviation"]))
        match = None
        if detected:
            match = next((r for r in typed
                          if r["state_abbreviation"].upper() == detected and _is_statewide(r)), None)
        if typed:
            return {"tier": "state", "match": match, "candidates": typed,
                    "state": detected, "kind": stok, "detected": detected}

    # 3) Local (city / school district) — best-guess; user confirms.
    if st:
        kind = _local_kind(name)
        place = _clean_place(name, st)
        pool = [r for r in catalog if r["state_abbreviation"].upper() == st]

        def _kind_ok(r):
            t = _uti_token(r["unique_tax_id"])
            if kind == "SCHL":
                return t == "SCHL"
            if kind == "CITY":
                return t == "CITY"
            if kind in ("JEDD", "JEDZ"):
                return t == kind
            return True

        kinded = [r for r in pool if _kind_ok(r)] or pool

        def _score(r):
            return difflib.SequenceMatcher(None, place, _catalog_place(r["tax_name"], st)).ratio()

        ranked = sorted(kinded, key=_score, reverse=True)
        # Looser threshold when the state came from the NAME; stricter when the
        # state is only a dominant-state fallback — otherwise we'd force a garbage
        # match in the wrong state (e.g. "Philadelphia" → an NJ tax at 0.35).
        threshold = 0.34 if detected else 0.6
        best = ranked[0] if (ranked and place and _score(ranked[0]) >= threshold) else None
        return {"tier": "local", "match": best, "candidates": ranked,
                "state": st, "kind": kind, "detected": detected}

    return {"tier": "unknown", "match": None, "candidates": [],
            "state": None, "kind": None, "detected": detected}


def build_taxes_mapping_rows(taxes, resolved):
    """taxes: extracted list (Type Code/Description/Section).
    resolved: {tax_row_key -> chosen catalog-row dict or None}.
    Emits rows in MAPPING_TAX_COLUMNS order (verbatim source names)."""
    rows = []
    for t in taxes:
        m = resolved.get(tax_row_key(t["Section"], t["Type Code"], t["Type Description"]))
        rows.append({
            "Source Tax Code": t["Type Code"],
            "Source Tax Code Name": t["Type Description"],
            "Source Tax Code Description": t["Section"],
            "Uzio Tax Code": (m or {}).get("tax_code", ""),
            "Unique Tax ID": (m or {}).get("unique_tax_id", ""),
            "Uzio Tax Code Description": (m or {}).get("tax_name", ""),
            "Uzio Sub-Tax Description": (m or {}).get("sub_tax_desc", ""),
        })
    return rows


# ---------- Streamlit UI ----------

def _render_name_editor(title, rows, name_field, key_prefix, caption=None):
    """Collapsible accordion of editable UZIO-name text boxes — one per row.

    Writes the user's chosen name back into each `row[name_field]` IN PLACE, so a
    single edit here flows into every downstream consumer that reads the enriched
    rows: the on-screen setup table, the setup Excel, AND the API mapping file.

    `rows`        already-filtered list of enriched dicts (each has Type Code /
                  Type Description / `name_field`). Pass only creatable rows.
    `name_field`  the dict key holding the UZIO display name
                  ("Earning Name" / "UZIO Deduction Name" / "Contribution Name").
    `key_prefix`  unique widget-key namespace (e.g. "ppsh_earnname").

    Smart default: each box pre-fills with the row's computed name and keeps
    FOLLOWING it on reruns (e.g. when a deduction's Master is remapped) until the
    user types a custom value — after which the user's value sticks even if the
    computed default later changes.
    """
    if not rows:
        return
    with st.expander(title, expanded=False):
        if caption:
            st.caption(caption)
        for r in rows:
            code = r.get("Type Code", "")
            td = r.get("Type Description", "")
            default = str(r.get(name_field, td) or td)
            wkey = f"{key_prefix}::{autosync_row_key(code, td)}"
            defkey = wkey + "::__def__"
            prev_def = st.session_state.get(defkey)
            # Re-seed the box with the new default ONLY if the user hasn't
            # customized it (still equal to the previously-seeded default, or
            # never seeded). Must run before st.text_input is instantiated.
            if wkey not in st.session_state or st.session_state.get(wkey) == prev_def:
                st.session_state[wkey] = default
            st.session_state[defkey] = default
            nm_col, ed_col = st.columns([1, 1])
            with nm_col:
                st.markdown(f"**{code}** — {td}")
            with ed_col:
                st.text_input(
                    f"{key_prefix}-{wkey}", key=wkey, label_visibility="collapsed",
                )
            val = st.session_state.get(wkey, default)
            # Never let an empty box null out the name — fall back to the default.
            r[name_field] = val.strip() if isinstance(val, str) and val.strip() else default


@st.cache_data(show_spinner=False)
def _cached_tax_catalog(path=UZIO_TAX_CATALOG_PATH):
    """Load the bundled UZIO tax catalog once per session (cached)."""
    return load_uzio_tax_catalog(path)


def _tax_autoselect_state(state_key, tax_key, lblmap_key):
    """on_change for a tax dropdown: when a tax is picked, set its State picker to
    that tax's state. Lets the 'Search all states' mode auto-fill the State field
    once you choose a result."""
    label = st.session_state.get(tax_key)
    state_for_label = (st.session_state.get(lblmap_key) or {}).get(label)
    if state_for_label:
        st.session_state[state_key] = state_for_label


def render_ui():
    st.title("Paycom - Prior Payroll Setup Helper")

    prior_files = st.file_uploader(
        "Paycom Prior Payroll file(s)",
        type=["xlsx", "xls", "csv"],
        key="ppsh_prior",
        accept_multiple_files=True,
    )
    client_name = st.text_input(
        "Client Name",
        value="",
        placeholder="e.g. Chief Delivery",
        key="ppsh_client",
        help="Used in the downloaded file name: <Client Name>_Payroll_Setup_Helper.xlsx",
    )
    # NOTE: the Paycom Scheduled Deductions Report uploader has been removed
    # for now. It will return in a later enhancement, along with the existing
    # "What to set up in Uzio", "Pre-tax vs post-tax", and "Bonus" answers
    # that depend on it. Those helper functions
    # (split_contribs_deductions, classify_pre_post_tax, classify_bonus,
    # build_simplified_xlsx_bytes, etc.) are preserved above so re-enabling
    # is just a UI change.

    if not prior_files:
        st.info("Upload at least one Prior Payroll file to begin.")
        return

    if len(prior_files) > 1:
        st.caption(f"Combining {len(prior_files)} Prior Payroll file(s) for analysis.")

    # Running stores results in session_state so the Auto-Sync toggle bar (which
    # triggers Streamlit reruns) doesn't lose the analysis. Without this, every
    # toggle click would re-collapse the page because the Run button is no longer
    # "pressed" on the rerun.
    if st.button("Run", type="primary"):
        with st.spinner("Analyzing..."):
            try:
                # Concatenate all uploaded Prior Payroll files into a single
                # DataFrame. pd.concat tolerates differing column sets (missing
                # columns become NaN); downstream functions already guard against
                # missing required columns.
                prior_frames = [_read_either(f) for f in prior_files]
                prior_df = pd.concat(prior_frames, ignore_index=True) if len(prior_frames) > 1 else prior_frames[0]
                earnings_from_prior = extract_unique_earnings_from_prior(prior_df)
                deductions_from_prior = extract_unique_deductions_from_prior(prior_df)
                # Filter out UZIO defaults (e.g. Earned Wage Access) BEFORE
                # bifurcation so they never reach the display, the xlsx, or
                # the downstream Tampermonkey automation.
                deductions_from_prior, skipped_default_deds = filter_default_uzio_deductions(deductions_from_prior)
                prior_contribs, prior_deds = bifurcate_match_memo(deductions_from_prior)
                # Attach Pre-tax / Post Tax verdict to the Deductions side only.
                prior_deds = classify_pre_post_from_calc_description(prior_df, prior_deds)
                st.session_state["ppsh_results"] = {
                    "earnings": earnings_from_prior,
                    "contribs": prior_contribs,
                    "deds": prior_deds,
                    "skipped": skipped_default_deds,
                    # RAW (unstripped) source names for the API mapping files, so
                    # "Source ... Code Name" matches the source file verbatim.
                    "earn_raw": _raw_name_map(prior_df, "Earnings"),
                    "ded_raw": _raw_name_map(prior_df, "Deductions"),
                    # Taxes: employee (W/H Taxes) + employer (Client Side Liabilities).
                    "taxes": extract_unique_taxes_from_prior(prior_df),
                }
            except Exception as e:
                st.session_state.pop("ppsh_results", None)
                st.error(f"Failed to analyze the files: {e}")
                raise

    results = st.session_state.get("ppsh_results")
    if not results:
        return

    earnings_from_prior = results["earnings"]
    prior_contribs = results["contribs"]
    prior_deds = results["deds"]
    skipped_default_deds = results["skipped"]
    earn_raw = results.get("earn_raw", {})
    ded_raw = results.get("ded_raw", {})
    taxes_from_prior = results.get("taxes", [])

    # Earnings: first drop UZIO's auto-created defaults (Regular Wage, Overtime,
    # Holiday, etc. — recognized from the Paycom names), then enrich the rest
    # with the UZIO Add-Earning form fields. Earning Type drives Hourly +
    # Taxability; "Other" stays editable. All values editable in the Excel.
    kept_earnings, skipped_earnings = filter_default_uzio_earnings(earnings_from_prior)
    st.markdown("## Earnings (from Prior Payroll file(s))")
    st.caption(
        f"Filtered to `Code Description == \"Earnings\"`, deduped on "
        f"(Type Code, Type Description).  \n"
        f"**{len(kept_earnings)} earning(s) to create**, "
        f"**{len(skipped_earnings)} skipped** as UZIO defaults (auto-created). "
        f"Earning Type drives Hourly/Taxability; `{EARNING_TYPE_OTHER}` rows are "
        f"editable in the Excel."
    )
    if skipped_earnings:
        skip_lines = "; ".join(
            f"{r['Type Code']} {r['Type Description']} → {r['UZIO Default Earning']}"
            for r in skipped_earnings
        )
        st.info(
            f"**Skipped {len(skipped_earnings)} UZIO default earning(s)** "
            f"(already auto-created by UZIO; will not be re-created): `{skip_lines}`"
        )

    # ── "Is the earning Non Discretionary?" toggle section ─────────────────
    # Show a toggle ONLY for earnings whose EFFECTIVE Earning Type is "Bonus"
    # (the only UZIO type with the Include-in-Overtime field). We read the type
    # the user picked in the "Map earnings to an earning type" accordion from
    # session_state (falling back to the tool's inferred type), so changing a
    # bonus to Other removes its toggle and changing something to Bonus adds one.
    # Yes => non-discretionary (included in OT rate); No => discretionary (default).
    def _effective_earn_type(code, td):
        return st.session_state.get(
            f"ppsh_earntype2::{autosync_row_key(code, td)}"
        ) or map_paycom_to_earning_type(code, td)

    bonus_rows = [
        (r["Type Code"], r["Type Description"])
        for r in kept_earnings
        if _effective_earn_type(r["Type Code"], r["Type Description"]) == "Bonus"
    ]
    btd_counts = {}
    for _, td in bonus_rows:
        btd_counts[td] = btd_counts.get(td, 0) + 1

    def _bonus_key(code, td):
        return f"ppsh_nondisc::{autosync_row_key(code, td)}"

    include_in_ot_map = {}
    if bonus_rows:
        st.markdown("### Is the earning Non Discretionary?")
        st.caption(
            f"**{len(bonus_rows)} bonus earning(s)** (excluding Lookback / Realtime). "
            "Toggle **ON = Non-Discretionary** (included in the overtime rate "
            "calculation → *Include Bonus in Overtime = Yes*); "
            "**OFF = Discretionary** (*Include Bonus in Overtime = No*, default)."
        )
        cols = st.columns(2)
        for i, (code, td) in enumerate(bonus_rows):
            label = f"{td} ({code})" if btd_counts[td] > 1 else td
            with cols[i % 2]:
                on = st.toggle(label, value=False, key=_bonus_key(code, td))
            include_in_ot_map[autosync_row_key(code, td)] = "Yes" if on else "No"

    # ── Map earnings to an Earning Type (collapsed accordion; ALL earnings) ──
    # Lists every earning (UZIO system-default earnings are already excluded from
    # kept_earnings). Each dropdown defaults to the type the tool inferred; change
    # it to override. Collapsed by default.
    earning_type_override_map = {}
    if kept_earnings:
        et_opts = list(UZIO_EARNING_TYPES)
        with st.expander("Map earnings to an earning type", expanded=False):
            st.caption(
                f"**{len(kept_earnings)} earning(s)** (UZIO system-default earnings excluded). "
                "Each dropdown defaults to the Earning Type the tool inferred — change it to "
                f"override. `{EARNING_TYPE_OTHER}` exposes the editable fields in UZIO."
            )
            for code, td in [(r["Type Code"], r["Type Description"]) for r in kept_earnings]:
                computed = map_paycom_to_earning_type(code, td)
                idx = et_opts.index(computed) if computed in et_opts else et_opts.index(EARNING_TYPE_OTHER)
                nm_col, pk_col = st.columns([1, 1])
                with nm_col:
                    st.markdown(f"**{code}** — {td}")
                with pk_col:
                    sel = st.selectbox(
                        f"EarnType {code}", options=et_opts, index=idx,
                        key=f"ppsh_earntype2::{autosync_row_key(code, td)}",
                        label_visibility="collapsed",
                    )
                if sel != computed:
                    earning_type_override_map[autosync_row_key(code, td)] = sel

    enriched_earnings = enrich_earnings_for_uzio(
        kept_earnings, include_in_ot_map=include_in_ot_map,
        earning_type_override_map=earning_type_override_map,
    )

    # ── Edit UZIO Earning names ────────────────────────────────────────────
    # Rename any earning as it should appear in UZIO. The edit flows into the
    # setup table below, the Excel, and the earnings mapping file in one shot.
    _render_name_editor(
        "✏️ Edit UZIO Earning names (optional)",
        enriched_earnings, "Earning Name", "ppsh_earnname",
        caption=(
            "Set the exact name each earning should have in UZIO. Flows into the "
            "setup Excel **and** the earnings mapping file — no need to edit two "
            "files. (Skipped UZIO-default earnings aren't listed here; their names "
            "must match UZIO's seeded spellings.)"
        ),
    )

    st.markdown("### UZIO Earning Setup (all form fields)")
    if not enriched_earnings:
        st.caption("(no new earnings to create — all were UZIO defaults, or none found)")
    else:
        st.dataframe(
            pd.DataFrame(enriched_earnings, columns=EARNING_OUTPUT_COLUMNS),
            hide_index=True, use_container_width=True,
        )

    st.markdown("---")

    # Surface what got filtered out before bifurcation so it's transparent.
    if skipped_default_deds:
        skip_lines = "; ".join(
            f"{r['Type Code']} - {r['Type Description']}" for r in skipped_default_deds
        )
        st.info(
            f"**Skipped {len(skipped_default_deds)} default UZIO deduction(s)** "
            f"that UZIO creates automatically (the automation must not re-create them): "
            f"`{skip_lines}`"
        )

    # Deductions extracted from Prior Payroll file(s) via the Match/Memo split:
    # Match/Memo rows are employer Contributions (their own section further below);
    # everything else is an employee-paid Deduction, shown here.
    st.markdown("## Deductions")
    st.caption(
        f"From the Prior Payroll file(s): `Code Description == \"Deductions\"`, deduped on "
        f"(Type Code, Type Description), split by the **Match/Memo** rule. "
        f"**{len(prior_deds)} deduction(s)** here; the **{len(prior_contribs)} Match/Memo "
        f"contribution(s)** are in the Contributions section below.  \n"
        f"**Pre/Post Tax** comes from the Paycom `Calc Description` column: `S125 Pre-Tax` → "
        f"Pre-tax, `After Tax Deduction` → Post Tax, "
        f"`FICA/FUTA/SUTA Taxable Only (401k|403b|457b)` → Pre-tax, anything else → Unknown."
    )
    if not prior_deds:
        st.caption("(none)")
    else:
        st.dataframe(pd.DataFrame(prior_deds), hide_index=True, use_container_width=True)
        # Surface unknown verdicts loudly so we can extend the mapping.
        unknowns = [r for r in prior_deds if r.get("Pre/Post Tax") == "Unknown"]
        if unknowns:
            unknown_descs = sorted({r.get("Calc Description", "") for r in unknowns if r.get("Calc Description")})
            st.warning(
                f"**{len(unknowns)} deduction(s) have an unrecognized Calc Description** "
                f"and were marked **Unknown**. Values seen: `{', '.join(unknown_descs) or '(blank)'}`. "
                f"Tell me how these should map (Pre-tax / Post Tax / something else) and "
                f"I'll extend the mapping."
            )

    # ── Auto-Sync toggle bar ──────────────────────────────────────────────
    # Benefit-type deductions (dental, medical, vision, voluntary life, etc.)
    # expose UZIO's "Auto-Sync from Uzio Benefits" radio. The implementor sets
    # Yes/No per deduction here; non-benefit deductions get "N/A" automatically.
    st.markdown("### Auto-Sync from Uzio Benefits")

    # First pass enrich (empty map) to discover benefit-type rows. Carry
    # (code, td) so toggles are uniquely keyed even when two rows share a
    # Type Description.
    benefit_rows = [
        (r["Type Code"], r["Type Description"])
        for r in enrich_deductions_for_uzio(prior_deds)
        if r.get("Is Benefit Type")
    ]
    # Count descriptions so we only disambiguate labels when one repeats.
    td_counts = {}
    for _, td in benefit_rows:
        td_counts[td] = td_counts.get(td, 0) + 1

    def _widget_key(code, td):
        return f"ppsh_autosync::{autosync_row_key(code, td)}"

    auto_sync_map = {}
    if not benefit_rows:
        st.caption(
            "No benefit-type deductions detected, so Auto-Sync is `N/A` for "
            "every deduction. (Benefit types: dental, medical, vision, voluntary "
            "life, critical illness, accident/cancer insurance, hospital "
            "indemnity, STD, AD&D.)"
        )
    else:
        def _on_select_all():
            v = st.session_state.get("ppsh_autosync_all", False)
            for code, td in benefit_rows:
                st.session_state[_widget_key(code, td)] = v

        st.caption(
            f"**{len(benefit_rows)} benefit-type deduction(s)** can Auto-Sync. "
            "Toggle ON = Auto-Sync **Yes**, OFF = **No** (default OFF). "
            "**Select All** flips every one at once. The setup table below "
            "updates instantly as you toggle."
        )
        st.toggle(
            "Select All",
            value=False,
            key="ppsh_autosync_all",
            on_change=_on_select_all,
        )
        cols = st.columns(2)
        for i, (code, td) in enumerate(benefit_rows):
            label = f"{td} ({code})" if td_counts[td] > 1 else td
            with cols[i % 2]:
                on = st.toggle(label, value=False, key=_widget_key(code, td))
            auto_sync_map[autosync_row_key(code, td)] = "Yes" if on else "No"

    # ── Map deductions to a Master Deduction (collapsed accordion; ALL) ─────
    # base_deds gives the tool's inferred master per deduction; list them all so
    # the user can set/confirm any. Each dropdown defaults to the inferred master
    # (NEEDS_REVIEW for ones the tool couldn't match). Change it to override.
    base_deds = enrich_deductions_for_uzio(prior_deds, auto_sync_map)
    master_override_map = {}
    if base_deds:
        master_opts = [NEEDS_REVIEW] + UZIO_MASTER_DEDUCTIONS
        with st.expander("Map deductions to a master deduction", expanded=False):
            st.caption(
                f"**{len(base_deds)} deduction(s)** (UZIO auto-created defaults already excluded). "
                "Each dropdown defaults to the master the tool inferred — change it to override. "
                f"`{NEEDS_REVIEW}` means the tool couldn't match one."
            )
            for r in base_deds:
                code, td = r["Type Code"], r["Type Description"]
                computed = r.get("UZIO Master Deductions List", NEEDS_REVIEW)
                idx = master_opts.index(computed) if computed in master_opts else 0
                nm_col, pk_col = st.columns([1, 1])
                with nm_col:
                    st.markdown(f"**{code}** — {td}")
                with pk_col:
                    sel = st.selectbox(
                        f"Master {code}", options=master_opts, index=idx,
                        key=f"ppsh_master2::{autosync_row_key(code, td)}",
                        label_visibility="collapsed",
                    )
                if sel != computed:
                    master_override_map[autosync_row_key(code, td)] = sel

    # Final enrich with the chosen Auto-Sync values + any master overrides.
    enriched_deds = enrich_deductions_for_uzio(
        prior_deds, auto_sync_map, master_override_map=master_override_map
    )

    # ── Edit UZIO Deduction names ──────────────────────────────────────────
    # Rename the deduction's DISPLAY name (the Master above still drives the
    # form's type/method). Only creatable rows (master != NEEDS_REVIEW) are shown,
    # since an unmapped row won't be created. The edit flows into the setup table,
    # the Excel, the deductions mapping, AND the contribution links below.
    _render_name_editor(
        "✏️ Edit UZIO Deduction names (optional)",
        [r for r in enriched_deds
         if r.get("UZIO Master Deductions List") != NEEDS_REVIEW],
        "UZIO Deduction Name", "ppsh_dedname",
        caption=(
            "Set the exact display name each deduction should have in UZIO (the "
            "Master Deduction above keeps driving the form). Flows into the setup "
            "Excel, the deductions mapping, and the contribution-link options below."
        ),
    )

    # Show the full UZIO setup table (all form-field columns) so the implementor
    # can eyeball it before downloading. Drop the helper "Is Benefit Type" flag.
    st.markdown("### UZIO Deduction Setup (all form fields)")
    if enriched_deds:
        df_enriched = pd.DataFrame(
            [{k: v for k, v in r.items() if k != "Is Benefit Type"} for r in enriched_deds],
            columns=DEDUCTION_OUTPUT_COLUMNS,
        )
        st.dataframe(df_enriched, hide_index=True, use_container_width=True)
        still_unmapped = [r for r in enriched_deds if r.get("UZIO Master Deductions List") == NEEDS_REVIEW]
        if still_unmapped:
            nr_lines = ", ".join(f"{r['Type Code']} - {r['Type Description']}" for r in still_unmapped)
            st.warning(
                f"**{len(still_unmapped)} deduction(s) still unmapped** (`{NEEDS_REVIEW}`): "
                f"`{nr_lines}`. Use the dropdowns above to map them, or tell me the "
                f"correct master and I'll bake it into the rules."
            )
    else:
        st.caption("(no deductions to set up)")

    # ── Map contributions to a company deduction (collapsed accordion) ──────
    # Each contribution can LINK to a company deduction (by its UZIO DISPLAY NAME —
    # the same label UZIO shows; the name editor above may have renamed it, so we
    # read current display names). 401k/Roth 401k default-link to 401k / Roth 401k;
    # Medical ER memo → Medical Pre/After-tax. Options are the deductions actually
    # created for this client, so we never link to a missing one.
    st.markdown("---")
    st.markdown("## Contributions")
    st.caption(
        f"`Match` or `Memo` in Type Code/Description → employer-side **Contributions** "
        f"(**{len(prior_contribs)} found**). Everything else is a Deduction (above)."
    )
    if not prior_contribs:
        st.caption("(none)")
    else:
        st.dataframe(pd.DataFrame(prior_contribs), hide_index=True, use_container_width=True)

    creatable_deds = [
        r for r in enriched_deds
        if r.get("UZIO Master Deductions List") and r["UZIO Master Deductions List"] != NEEDS_REVIEW
    ]
    available_masters = sorted({r["UZIO Master Deductions List"] for r in creatable_deds})
    available_display = sorted({
        r.get("UZIO Deduction Name") for r in creatable_deds if r.get("UZIO Deduction Name")
    })
    master_to_display = {}
    for r in creatable_deds:
        master_to_display.setdefault(
            r["UZIO Master Deductions List"], r.get("UZIO Deduction Name", "")
        )
    link_options = [CONTRIB_LINK_NONE] + available_display

    link_map = {}
    if not prior_contribs:
        st.caption("Contributions — none found.")
    else:
        with st.expander("Map contributions to a company deduction", expanded=False):
            st.caption(
                f"**{len(prior_contribs)} contribution(s).** Pick the company deduction "
                f"to link each one to (default-mapped for 401k / Roth 401k / Medical). "
                f"Choose `{CONTRIB_LINK_NONE}` to create the contribution without a link. "
                f"**401k / Roth 401k** use Method **{CONTRIB_METHOD_FORMULA}** "
                f"({_format_formula(CONTRIB_FORMULA_TIERS)}); all others use **{CONTRIB_METHOD_FIXED}**."
            )
            for code, td in [(c["Type Code"], c["Type Description"]) for c in prior_contribs]:
                default_master = map_contribution_to_deduction(code, td, available_masters)
                default_opt = master_to_display.get(default_master, CONTRIB_LINK_NONE)
                if default_opt not in link_options:
                    default_opt = CONTRIB_LINK_NONE
                name_col, pick_col = st.columns([1, 1])
                with name_col:
                    st.markdown(f"**{code}** — {td}")
                with pick_col:
                    sel = st.selectbox(
                        f"Link {code}",
                        options=link_options,
                        index=link_options.index(default_opt),
                        key=f"ppsh_contriblinkD::{autosync_row_key(code, td)}",
                        label_visibility="collapsed",
                    )
                link_map[autosync_row_key(code, td)] = "" if sel == CONTRIB_LINK_NONE else sel

    enriched_contribs = enrich_contributions_for_uzio(prior_contribs, link_map)

    # ── Edit UZIO Contribution names ───────────────────────────────────────
    _render_name_editor(
        "✏️ Edit UZIO Contribution names (optional)",
        enriched_contribs, "Contribution Name", "ppsh_conname",
        caption=(
            "Set the exact name each contribution should have in UZIO. Flows into "
            "the setup Excel and the contributions mapping file."
        ),
    )

    if enriched_contribs:
        st.markdown("### UZIO Contribution Setup (all form fields)")
        st.dataframe(
            pd.DataFrame(enriched_contribs, columns=CONTRIBUTION_OUTPUT_COLUMNS),
            hide_index=True, use_container_width=True,
        )

    # ── Taxes (Paycom → UZIO tax-catalog mapping) ──────────────────────────
    # Employee (W/H Taxes) + employer (Client Side Liabilities) taxes. Federal &
    # state auto-map deterministically; local city/school taxes are a best guess
    # the user confirms via a searchable dropdown. Taxes are NOT created in UZIO —
    # we only produce the mapping file.
    st.markdown("---")
    st.markdown("## Taxes")

    catalog = _cached_tax_catalog()
    with st.expander("Tax catalog source (advanced)", expanded=False):
        up = st.file_uploader(
            "Override the UZIO tax catalog with a newer CSV (optional)",
            type=["csv"], key="ppsh_taxcat",
        )
        if up is not None:
            try:
                up.seek(0)
                catalog = load_uzio_tax_catalog(up)
            except Exception as e:
                st.warning(f"Couldn't read the uploaded catalog ({e}); using the bundled one.")
        st.caption(f"Using **{len(catalog)}** UZIO tax entries (bundled catalog unless overridden).")

    tax_map_rows = []
    if not taxes_from_prior:
        st.caption("No taxes found in the prior payroll (W/H Taxes / Client Side Liabilities).")
    elif not catalog:
        st.warning("UZIO tax catalog not found (`uzio_tax_catalog.csv`) — can't map taxes.")
    else:
        dd_state = dominant_tax_state(taxes_from_prior)
        wh = [t for t in taxes_from_prior if t["Section"] == TAX_SECTION_WH]
        emp = [t for t in taxes_from_prior if t["Section"] == TAX_SECTION_EMPLOYER]
        st.caption(
            f"**{len(taxes_from_prior)} tax(es)** — {len(wh)} employee (W/H Taxes) + "
            f"{len(emp)} employer (Client Side Liabilities). Federal & state auto-map; "
            f"**local city/school are a best guess — confirm them**. Dominant state: "
            f"`{dd_state or 'n/a'}`."
        )

        def _tax_opt_label(r):
            base = f"{r['tax_name']}  [{r['tax_code']}]"
            return base + (f" — {r['sub_tax_desc']}" if r['sub_tax_desc'] else "")

        TAX_LEAVE = "— leave unmapped —"
        TAX_ANY_STATE = "🔍 Search all states"
        tax_resolved = {}
        state_options = sorted({r["state_abbreviation"].upper() for r in catalog})
        with st.expander("Map taxes to a UZIO tax", expanded=False):
            st.caption(
                "Federal & State auto-map (reliable). **Local** taxes — and anything the "
                "tool couldn't map — get a **finder**: pick the **State**, **type** part of "
                "the city / school name, then choose the exact tax. The dropdown shows the "
                "UZIO code and Municipal/School so you can tell duplicates apart."
            )
            for t in taxes_from_prior:
                code, td, section = t["Type Code"], t["Type Description"], t["Section"]
                key = tax_row_key(section, code, td)
                res = map_tax_to_uzio(code, td, section, catalog, file_default_state=dd_state)
                tier = res["tier"]
                badge = {"federal": "Federal", "state": "State",
                         "local": "Local — find & confirm", "unknown": "Unmapped — find"}.get(tier, tier)
                # Every tax row gets the same finder (State picker + Search +
                # filtered dropdown) for consistency — pre-selected to the tool's
                # best guess, so confident Federal/State picks still default right.
                use_finder = True

                nm_col, pk_col = st.columns([1, 1])
                with nm_col:
                    st.markdown(f"**{code}** — {td}")
                    st.caption(f"{section} · {badge}")

                if not use_finder:
                    cand = res["candidates"][:60]
                    if res["match"] and res["match"] not in cand:
                        cand = [res["match"]] + cand
                    labels = [TAX_LEAVE] + [_tax_opt_label(r) for r in cand]
                    rowmap = {_tax_opt_label(r): r for r in cand}
                    bl = _tax_opt_label(res["match"]) if res["match"] is not None else None
                    default_index = labels.index(bl) if (bl and bl in labels) else 0
                    with pk_col:
                        sel = st.selectbox(
                            f"Tax {section} {code}", options=labels, index=default_index,
                            key=f"ppsh_tax::{key}", label_visibility="collapsed",
                        )
                    tax_resolved[key] = rowmap.get(sel)
                else:
                    state_key = f"ppsh_taxstate::{key}"
                    search_key = f"ppsh_taxsearch::{key}"
                    tax_key = f"ppsh_tax::{key}"
                    lblmap_key = f"ppsh_taxlblstate::{key}"
                    # Smart default: confident match's state → detected state →
                    # "search all states" (so we never pre-select a wrong state).
                    if res["match"]:
                        guess_state = res["match"].get("state_abbreviation", "").upper()
                    elif res.get("detected"):
                        # state came from the NAME → scope to it
                        guess_state = res["detected"]
                    else:
                        # only a dominant-state fallback (or none) → start in
                        # "Search all states" so we never pre-select a wrong state
                        guess_state = TAX_ANY_STATE
                    if guess_state != TAX_ANY_STATE and guess_state not in state_options:
                        guess_state = TAX_ANY_STATE
                    if state_key not in st.session_state:
                        st.session_state[state_key] = guess_state

                    with pk_col:
                        state_sel = st.selectbox(
                            f"State {section} {code}", options=[TAX_ANY_STATE] + state_options,
                            key=state_key, label_visibility="collapsed",
                        )
                        all_states = (state_sel == TAX_ANY_STATE)
                        search = st.text_input(
                            f"Search {section} {code}",
                            value=_clean_place(td, "" if all_states else state_sel),
                            key=search_key, placeholder="type a city / school name…",
                            label_visibility="collapsed",
                        )
                        # Token-based filter over the chosen scope (one state, or
                        # the whole catalog when "Search all states"). Numeric codes
                        # / 1-2 char noise ignored, so "510101 philadelphia" works.
                        q = (search or "").strip().lower()
                        toks = [w for w in re.split(r"[^a-z0-9]+", q) if len(w) >= 3 and not w.isdigit()]
                        pool = (list(catalog) if all_states
                                else [r for r in catalog if r["state_abbreviation"].upper() == state_sel])
                        if toks:
                            strict = [r for r in pool if all(w in r["tax_name"].lower() for w in toks)]
                            pool = strict if strict else [r for r in pool
                                                          if any(w in r["tax_name"].lower() for w in toks)]
                        ref = " ".join(toks) or _clean_place(td, "" if all_states else state_sel)
                        pool.sort(key=lambda r: difflib.SequenceMatcher(
                            None, ref, _catalog_place(r["tax_name"], r["state_abbreviation"].upper())).ratio(),
                            reverse=True)
                        pool = pool[:50]
                        # Keep the tool's best guess visible (if it fits the scope).
                        if res["match"] and res["match"] not in pool and (
                                all_states or res["match"].get("state_abbreviation", "").upper() == state_sel):
                            pool = [res["match"]] + pool

                        labels = [TAX_LEAVE] + [_tax_opt_label(r) for r in pool]
                        rowmap = {_tax_opt_label(r): r for r in pool}
                        # label → state, so picking a result can auto-fill the State.
                        st.session_state[lblmap_key] = {
                            _tax_opt_label(r): r["state_abbreviation"].upper() for r in pool
                        }
                        # Stable key + validate: if the persisted pick is no longer
                        # an option (scope/search changed), reset to the default.
                        default_label = (_tax_opt_label(res["match"])
                                         if (res["match"] and _tax_opt_label(res["match"]) in labels)
                                         else TAX_LEAVE)
                        if st.session_state.get(tax_key) not in labels:
                            st.session_state[tax_key] = default_label
                        sel = st.selectbox(
                            f"Tax {section} {code}", options=labels, key=tax_key,
                            on_change=_tax_autoselect_state,
                            args=(state_key, tax_key, lblmap_key),
                            label_visibility="collapsed",
                        )
                    tax_resolved[key] = rowmap.get(sel)

        tax_map_rows = build_taxes_mapping_rows(taxes_from_prior, tax_resolved)
        st.markdown("### UZIO Tax Mapping")
        st.dataframe(
            pd.DataFrame(tax_map_rows, columns=MAPPING_TAX_COLUMNS),
            hide_index=True, use_container_width=True,
        )
        unmapped = [r for r in tax_map_rows if not r["Uzio Tax Code"]]
        if unmapped:
            nm = ", ".join(f"{r['Source Tax Code']} {r['Source Tax Code Name']}" for r in unmapped)
            st.warning(
                f"**{len(unmapped)} tax(es) not yet mapped** (`{nm}`) — pick a UZIO tax in the "
                f"accordion above, or they'll be blank in the mapping file."
            )

    # ── Download — single xlsx, three tabs: Earnings | Deductions | Contributions
    st.markdown("---")
    # Filename = "<Client Name>_Payroll_Setup_Helper.xlsx". Falls back to
    # "Client" if the user left the field blank. We strip characters that are
    # invalid in Windows filenames so download doesn't fail on names like
    # "Acme: West Coast" or "Bob's / Tom's".
    raw_name = (client_name or "").strip() or "Client"
    safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1F]+', "", raw_name).strip() or "Client"
    xlsx_bytes = build_3tab_setup_xlsx(
        earnings=enriched_earnings,
        deductions=enriched_deds,
        contributions=enriched_contribs,
    )
    st.download_button(
        "📥 Download Setup Helper (Earnings | Deductions | Contributions)",
        data=xlsx_bytes,
        file_name=f"{safe_name}_Payroll_Setup_Helper.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    # ── Download — API Mapping files (Source name → UZIO name) ─────────────
    # The API run needs the original Source file PLUS three mapping CSVs. Each
    # maps the verbatim source name → the exact UZIO name. The earnings mapping
    # ALSO includes the skipped UZIO-default earnings (Regular Wage, Overtime,
    # Holiday, ...) because the API still uploads data against those.
    st.markdown("---")
    st.markdown("## API Mapping files (Source → UZIO names)")

    earn_map_rows = build_earnings_mapping_rows(enriched_earnings, skipped_earnings, earn_raw)
    ded_map_rows = build_deductions_mapping_rows(enriched_deds, ded_raw)
    contrib_map_rows = build_contributions_mapping_rows(enriched_contribs, ded_raw)

    st.caption(
        f"For the API run, alongside the Source file. **{len(earn_map_rows)} earning(s)** "
        f"(incl. {len(skipped_earnings)} skipped UZIO default(s)), "
        f"**{len(ded_map_rows)} deduction(s)**, **{len(contrib_map_rows)} contribution(s)**, "
        f"**{len(tax_map_rows)} tax(es)**. "
        "⚠️ This is a **critical** file — the *Source* name must match the source file "
        "exactly (spaces included), so it's taken verbatim from the upload; the *UZIO* "
        "name is taken from the exact field the script types into UZIO."
    )

    # Single button → separate CSV downloads with the exact filenames:
    #   <Client>_earnings_mapping.csv / _deductions_mapping.csv /
    #   _contributions_mapping.csv / _taxes_mapping.csv
    # (st.download_button can only send one file per click, so we embed each CSV
    # as base64 and trigger the downloads from a small HTML/JS component.)
    _map_files = [
        [f"{safe_name}_earnings_mapping.csv",
         base64.b64encode(_mapping_csv_bytes(earn_map_rows, MAPPING_EARNING_COLUMNS)).decode("ascii")],
        [f"{safe_name}_deductions_mapping.csv",
         base64.b64encode(_mapping_csv_bytes(ded_map_rows, MAPPING_DEDUCTION_COLUMNS)).decode("ascii")],
        [f"{safe_name}_contributions_mapping.csv",
         base64.b64encode(_mapping_csv_bytes(contrib_map_rows, MAPPING_CONTRIBUTION_COLUMNS)).decode("ascii")],
        [f"{safe_name}_taxes_mapping.csv",
         base64.b64encode(_mapping_csv_bytes(tax_map_rows, MAPPING_TAX_COLUMNS)).decode("ascii")],
    ]
    components.html(
        MAPPING_DOWNLOAD_HTML.replace("__FILES__", json.dumps(_map_files)),
        height=110,
    )

    with st.expander("Preview mapping files (verify exact names before the API run)"):
        st.markdown("**Earnings mapping** (created + skipped UZIO defaults)")
        st.dataframe(
            pd.DataFrame(earn_map_rows, columns=MAPPING_EARNING_COLUMNS),
            hide_index=True, use_container_width=True,
        )
        st.markdown("**Deductions mapping**")
        st.dataframe(
            pd.DataFrame(ded_map_rows, columns=MAPPING_DEDUCTION_COLUMNS),
            hide_index=True, use_container_width=True,
        )
        st.markdown("**Contributions mapping**")
        st.dataframe(
            pd.DataFrame(contrib_map_rows, columns=MAPPING_CONTRIBUTION_COLUMNS),
            hide_index=True, use_container_width=True,
        )
        st.markdown("**Taxes mapping**")
        st.dataframe(
            pd.DataFrame(tax_map_rows, columns=MAPPING_TAX_COLUMNS),
            hide_index=True, use_container_width=True,
        )
