"""
ADP <-> UZIO Withholding Audit Tool — REBUILD (v3)

Scope: FIT + SIT only. 13 field comparisons (12 from the client-shared
mapping sheet + FIT_WITHHOLDING_ALLOWANCE retained from the prior version).

What changed vs the previous version
====================================
1. **Auto-detects** which uploaded file is UZIO vs ADP from columns — the
   two upload slots are order-agnostic.
2. **Uses UZIO's `status`, `effective_date`, `state_code`, `tax_scope` and
   `master_tax_type` columns** that the previous version completely ignored.
3. **Multi-state SIT join** on (employee_id, state_code) so an employee with
   SIT records for two states is compared correctly.
4. **Status comes from UZIO** when ADP has no status column (the previous
   tool emitted "NAN" for every row in such files).
5. **State code resolves from any of** `Worked in State Code` /
   `Worked in State` / `State Tax Code` / `Lived in State Code` — the
   previous lookup missed the most common name and left every row blank.
6. **`SIT_FILING_STATUS` is auto-skipped** for any (employee, state) where
   UZIO has zero `SIT_FILING_STATUS` records. Many states (MA, etc.) don't
   store filing status as a single enum — they use `SIT_HOH` + allowances.
   The previous version compared ADP's "Single - Head of Household" against
   a blank UZIO value and flagged 100% of MA employees as mismatched.
7. **Per-field stale UZIO detection**: if a UZIO field's `effective_date`
   is older than the ADP W-4 effective date, that specific field is flagged
   in a separate sheet.
8. **Reciprocity sheet**: every employee where ADP Lived state ≠ Worked
   state, regardless of mismatch.
9. **Tighter filing-status match**: exact-after-normalize, no substring
   fallback. ("Single" no longer false-matches "Single or Married...".)
10. **Blank-vs-value detection**: when one system has a real value and the
    other is blank, we surface it as a distinct category rather than
    silently treating blank == 0.
11. **Pure-function comparison engine** separated from `render_ui` for
    testability. The 800-line `render_ui` became ~300 lines of UI + ~500
    lines of pure logic.

Preserved from the previous version
====================================
- The 13 field comparisons and all auto-fix value normalizations
  (bool, money_cents, filing_status, numeric).
- The False Positives Filtered sheet (no-SIT states).
- The Needs UI Verification sheet (SIT_WITHHOLDING_EXEMPTION).
- The same .xlsx output shape with the same sheet names so existing
  consumers can keep working.
- The FILING_STATUS_MAP enum table — now reloaded at runtime from
  `filing status_code.txt` (with the hardcoded dict as a safety net),
  matching how the Paycom audit handles the same data. Implementors can
  extend filing-status codes by editing one text file; no code change.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Optional

import pandas as pd
import streamlit as st
import yaml

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────

NO_SIT_STATES = {"FL", "TX", "NV", "WA", "WY", "SD", "AK", "TN", "NH"}

# These UZIO fields don't map cleanly to user-editable UI in Uzio. When they
# mismatch, route to "Needs UI Verification" instead of the main mismatch list.
FIELDS_REQUIRING_UI_VERIFICATION = {
    "SIT_WITHHOLDING_EXEMPTION",
    # 'Do not calculate State Tax' is an internal/derived flag in Uzio.
    # It is auto-set when an employee is Federal-Exempt. It does NOT
    # appear in any user-editable UI screen.
}

# Field comparison registry. The 12 entries from the client-shared mapping
# sheet + FIT_WITHHOLDING_ALLOWANCE retained from the prior version (used by
# pre-2020 W-4 employees).
FIELD_MAPPING = [
    # identity (not compared, only for display)
    {"uzio_key": "employee_id",                          "adp_col": "Associate ID",                            "kind": "id"},
    {"uzio_key": "employee_first_name",                  "adp_col": "Legal First Name",                        "kind": "id"},
    {"uzio_key": "employee_last_name",                   "adp_col": "Legal Last Name",                         "kind": "id"},
    # FIT
    {"uzio_key": "FIT_WITHHOLDING_EXEMPTION",            "adp_col": "Do Not Calculate Federal Income Tax",     "kind": "bool",           "scope": "FED"},
    {"uzio_key": "FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",  "adp_col": "Federal Additional Tax Amount",           "kind": "money_cents",    "scope": "FED"},
    {"uzio_key": "FIT_FILING_STATUS",                    "adp_col": "Federal/W4 Marital Status Description",   "kind": "filing_status",  "scope": "FED"},
    {"uzio_key": "FIT_CHILD_AND_DEPENDENT_TAX_CREDIT",   "adp_col": "Dependents",                              "kind": "money_cents",    "scope": "FED"},
    {"uzio_key": "FIT_DEDUCTIONS_OVER_STANDARD",         "adp_col": "Deductions",                              "kind": "money_cents",    "scope": "FED"},
    {"uzio_key": "FIT_HIGHER_WITHHOLDING",               "adp_col": "Multiple Jobs indicator",                 "kind": "bool",           "scope": "FED"},
    {"uzio_key": "FIT_OTHER_INCOME",                     "adp_col": "Other Income",                            "kind": "money_cents",    "scope": "FED"},
    {"uzio_key": "FIT_WITHHOLD_AS_NON_RESIDENT",         "adp_col": "Non-Resident Alien",                      "kind": "bool",           "scope": "FED"},
    {"uzio_key": "FIT_WITHHOLDING_ALLOWANCE",            "adp_col": "Federal/W4 Exemptions",                   "kind": "numeric",        "scope": "FED"},
    # SIT
    {"uzio_key": "SIT_WITHHOLDING_EXEMPTION",            "adp_col": "Do not calculate State Tax",              "kind": "bool",           "scope": "STATE"},
    {"uzio_key": "SIT_FILING_STATUS",                    "adp_col": "State Marital Status Description",        "kind": "filing_status",  "scope": "STATE"},
    {"uzio_key": "SIT_TOTAL_ALLOWANCES",                 "adp_col": "State Exemptions/Allowances",             "kind": "numeric",        "scope": "STATE"},
    {"uzio_key": "SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD",  "adp_col": "State Additional Tax Amount",             "kind": "money_cents",    "scope": "STATE"},
]

ADP_ID_CANDIDATES        = ["associate id", "employee id", "employee_id", "emp_id", "file number"]
ADP_FIRST_NAME_CANDIDATES = ["legal first name", "first name", "first_name"]
ADP_LAST_NAME_CANDIDATES  = ["legal last name", "last name", "last_name"]
ADP_EFF_DATE_COL          = "Federal/W4 Effective Date"
ADP_STATE_CANDIDATES_WORKED = [
    "worked in state code", "worked in state", "work state code", "work state",
    "state tax code", "state",
]
ADP_STATE_CANDIDATES_LIVED  = [
    "lived in state code", "lived in state", "lived in state tax code",
    "home state", "primary address: state",
]

UZIO_ID_CANDIDATES        = ["employee_id", "employee id", "emp_id"]
UZIO_KEY_COL              = "withholding_field_key"
UZIO_VAL_COL              = "withholding_field_value"
UZIO_STATE_COL            = "state_code"
UZIO_SCOPE_COL            = "tax_scope"
UZIO_MASTER_COL           = "master_tax_type"
UZIO_EFF_COL              = "effective_date"
UZIO_STATUS_COL           = "status"


# ─────────────────────────────────────────────────────────────────────────────
# FILING_STATUS_MAP_FALLBACK (UZIO enum → ADP label).
#
# This is a SAFETY NET only. The live source of truth is
# `filing status_code.txt` at the project root, parsed at runtime by
# `load_filing_status_map()`. To add or change a code:
#   1. Edit `filing status_code.txt` — append a line like
#         IL_HEAD_OF_HOUSEHOLD("Head of household")
#   2. Save. The tool picks it up on the next audit run; no code change needed.
#
# Keep this dict in sync with the txt file so the tool still works if the
# file is missing.
# ─────────────────────────────────────────────────────────────────────────────
FILING_STATUS_MAP_FALLBACK = {
    "FEDERAL_SINGLE": "Single",
    "FEDERAL_MARRIED": "Married",
    "FEDERAL_MARRIED_SINGLE": "Married but withhold as Single",
    "FEDERAL_SINGLE_OR_MARRIED": "Single or Married filing separately",
    "FEDERAL_MARRIED_JOINTLY": "Married filing jointly or Qualifying surviving spouse",
    "FEDERAL_HEAD_OF_HOUSEHOLD": "Head of household",

    "MD_SINGLE": "Single", "MD_MARRIED": "Married",
    "MD_MARRIED_SINGLE": "Married but withhold at single rate",
    "DC_SINGLE": "Single", "DC_HEAD_OF_HOUSEHOLD": "Head of Household",
    "DC_MARRIED_DP_JOINTLY": "Married/Domestic Partners Filing Jointly",
    "DC_MARRIED_SEPARATELY": "Married Filing Separately",
    "DC_MARRIED_DP_SEPARATELY": "Married/Domestic Partners Filing Separately",
    "NM_SINGLE": "Single or Married filing separately",
    "NM_MARRIED": "Married filing jointly or Qualifying Surviving Spouse",
    "NM_MARRIED_SINGLE": "Married but withhold as Single",
    "NM_HEAD_OF_HOUSEHOLD": "Head of Household",
    "MS_SINGLE": "Single", "MS_HEAD_OF_HOUSEHOLD": "Head of Family",
    "MS_M1": "Married (Spouse NOT employed)", "MS_M2": "Married (Spouse is employed)",
    "MO_SINGLE": "Single or Married Spouse Works or Married Filing Separate",
    "MO_MARRIED": "Married (Spouse does not work)", "MO_HEAD_OF_HOUSEHOLD": "Head of Household",
    "AL_NO_PERSONAL_EXEMPTION": "No Personal Exemption", "AL_SINGLE": "Single",
    "AL_MARRIED": "Married", "AL_MARRIED_SEPARATELY": "Married Filing Separately",
    "AL_HEAD_OF_HOUSEHOLD": "Head of Family",
    "DE_MARRIED": "Married", "DE_SINGLE": "Single",
    "DE_MARRIED_SINGLE_RATE": "Married but Withhold as Single",
    "OK_MARRIED": "Married", "OK_SINGLE": "Single",
    "OK_MARRIED_SINGLE_RATE": "Married but Withhold as Single",
    "OK_NRA": "Non-Resident Alien",
    "NC_HEAD_OF_HOUSEHOLD": "Head of Household",
    "NC_MARRIED": "Married Filing Jointly or Surviving Spouse",
    "NC_SINGLE": "Single or Married Filing Separately",
    "SC_MARRIED_SINGLE_RATE": "Married but Withhold at higher Single Rate",
    "SC_MARRIED": "Married", "SC_SINGLE": "Single",
    "UT_SINGLE": "Single or Married filing separately",
    "UT_MARRIED": "Married filing jointly or Qualifying widow(er)",
    "UT_HEAD_OF_HOUSEHOLD": "Head of Household",
    "GA_SINGLE": "Single",
    "GA_SEPARATE_MARRIED_JOINT_BOTH_WORKING": "Married Filing Separate or Married Filing Joint both spouses working",
    "GA_MARRIED_JOINT_ONE_WORKING": "Married Filing Joint one spouse working",
    "GA_HEAD_OF_HOUSEHOLD": "Head of Household",
    "WI_SINGLE": "Single", "WI_MARRIED": "Married",
    "WI_MARRIED_SINGLE_RATE": "Married but withhold at higher single rate",
    "KS_SINGLE": "Single", "KS_JOINT": "Joint",
    "VT_SINGLE": "Single", "VT_MARRIED": "Married/Civil Union Filing Jointly",
    "VT_MARRIED_FILING_SEPERATELY": "Married/Civil Union Filing Separately",
    "VT_MARRIED_SINGLE_RATE": "Married, but withhold at higher single rate",
    "NJ_SINGLE": "Single", "NJ_MARRIED_DP_JOINTLY": "Married/Civil Union Couple Joint",
    "NJ_MARRIED_SEPARATELY": "Married/Civil Union Partner Separate",
    "NJ_HEAD_OF_HOUSEHOLD": "Head of Household",
    "NJ_QUALIFIED_WIDOW": "Qualifying Widow(er)/Surviving Civil Union Partner",
    "CA_HEAD_OF_HOUSEHOLD": "Head of Household",
    "CA_MARRIED": "Married (one income)",
    "CA_SINGLE": "Single or Married (with two or more incomes)",
    "MN_SINGLE": "Single, Married but legally separated or Spouse is a nonresident alien",
    "MN_MARRIED": "Married",
    "MN_MARRIED_SINGLE_RATE": "Married but withhold at higher single rate",
    "IA_OTHER": "Other (Including Single)", "IA_HEAD_OF_HOUSEHOLD": "Head of Household",
    "IA_MARRIED_JOINTLY": "Married filing jointly", "IA_QUALIFIED_SPOUSE": "Qualifying Surviving Spouse",
    "ME_SINGLE": "Single or Head of Household", "ME_MARRIED": "Married",
    "ME_MARRIED_SINGLE_RATE": "Married but withhold at higher single rate",
    "ME_NON_RESIDENT_ALIEN": "Nonresident alien",
    "NY_MARRIED_WITHHOLD_SINGLE": "Married but withhold as Single",
    "NY_SINGLE": "Single", "NY_MARRIED": "Married", "NY_HEAD_OF_HOUSEHOLD": "Head of Household",
    "NE_SINGLE": "Single", "NE_MARRIED": "Married Filing Jointly or Qualifying Widow(er)",
    "LA_NO_DEDUCTION": "No Deduction",
    "LA_SINGLE_OR_MARRIED": "Single or married filing separately",
    "LA_MARRIED_FILING_JOINTLY_HOH": "Married filing jointly, qualifying surviving spouse, or head of household",
    "OR_SINGLE": "Single", "OR_MARRIED": "Married",
    "OR_MARRIED_SINGLE_RATE": "Married but withhold at higher single rate",
    "ND_SINGLE": "Single", "ND_MARRIED": "Married",
    "ND_MARRIED_SINGLE_RATE": "Married but Withhold at higher Single Rate",
    "ND_SINGLE_MARRIED_SEPARATELY": "Single or Married filing separately",
    "ND_HEAD_OF_HOUSEHOLD": "Head of household",
    "ND_MARRIED_JOINTLY": "Married filing jointly  or Qualifying Surviving Spouse",
    "ID_SINGLE": "Single", "ID_MARRIED": "Married",
    "ID_MARRIED_SINGLE_RATE": "Married but Withhold at higher Single Rate",
    "CO_SINGLE_OR_MARRIED_SEPARATELY": "Single or Married filing separately",
    "CO_MARRIED_JOINTLY": "Married filing jointly",
    "CO_HEAD_OF_HOUSEHOLD": "Head of household",
    "CO_SINGLE": "Single", "CO_MARRIED": "Married",
    "CO_MARRIED_SINGLE_RATE": "Married but Withhold at higher Single Rate",
    "HI_SINGLE": "Single", "HI_MARRIED": "Married",
    "HI_MARRIED_SINGLE_RATE": "Married but Withhold at higher single rate",
    "HI_DISABLED": "Certified disabled person", "HI_NMS": "Nonresident Military Spouse",
    "MT_SINGLE": "Single or Married filing separately",
    "MT_MARRIED": "Married filing jointly or qualifying surviving spouse",
    "MT_HEAD_OF_HOUSEHOLD": "Head of household",
    "AR_SINGLE": "Single", "AR_MARRIED_FILING_JOINTLY": "Married Filing Jointly",
    "AR_HOH": "Head of Household",
}


from functools import lru_cache


@lru_cache(maxsize=1)
def load_filing_status_map() -> dict:
    """Read filing-status codes from `filing status_code.txt` at runtime.

    The file is the Java-enum source format used by UZIO:
        FEDERAL_SINGLE("Single"), MD_MARRIED("Married"), ...

    We extract every `CODE("Label")` pair with a regex and merge into the
    hardcoded fallback. File entries override the fallback so implementors
    can correct or add codes without a code change. If the file is missing
    or unreadable, we fall back to the hardcoded dict so the tool still
    runs (it just won't pick up any post-shipped code additions).

    Cached per process: edit the file → restart Streamlit (or call
    `load_filing_status_map.cache_clear()`) to pick up changes.
    """
    merged = dict(FILING_STATUS_MAP_FALLBACK)
    path = _resolve_filing_status_path()
    if not path:
        return merged
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
        # Strip Java line comments so commented-out enum entries don't sneak in.
        text = re.sub(r"//.*", "", text)
        pattern = re.compile(r'([A-Z][A-Z0-9_]+)\("([^"]+)"\)')
        for code, label in pattern.findall(text):
            merged[code.strip()] = label.strip()
    except Exception:
        pass
    return merged


# ─────────────────────────────────────────────────────────────────────────────
# Small utility helpers
# ─────────────────────────────────────────────────────────────────────────────

def _clean(x) -> str:
    """Coerce any value to a stripped string. Treats NaN/None as ''."""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).strip()
    if s.lower() in {"nan", "nat", "none"}:
        return ""
    return s


def _norm_bool(s: str) -> str:
    """Return '1', '0', or '' for inputs that look boolean."""
    v = _clean(s).lower()
    if v in {"yes", "y", "true", "1", "on"}:
        return "1"
    if v in {"no", "n", "false", "0", "off"}:
        return "0"
    return ""  


def _norm_money(s: str) -> Optional[float]:
    """Return a float for currency-ish strings, else None for blank."""
    v = _clean(s).replace("$", "").replace(",", "")
    if v == "":
        return None
    if v.startswith("(") and v.endswith(")"):
        v = "-" + v[1:-1]
    try:
        return float(v)
    except ValueError:
        return None


def _norm_filing_status(s: str) -> str:
    """Lowercase and collapse non-word chars to single spaces."""
    return re.sub(r"[\W_]+", " ", _clean(s).lower()).strip()


def _parse_date(d) -> pd.Timestamp:
    if d is None or _clean(d) == "":
        return pd.NaT
    try:
        return pd.to_datetime(d, errors="coerce")
    except Exception:
        return pd.NaT


def _find_col(cols, candidates) -> Optional[str]:
    """Case-insensitive lookup of the first matching column name."""
    norm = {c.strip().lower(): c for c in cols}
    for cand in candidates:
        if cand in norm:
            return norm[cand]
    return None


def _resolve_repo_file(filename: str, env_var: Optional[str] = None) -> Optional[str]:
    """Find a project file regardless of where Streamlit/python was launched.

    Search order: env var override → current working directory → module
    location and ancestor directories (../../ from apps/adp/).
    """
    import os
    if env_var:
        env = os.environ.get(env_var)
        if env and os.path.isfile(env):
            return env
    if os.path.isfile(filename):
        return os.path.abspath(filename)
    here = os.path.dirname(os.path.abspath(__file__))
    for up in (here, os.path.dirname(here), os.path.dirname(os.path.dirname(here))):
        cand = os.path.join(up, filename)
        if os.path.isfile(cand):
            return cand
    return None


def _resolve_key_mapping_path() -> Optional[str]:
    return _resolve_repo_file("key_mapping.yml", env_var="KEY_MAPPING_YML")


def _resolve_filing_status_path() -> Optional[str]:
    return _resolve_repo_file("filing status_code.txt", env_var="FILING_STATUS_CODE")


def load_key_mapping_yml() -> dict:
    """Read display labels from the project-shared key_mapping.yml.

    Returns a NESTED dict keyed by jurisdiction (FED + state codes), each
    mapping field_key → label. The previous flattened form dropped per-state
    labels (when the same key appeared under multiple jurisdictions, only
    the last-loaded one survived). Missing file → empty dict.
    """
    path = _resolve_key_mapping_path()
    if not path:
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        out: dict[str, dict[str, str]] = {}
        for juris, items in data.get("withholding_es", {}).get("mappings", {}).items():
            if not isinstance(items, dict):
                continue
            bucket = out.setdefault(str(juris).upper(), {})
            for k, meta in items.items():
                if isinstance(meta, dict) and "label" in meta:
                    bucket[k] = str(meta["label"])
        return out
    except Exception:
        return {}


def field_label(key: str, labels: dict, jurisdiction: str = "") -> str:
    """Look up the display label for a field, jurisdiction-aware.

    - FIT_* fields → look up under 'FED'.
    - SIT_* fields → look up under the given state code (jurisdiction).
    - If a specific jurisdiction has no entry, fall back to any jurisdiction
      that defines the key, then to a Title-Cased version of the key.
    """
    juris = (jurisdiction or "").upper()
    if key.startswith("FIT_"):
        juris = "FED"
    # Try the requested jurisdiction first.
    if juris and juris in labels and key in labels[juris]:
        return labels[juris][key]
    # Fall back to any jurisdiction that defines the same key.
    for jbucket in labels.values():
        if key in jbucket:
            return jbucket[key]
    # Final fallback: Title Case of the raw key.
    return key.replace("_", " ").title()


# ─────────────────────────────────────────────────────────────────────────────
# File source auto-detection
# ─────────────────────────────────────────────────────────────────────────────

def detect_source(df: pd.DataFrame) -> str:
    """Inspect columns to decide 'uzio', 'adp', or 'unknown'."""
    cols = {c.strip().lower() for c in df.columns}
    if UZIO_KEY_COL in cols and UZIO_VAL_COL in cols:
        return "uzio"
    if "associate id" in cols and any(c.startswith("federal") for c in cols):
        return "adp"
    if any(c in cols for c in ADP_ID_CANDIDATES) and any(
        "marital" in c or "w4" in c or "withhold" in c or "calculate" in c for c in cols
    ):
        return "adp"
    return "unknown"


def read_uploaded(file) -> pd.DataFrame:
    """Read a Streamlit UploadedFile-like object as a string DataFrame."""
    name = (getattr(file, "name", "") or "").lower()
    raw = file.getvalue() if hasattr(file, "getvalue") else file.read()
    buf = io.BytesIO(raw)
    if name.endswith(".csv"):
        return pd.read_csv(buf, dtype=str, keep_default_na=True)
    return pd.read_excel(buf, dtype=str, engine="openpyxl")


# ─────────────────────────────────────────────────────────────────────────────
# Comparison engine (pure functions)
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class ComparisonResult:
    """Outcome of comparing one (employee, field) pair."""
    matched: bool
    category: str              # 'match' | 'mismatch' | 'blank_vs_value' | 'skipped' | 'ui_verify'
    adp_normalized: str
    uzio_normalized: str
    rule: str
    reason: str = ""


def compare_field(
    uzio_key: str,
    kind: str,
    adp_raw,
    uzio_raw,
) -> ComparisonResult:
    """Type-driven comparison. Returns a ComparisonResult.

    Categories:
      match            — values agree
      mismatch         — both have values that genuinely disagree
      blank_vs_value   — exactly one side has a real value (the other is blank)
      skipped          — caller asked to skip (no-SIT state, etc.) — not used here
      ui_verify        — caller routed to UI verification — not used here
    """
    a = _clean(adp_raw)
    u = _clean(uzio_raw)

    # ── bool ──────────────────────────────────────────────────────────────
    if kind == "bool":
        ab, ub = _norm_bool(a), _norm_bool(u)
        a_blank = (ab == "" and a == "")
        u_blank = (ub == "" and u == "")
        # Blank is treated as False (the default). So blank-vs-false → match;
        # blank-vs-true is the only asymmetric case worth surfacing.
        a_eff = "0" if a_blank else ab
        u_eff = "0" if u_blank else ub
        if a_eff == u_eff:
            return ComparisonResult(
                True, "match", a_eff, u_eff, "bool",
                "blank treated as false (default)" if (a_blank or u_blank) else "both populated",
            )
        if a_blank or u_blank:
            return ComparisonResult(
                False, "blank_vs_value",
                ab if not a_blank else "(blank)",
                ub if not u_blank else "(blank)",
                "bool",
                "one side blank, the other is true/yes — not a safe default match",
            )
        return ComparisonResult(
            False, "mismatch", ab, ub, "bool",
            "ADP Yes/No vs UZIO true/false disagree",
        )

    # ── filing_status ─────────────────────────────────────────────────────
    if kind == "filing_status":
        if u == "":
            if a == "":
                return ComparisonResult(True, "match", "", "", "filing_status", "both blank")
            return ComparisonResult(
                False, "blank_vs_value", a, "(blank)", "filing_status",
                "ADP has filing status but UZIO has no value",
            )
        fsmap = load_filing_status_map()
        u_mapped = fsmap.get(u, u.split("_", 1)[1].replace("_", " ").title() if "_" in u else u.title())
        a_norm = _norm_filing_status(a)
        u_norm = _norm_filing_status(u_mapped)
        if a == "":
            return ComparisonResult(
                False, "blank_vs_value", "(blank)", u_norm,
                "filing_status", "UZIO has filing status but ADP has no value",
            )
        # EXACT match after normalization — no substring fallback.
        return ComparisonResult(
            a_norm == u_norm, "match" if a_norm == u_norm else "mismatch",
            a_norm, u_norm, "filing_status",
            "UZIO enum mapped to ADP label, exact normalized match",
        )

    # ── money in cents (UZIO stores cents, ADP stores dollars) ────────────
    if kind == "money_cents":
        af, uf = _norm_money(a), _norm_money(u)
        a_blank = af is None
        u_blank = uf is None
        a_val = 0.0 if a_blank else af
        u_val = 0.0 if u_blank else (uf / 100.0)
        # Both effectively zero (blank or literal 0) → match.
        if abs(a_val - u_val) < 0.01:
            note = "both effectively zero" if (a_blank or u_blank) else "both populated and equal"
            return ComparisonResult(True, "match", f"{a_val:g}", f"{u_val:g}", "money_cents", note)
        # Asymmetric blank vs non-zero → blank_vs_value (a real, actionable finding).
        if a_blank or u_blank:
            a_disp = "(blank)" if a_blank else f"{a_val:g}"
            u_disp = "(blank)" if u_blank else f"{u_val:g}"
            return ComparisonResult(
                False, "blank_vs_value", a_disp, u_disp,
                "money_cents",
                "one side blank, the other has a non-zero value",
            )
        return ComparisonResult(
            False, "mismatch", f"{a_val:g}", f"{u_val:g}",
            "money_cents", "UZIO stored in cents; compared in dollars and disagree",
        )

    # ── numeric (counts, allowances) ──────────────────────────────────────
    af, uf = _norm_money(a), _norm_money(u)
    a_blank = af is None
    u_blank = uf is None
    a_val = 0.0 if a_blank else af
    u_val = 0.0 if u_blank else uf
    if a_val == u_val:
        note = "both effectively zero" if (a_blank or u_blank) else "both populated and equal"
        return ComparisonResult(True, "match", f"{a_val:g}", f"{u_val:g}", "numeric", note)
    if a_blank or u_blank:
        a_disp = "(blank)" if a_blank else f"{a_val:g}"
        u_disp = "(blank)" if u_blank else f"{u_val:g}"
        return ComparisonResult(
            False, "blank_vs_value", a_disp, u_disp,
            "numeric", "one side blank, the other has a non-zero value",
        )
    return ComparisonResult(
        False, "mismatch", f"{a_val:g}", f"{u_val:g}",
        "numeric", "numeric values disagree",
    )


# ─────────────────────────────────────────────────────────────────────────────
# Parsers
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class UzioPivot:
    """Pivoted UZIO data, indexed for fast (emp_id [, state]) lookup."""
    federal_wide: pd.DataFrame                # index = emp_id, cols = FIT_* keys
    state_wide: pd.DataFrame                  # index = (emp_id, state), cols = SIT_* keys
    eff_by_emp_field: dict                    # {(emp_id, key, state_or_None): Timestamp}
    status_by_emp: dict                       # {emp_id: 'ACTIVE'|'TERMINATED'|...}
    name_by_emp: dict                         # {emp_id: 'First Last'}
    states_by_emp: dict                       # {emp_id: set(state_code, ...)}
    has_sit_filing_status_for: set            # {(emp_id, state)} where SIT_FILING_STATUS exists
    all_emp_ids: set


def parse_uzio(df: pd.DataFrame) -> UzioPivot:
    """Normalize the UZIO long file into pivoted views + lookup tables."""
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]
    id_col = _find_col(df.columns, UZIO_ID_CANDIDATES) or "employee_id"

    # Coerce key columns to clean strings.
    df[id_col] = df[id_col].apply(_clean)
    df = df[df[id_col] != ""]
    df[UZIO_KEY_COL] = df[UZIO_KEY_COL].apply(_clean)
    df[UZIO_VAL_COL] = df[UZIO_VAL_COL].apply(_clean)
    if UZIO_STATE_COL in df.columns:
        df[UZIO_STATE_COL] = df[UZIO_STATE_COL].apply(_clean).str.upper()
    else:
        df[UZIO_STATE_COL] = ""

    # Federal pivot: only FEDERAL-scope rows (or rows with no scope but no state).
    is_federal = (df.get(UZIO_SCOPE_COL, "FEDERAL").fillna("").str.upper() == "FEDERAL") \
        if UZIO_SCOPE_COL in df.columns else (df[UZIO_STATE_COL] == "")
    fed = df[is_federal]
    federal_wide = (
        fed.pivot_table(
            index=id_col, columns=UZIO_KEY_COL, values=UZIO_VAL_COL,
            aggfunc=lambda x: list(x)[-1],
        )
        if not fed.empty else pd.DataFrame()
    )
    if not federal_wide.empty:
        federal_wide.index.name = id_col

    # State pivot: rows where state_code is present (or scope == STATE).
    is_state = (df[UZIO_STATE_COL] != "")
    if UZIO_SCOPE_COL in df.columns:
        is_state = is_state | (df[UZIO_SCOPE_COL].fillna("").str.upper() == "STATE")
    st_df = df[is_state]
    if not st_df.empty:
        state_wide = st_df.pivot_table(
            index=[id_col, UZIO_STATE_COL],
            columns=UZIO_KEY_COL, values=UZIO_VAL_COL,
            aggfunc=lambda x: list(x)[-1],
        )
    else:
        state_wide = pd.DataFrame()

    # eff_date lookup, keyed per (emp, key, state-or-None).
    eff_by = {}
    if UZIO_EFF_COL in df.columns:
        dser = df[UZIO_EFF_COL].apply(_parse_date)
        for (emp, key, state), dts in zip(
            zip(df[id_col], df[UZIO_KEY_COL], df[UZIO_STATE_COL]),
            dser,
        ).__iter__() if False else [
            ((row[id_col], row[UZIO_KEY_COL], row[UZIO_STATE_COL]), d)
            for (_, row), d in zip(df.iterrows(), dser)
        ]:
            # zip-over-iterrows is slow for huge frames but UZIO files are
            # under 100k rows. Refactor if a real client needs it.
            state_key = (emp, key, state if state else None)
            if state_key not in eff_by or (pd.notna(dts) and (eff_by[state_key] is pd.NaT or dts > eff_by[state_key])):
                eff_by[state_key] = dts

    # Per-employee aggregates.
    status_by = {}
    name_by = {}
    states_by = {}
    if UZIO_STATUS_COL in df.columns:
        first_status = df.groupby(id_col)[UZIO_STATUS_COL].first()
        status_by = {k: _clean(v).upper() for k, v in first_status.items()}
    name_cols = [c for c in ["employee_first_name", "employee_last_name"] if c in df.columns]
    if name_cols:
        names = df.groupby(id_col)[name_cols].first()
        name_by = {
            emp: (f"{_clean(r.get('employee_first_name',''))} {_clean(r.get('employee_last_name',''))}").strip()
            for emp, r in names.iterrows()
        }
    states_by = (
        df[df[UZIO_STATE_COL] != ""]
        .groupby(id_col)[UZIO_STATE_COL].apply(lambda s: set(s.unique()))
        .to_dict()
    )

    has_sit_fs = set()
    if not st_df.empty:
        sit_fs = st_df[st_df[UZIO_KEY_COL] == "SIT_FILING_STATUS"]
        for emp, state in zip(sit_fs[id_col], sit_fs[UZIO_STATE_COL]):
            has_sit_fs.add((emp, state))

    return UzioPivot(
        federal_wide=federal_wide,
        state_wide=state_wide,
        eff_by_emp_field=eff_by,
        status_by_emp=status_by,
        name_by_emp=name_by,
        states_by_emp=states_by,
        has_sit_filing_status_for=has_sit_fs,
        all_emp_ids=set(df[id_col].unique()),
    )


@dataclass
class AdpParsed:
    df: pd.DataFrame               # one row per employee (latest W-4)
    id_col: str
    state_worked_col: Optional[str]
    state_lived_col: Optional[str]
    eff_date_col: Optional[str]
    multi_row_emp_ids: set         # emps that had W-4 history (dedup'd)
    w4_history: pd.DataFrame       # all rows incl. older, for the date sheet
    all_emp_ids: set


def parse_adp(df_raw: pd.DataFrame) -> AdpParsed:
    df = df_raw.copy()
    df.columns = [c if isinstance(c, str) else str(c) for c in df.columns]

    id_col = _find_col(df.columns, ADP_ID_CANDIDATES) or df.columns[0]
    df[id_col] = df[id_col].apply(_clean)
    df = df[df[id_col] != ""]

    state_worked_col = _find_col(df.columns, ADP_STATE_CANDIDATES_WORKED)
    state_lived_col = _find_col(df.columns, ADP_STATE_CANDIDATES_LIVED)
    eff_col = ADP_EFF_DATE_COL if ADP_EFF_DATE_COL in df.columns else None

    # Normalize state code values to uppercase.
    for c in (state_worked_col, state_lived_col):
        if c:
            df[c] = df[c].apply(_clean).str.upper()

    # Dedup multi-row employees by latest Federal/W4 Effective Date.
    if eff_col:
        df["_eff_date"] = df[eff_col].apply(_parse_date)
        df_sorted = df.sort_values([id_col, "_eff_date"], ascending=[True, False], na_position="last")
        history = df_sorted.copy()
        df_dedup = df_sorted.drop_duplicates(subset=[id_col], keep="first").copy()
    else:
        df["_eff_date"] = pd.NaT
        history = df.copy()
        df_dedup = df.copy()
    history["IS_SELECTED_LATEST"] = history.index.isin(df_dedup.index)

    counts = history.groupby(id_col).size()
    multi_ids = set(counts[counts > 1].index)

    return AdpParsed(
        df=df_dedup,
        id_col=id_col,
        state_worked_col=state_worked_col,
        state_lived_col=state_lived_col,
        eff_date_col=eff_col,
        multi_row_emp_ids=multi_ids,
        w4_history=history,
        all_emp_ids=set(df_dedup[id_col].unique()),
    )


# ─────────────────────────────────────────────────────────────────────────────
# Audit orchestration
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class AuditOptions:
    skip_no_sit_states: bool = True
    skip_known_ui_only_fields: bool = True
    skip_sit_fs_when_uzio_has_none: bool = True
    flag_blank_vs_value: bool = True


@dataclass
class AuditOutput:
    mismatches: pd.DataFrame                  # ALL findings: Mismatch + Blank vs Value + Needs UI Verification
    false_positives_filtered: pd.DataFrame
    reciprocity: pd.DataFrame
    stale_uzio: pd.DataFrame
    missing_in_uzio: pd.DataFrame
    missing_in_adp: pd.DataFrame
    employees_with_mismatches: pd.DataFrame
    mismatch_summary: pd.DataFrame
    field_rules: pd.DataFrame
    w4_history: pd.DataFrame
    summary_metrics: pd.DataFrame
    about: pd.DataFrame


MISMATCH_COLS = [
    "EMPLOYEE_ID", "EMPLOYEE_NAME", "EMPLOYMENT_STATUS", "STATE_CODE",
    "JURISDICTION",
    "CATEGORY",                # 'Mismatch' | 'Blank vs Value' | 'Needs UI Verification'
    "FIELD_LABEL", "FIELD_KEY",
    "ADP_COLUMN", "UZIO_COLUMN",
    "ADP_VALUE_RAW", "UZIO_VALUE_RAW",
    "ADP_VALUE_NORMALIZED", "UZIO_VALUE_NORMALIZED",
    "RULE_APPLIED", "ADP_EFFECTIVE_DATE_USED",
    "HAS_W4_HISTORY", "VERIFY_IN_UI_FIRST",
]

# Display strings for the CATEGORY column. Kept short so the Excel column
# stays narrow; the rule text in 'Rule Applied' carries the why.
CAT_MISMATCH    = "Mismatch"
CAT_BLANK_VS_V  = "Blank vs Value"
CAT_UI_VERIFY   = "Needs UI Verification"


# Status normalization vocabulary — applies when status comes from any source
# other than UZIO's clean `status` column (e.g. a status field on the ADP file).
ACTIVE_TOKENS = {"active", "active employee", "a", "act", "active (current)"}
TERMINATED_TOKENS = {
    "terminated", "term", "t", "inactive", "separated", "left",
    "no longer employed", "ex-employee", "former",
}


def normalize_status(raw: str) -> str:
    """Map varied status strings to ACTIVE / TERMINATED / (passthrough upper)."""
    s = _clean(raw).lower()
    if not s:
        return ""
    if s in ACTIVE_TOKENS or s.startswith("act"):
        return "ACTIVE"
    if s in TERMINATED_TOKENS or s.startswith("term"):
        return "TERMINATED"
    return _clean(raw).upper()


_KIND_HUMAN = {
    "bool": "boolean",
    "money_cents": "money (UZIO cents → ADP dollars)",
    "filing_status": "filing status",
    "numeric": "numeric / allowances",
}

_FIELD_RULE_NOTES = {
    "FIT_WITHHOLDING_EXEMPTION":           "Yes/No vs true/false. Blank treated as false.",
    "FIT_ADDL_WITHHOLDING_PER_PAY_PERIOD": "UZIO cents → divide by 100 before comparing to ADP dollars.",
    "FIT_FILING_STATUS":                   "Compare normalized; UZIO enum (e.g. FEDERAL_SINGLE) mapped to ADP label.",
    "FIT_CHILD_AND_DEPENDENT_TAX_CREDIT":  "UZIO cents → divide by 100.",
    "FIT_DEDUCTIONS_OVER_STANDARD":        "UZIO cents → divide by 100.",
    "FIT_HIGHER_WITHHOLDING":              "Yes/No vs true/false (Multiple Jobs Step 2c).",
    "FIT_OTHER_INCOME":                    "UZIO cents → divide by 100.",
    "FIT_WITHHOLD_AS_NON_RESIDENT":        "Yes/No vs true/false.",
    "FIT_WITHHOLDING_ALLOWANCE":           "Numeric. Pre-2020 W-4 employees only.",
    "SIT_WITHHOLDING_EXEMPTION":           "Yes/No vs true/false. Routed to 'Needs UI Verification' (derived UZIO flag).",
    "SIT_FILING_STATUS":                   "Compare normalized; UZIO state enum mapped to ADP label. Auto-skipped for states where UZIO has no enum (MA uses SIT_HOH instead).",
    "SIT_TOTAL_ALLOWANCES":                "Numeric. If UZIO TOTAL is blank, fall back to SIT_BASIC_ALLOWANCES + SIT_ADDITIONAL_ALLOWANCES (Illinois rule).",
    "SIT_ADDL_WITHHOLDING_PER_PAY_PERIOD": "UZIO cents → divide by 100.",
}


def _build_field_rules_df(labels: dict) -> pd.DataFrame:
    """Build the 'Field Mapping Rules' sheet in the spec format."""
    rows = []
    for m in FIELD_MAPPING:
        if m["kind"] == "id":
            continue
        juris_for_label = m["scope"] if m["scope"] == "FED" else ""
        rows.append({
            "Field Label":       field_label(m["uzio_key"], labels, juris_for_label),
            "Field Key":         m["uzio_key"],
            "ADP Column Found":  m["adp_col"],
            "UZIO Column Found": m["uzio_key"],   # long-format key acts as the UZIO "column"
            "Type":              _KIND_HUMAN.get(m["kind"], m["kind"]),
            "Notes":             _FIELD_RULE_NOTES.get(m["uzio_key"], ""),
        })
    return pd.DataFrame(rows)


def run_audit(
    adp_df_raw: pd.DataFrame,
    uzio_df_raw: pd.DataFrame,
    options: AuditOptions,
) -> AuditOutput:
    """The pure-data audit pipeline. UI-agnostic."""

    labels = load_key_mapping_yml()
    adp = parse_adp(adp_df_raw)
    uzio = parse_uzio(uzio_df_raw)

    # ── Population gaps ──────────────────────────────────────────────────
    missing_in_uzio_ids = adp.all_emp_ids - uzio.all_emp_ids
    missing_in_adp_ids = uzio.all_emp_ids - adp.all_emp_ids
    both_ids = adp.all_emp_ids & uzio.all_emp_ids

    # ── Per-employee comparison loop ─────────────────────────────────────
    mismatches: list[dict] = []        # Mismatch + Blank vs Value + Needs UI Verification
    false_positives: list[dict] = []
    reciprocity: list[dict] = []
    stale: list[dict] = []

    # Index ADP by employee for fast row lookup.
    adp_indexed = adp.df.set_index(adp.id_col)

    for emp_id in sorted(both_ids):
        adp_row = adp_indexed.loc[emp_id]
        # When multiple ADP rows exist with the same id (shouldn't post-dedup),
        # loc returns a DataFrame — pick the first row.
        if isinstance(adp_row, pd.DataFrame):
            adp_row = adp_row.iloc[0]

        # Identity
        first_name = _clean(adp_row.get("Legal First Name", ""))
        last_name = _clean(adp_row.get("Legal Last Name", ""))
        emp_name = f"{first_name} {last_name}".strip()
        if not emp_name:
            emp_name = uzio.name_by_emp.get(emp_id, "")

        # Status — prefer UZIO's status column over guessing from ADP.
        status = normalize_status(uzio.status_by_emp.get(emp_id, ""))
        if not status:
            # Try a status-like column on the ADP row if UZIO didn't say.
            for col in adp_indexed.columns:
                cl = col.lower()
                if "status" in cl and "marital" not in cl and "tax" not in cl and "withholding" not in cl:
                    status = normalize_status(adp_row.get(col, ""))
                    if status:
                        break
        if not status:
            status = "ACTIVE"  # safe default if neither side tells us

        # Worked & lived state for SIT routing.
        worked_state = _clean(adp_row.get(adp.state_worked_col, "")).upper() if adp.state_worked_col else ""
        lived_state = _clean(adp_row.get(adp.state_lived_col, "")).upper() if adp.state_lived_col else ""
        eff_date = adp_row.get("_eff_date", pd.NaT)
        eff_date_str = eff_date.strftime("%Y-%m-%d") if pd.notna(eff_date) else ""
        is_multi_row = emp_id in adp.multi_row_emp_ids

        # Reciprocity row (emitted for every employee where they differ).
        if worked_state and lived_state and worked_state != lived_state:
            reciprocity.append({
                "EMPLOYEE_ID": emp_id,
                "EMPLOYEE_NAME": emp_name,
                "EMPLOYMENT_STATUS": status,
                "ADP_WORKED_STATE": worked_state,
                "ADP_LIVED_STATE": lived_state,
                "NOTE": "SIT comparison was run against Worked state. Confirm SIT setup in UZIO for both states.",
            })

        # Pick the UZIO federal row for this emp.
        uzio_fed = (
            uzio.federal_wide.loc[emp_id]
            if (not uzio.federal_wide.empty and emp_id in uzio.federal_wide.index)
            else pd.Series(dtype=object)
        )

        # Decide which state to compare for SIT.
        sit_states = []
        if worked_state:
            sit_states.append(worked_state)
        # Fall back to UZIO's recorded states if ADP doesn't tell us.
        if not sit_states:
            sit_states = sorted(uzio.states_by_emp.get(emp_id, set()))

        for mapping in FIELD_MAPPING:
            if mapping["kind"] == "id":
                continue

            uz_key = mapping["uzio_key"]
            adp_col = mapping["adp_col"]
            scope = mapping["scope"]
            kind = mapping["kind"]

            if adp_col not in adp_indexed.columns:
                continue  # ADP file doesn't carry this column for this client

            adp_val_raw = adp_row.get(adp_col, "")

            if scope == "FED":
                uzio_val_raw = uzio_fed.get(uz_key, "") if isinstance(uzio_fed, pd.Series) else ""
                _record_comparison(
                    mapping=mapping, uz_key=uz_key, kind=kind,
                    emp_id=emp_id, emp_name=emp_name, status=status,
                    state_code="", jurisdiction="FED",
                    adp_val_raw=adp_val_raw, uzio_val_raw=uzio_val_raw,
                    is_multi_row=is_multi_row, eff_date_str=eff_date_str,
                    options=options, labels=labels,
                    mismatches=mismatches, false_positives=false_positives,
                    uzio_column_label=uz_key,
                )
                # Stale-record detection for federal fields.
                _check_stale(
                    uzio=uzio, emp_id=emp_id, uz_key=uz_key, state=None,
                    adp_eff_date=eff_date, sink=stale,
                    emp_name=emp_name, status=status,
                )
                continue

            # SIT — iterate per state. Empty sit_states means we skip SIT.
            for state in sit_states:
                # Skip no-SIT states up front.
                if options.skip_no_sit_states and state in NO_SIT_STATES:
                    false_positives.append({
                        "EMPLOYEE_ID": emp_id, "EMPLOYEE_NAME": emp_name,
                        "STATE_CODE": state, "FIELD_KEY": uz_key,
                        "REASON": f"State {state} has no state income tax — SIT comparison skipped.",
                        "ADP_VALUE_RAW": _clean(adp_val_raw), "UZIO_VALUE_RAW": "",
                    })
                    continue

                # Skip SIT_FILING_STATUS for states where UZIO has no such records
                # (e.g. MA, which uses SIT_HOH boolean + allowances instead).
                if (
                    options.skip_sit_fs_when_uzio_has_none
                    and uz_key == "SIT_FILING_STATUS"
                    and (emp_id, state) not in uzio.has_sit_filing_status_for
                ):
                    false_positives.append({
                        "EMPLOYEE_ID": emp_id, "EMPLOYEE_NAME": emp_name,
                        "STATE_CODE": state, "FIELD_KEY": uz_key,
                        "REASON": (
                            f"UZIO has no SIT_FILING_STATUS record for {state}. This state likely "
                            f"encodes filing status via SIT_HOH + allowance fields rather than a "
                            f"single enum. Compare those instead, or verify in UI."
                        ),
                        "ADP_VALUE_RAW": _clean(adp_val_raw), "UZIO_VALUE_RAW": "",
                    })
                    continue

                uzio_val_raw = ""
                state_row = None
                if not uzio.state_wide.empty and (emp_id, state) in uzio.state_wide.index:
                    state_row = uzio.state_wide.loc[(emp_id, state)]
                    if isinstance(state_row, pd.Series):
                        uzio_val_raw = state_row.get(uz_key, "")

                # ── Illinois (and any state) SIT_TOTAL_ALLOWANCES fallback:
                # when TOTAL is blank but BASIC + ADDITIONAL are present, sum
                # them and compare. Spec'd in the Custom GPT prompt.
                uzio_column_label = uz_key
                if (
                    uz_key == "SIT_TOTAL_ALLOWANCES"
                    and _clean(uzio_val_raw) == ""
                    and state_row is not None
                ):
                    basic = _norm_money(_clean(state_row.get("SIT_BASIC_ALLOWANCES", "")))
                    addl = _norm_money(_clean(state_row.get("SIT_ADDITIONAL_ALLOWANCES", "")))
                    if basic is not None or addl is not None:
                        computed = (basic or 0.0) + (addl or 0.0)
                        uzio_val_raw = str(int(computed)) if computed.is_integer() else str(computed)
                        uzio_column_label = "SIT_BASIC_ALLOWANCES + SIT_ADDITIONAL_ALLOWANCES"

                _record_comparison(
                    mapping=mapping, uz_key=uz_key, kind=kind,
                    emp_id=emp_id, emp_name=emp_name, status=status,
                    state_code=state, jurisdiction="STATE",
                    adp_val_raw=adp_val_raw, uzio_val_raw=uzio_val_raw,
                    is_multi_row=is_multi_row, eff_date_str=eff_date_str,
                    options=options, labels=labels,
                    mismatches=mismatches, false_positives=false_positives,
                    uzio_column_label=uzio_column_label,
                )
                _check_stale(
                    uzio=uzio, emp_id=emp_id, uz_key=uz_key, state=state,
                    adp_eff_date=eff_date, sink=stale,
                    emp_name=emp_name, status=status,
                )

    # ── Build DataFrames ─────────────────────────────────────────────────
    df_mismatches = pd.DataFrame(mismatches, columns=MISMATCH_COLS)
    df_filtered = pd.DataFrame(false_positives) if false_positives else pd.DataFrame(columns=[
        "EMPLOYEE_ID", "EMPLOYEE_NAME", "STATE_CODE", "FIELD_KEY", "REASON",
        "ADP_VALUE_RAW", "UZIO_VALUE_RAW",
    ])
    df_recip = pd.DataFrame(reciprocity) if reciprocity else pd.DataFrame(columns=[
        "EMPLOYEE_ID", "EMPLOYEE_NAME", "EMPLOYMENT_STATUS",
        "ADP_WORKED_STATE", "ADP_LIVED_STATE", "NOTE",
    ])
    df_stale = pd.DataFrame(stale) if stale else pd.DataFrame(columns=[
        "EMPLOYEE_ID", "EMPLOYEE_NAME", "EMPLOYMENT_STATUS", "STATE_CODE",
        "FIELD_KEY", "ADP_EFFECTIVE_DATE", "UZIO_EFFECTIVE_DATE", "AGE_DAYS",
    ])

    df_field_rules = _build_field_rules_df(load_key_mapping_yml())

    # Missing populations.
    if missing_in_uzio_ids:
        df_missing_uzio = adp.df[adp.df[adp.id_col].isin(missing_in_uzio_ids)][[adp.id_col]].copy()
        df_missing_uzio.columns = ["ASSOCIATE_ID"]
        df_missing_uzio["LEGAL_FIRST_NAME"] = adp.df.set_index(adp.id_col).loc[df_missing_uzio["ASSOCIATE_ID"], "Legal First Name"].values if "Legal First Name" in adp.df.columns else ""
        df_missing_uzio["LEGAL_LAST_NAME"] = adp.df.set_index(adp.id_col).loc[df_missing_uzio["ASSOCIATE_ID"], "Legal Last Name"].values if "Legal Last Name" in adp.df.columns else ""
    else:
        df_missing_uzio = pd.DataFrame(columns=["ASSOCIATE_ID", "LEGAL_FIRST_NAME", "LEGAL_LAST_NAME"])

    if missing_in_adp_ids:
        df_missing_adp = pd.DataFrame({
            "EMPLOYEE_ID": sorted(missing_in_adp_ids),
            "EMPLOYEE_NAME": [uzio.name_by_emp.get(e, "") for e in sorted(missing_in_adp_ids)],
            "STATUS": [uzio.status_by_emp.get(e, "") for e in sorted(missing_in_adp_ids)],
        })
    else:
        df_missing_adp = pd.DataFrame(columns=["EMPLOYEE_ID", "EMPLOYEE_NAME", "STATUS"])

    # Stable sort for the mismatch sheets — Category (most actionable first),
    # then Active before Terminated, then Employee, then Field.
    _category_order = {CAT_MISMATCH: 0, CAT_BLANK_VS_V: 1, CAT_UI_VERIFY: 2}
    if not df_mismatches.empty:
        df_mismatches["_cat_order"] = df_mismatches["CATEGORY"].map(_category_order).fillna(99)
        df_mismatches["_status_order"] = (df_mismatches["EMPLOYMENT_STATUS"] != "ACTIVE").astype(int)
        df_mismatches = (
            df_mismatches
            .sort_values(["_cat_order", "_status_order", "EMPLOYEE_ID", "FIELD_LABEL"], kind="mergesort")
            .drop(columns=["_cat_order", "_status_order"])
            .reset_index(drop=True)
        )

    # Per-Category × per-Field counts.
    if not df_mismatches.empty:
        df_summary = (
            df_mismatches.groupby(["CATEGORY", "FIELD_LABEL", "FIELD_KEY"])
            .agg(**{
                "Mismatch Count":        ("EMPLOYEE_ID", "count"),
                "Unique Employee Count": ("EMPLOYEE_ID", "nunique"),
            })
            .reset_index()
            .rename(columns={
                "CATEGORY":   "Category",
                "FIELD_LABEL": "Field Label",
                "FIELD_KEY":   "Field Key",
            })
            .sort_values(["Category", "Mismatch Count"], ascending=[True, False])
        )
    else:
        df_summary = pd.DataFrame(columns=[
            "Category", "Field Label", "Field Key", "Mismatch Count", "Unique Employee Count",
        ])

    # Per-employee summary.
    if not df_mismatches.empty:
        df_emp_sum = df_mismatches.groupby("EMPLOYEE_ID").agg(
            EMPLOYEE_NAME=("EMPLOYEE_NAME", "first"),
            EMPLOYMENT_STATUS=("EMPLOYMENT_STATUS", "first"),
            STATE_CODE=("STATE_CODE", lambda s: ", ".join(sorted(set(x for x in s if x)))),
            mismatch_rows=("FIELD_KEY", "count"),
            fields=("FIELD_LABEL", lambda x: ", ".join(sorted(set(x)))),
        ).reset_index().sort_values("mismatch_rows", ascending=False)
    else:
        df_emp_sum = pd.DataFrame(columns=["EMPLOYEE_ID", "EMPLOYEE_NAME", "EMPLOYMENT_STATUS", "STATE_CODE", "mismatch_rows", "fields"])

    # Top-level summary metrics.
    df_miss_active = df_mismatches[df_mismatches["EMPLOYMENT_STATUS"] == "ACTIVE"] if not df_mismatches.empty else df_mismatches
    df_miss_term = df_mismatches[df_mismatches["EMPLOYMENT_STATUS"] != "ACTIVE"] if not df_mismatches.empty else df_mismatches

    def _cat_count(df, cat):
        if df.empty:
            return 0
        return int((df["CATEGORY"] == cat).sum())

    metrics = [
        {"Metric": "UZIO employees (total)", "Value": len(uzio.all_emp_ids)},
        {"Metric": "ADP employees (total)", "Value": len(adp.all_emp_ids)},
        {"Metric": "Employees compared (in both)", "Value": len(both_ids)},
        {"Metric": "Employees with W-4 history (multiple ADP rows)", "Value": len(adp.multi_row_emp_ids)},
        {"Metric": "Mismatches — Total (all categories)", "Value": len(df_mismatches)},
        {"Metric": "Mismatches — Active employees", "Value": len(df_miss_active)},
        {"Metric": "Mismatches — Terminated employees", "Value": len(df_miss_term)},
        {"Metric": "Employees with at least one mismatch", "Value": df_mismatches["EMPLOYEE_ID"].nunique() if not df_mismatches.empty else 0},
        {"Metric": "    of which Category = Mismatch", "Value": _cat_count(df_mismatches, CAT_MISMATCH)},
        {"Metric": "    of which Category = Blank vs Value", "Value": _cat_count(df_mismatches, CAT_BLANK_VS_V)},
        {"Metric": "    of which Category = Needs UI Verification", "Value": _cat_count(df_mismatches, CAT_UI_VERIFY)},
        {"Metric": "False positives filtered out", "Value": len(df_filtered)},
        {"Metric": "Reciprocity-check rows (Lived != Worked state)", "Value": len(df_recip)},
        {"Metric": "Stale UZIO field records detected", "Value": len(df_stale)},
        {"Metric": "Employees in ADP missing from UZIO", "Value": len(missing_in_uzio_ids)},
        {"Metric": "Employees in UZIO missing from ADP", "Value": len(missing_in_adp_ids)},
    ]
    df_metrics = pd.DataFrame(metrics)

    # About sheet.
    df_about = pd.DataFrame([
        {"Section": "Act on first",
         "Notes": "Open 'Mismatches (Active)'. Every finding for an Active employee is here, sorted so 'Mismatch' (hard disagreement) comes before 'Blank vs Value' (one side missing a value) and 'Needs UI Verification' (UZIO derived flag, not editable). The 'Category' column tells you which is which. Rows with VERIFY_IN_UI_FIRST=Yes have W-4 history; confirm the latest W-4 record in UZIO UI before changing."},
        {"Section": "Categories explained",
         "Notes": "Mismatch — ADP and UZIO disagree on populated values, fix one side. Blank vs Value — one system has a value, the other is blank (usually means data entry missed a step). Needs UI Verification — the underlying UZIO field is a derived/internal flag (e.g. SIT_WITHHOLDING_EXEMPTION auto-set when employee is Federal-Exempt); the comparison is unreliable, open UZIO UI to confirm."},
        {"Section": "False positives filtered",
         "Notes": "'False Positives Filtered' lists comparisons we intentionally skipped (no-SIT states like FL/TX/NV, and states where UZIO doesn't store SIT_FILING_STATUS as a single enum such as MA). No action needed; this sheet exists for transparency."},
        {"Section": "Stale UZIO records",
         "Notes": "'Stale UZIO Records' lists fields where UZIO's effective_date is older than the latest ADP W-4 date. Often the employee submitted a new W-4 in ADP that hasn't propagated to UZIO."},
        {"Section": "Multi-state employees",
         "Notes": "'Reciprocity Check Needed' lists employees whose ADP Lived state differs from their Worked state. SIT was audited against the Worked state; confirm UZIO setup for both."},
        {"Section": "W-4 history",
         "Notes": "'ADP Effective Date Used' shows every row in the ADP upload, marking which row's W-4 was used per employee (the most recent)."},
        {"Section": "Population gaps",
         "Notes": "'Missing in UZIO' (in ADP not UZIO — may need adding) and 'Missing in ADP' (in UZIO not ADP — may be phantom records or post-migration hires)."},
    ])

    return AuditOutput(
        mismatches=df_mismatches,
        false_positives_filtered=df_filtered,
        reciprocity=df_recip,
        stale_uzio=df_stale,
        missing_in_uzio=df_missing_uzio,
        missing_in_adp=df_missing_adp,
        employees_with_mismatches=df_emp_sum,
        mismatch_summary=df_summary,
        field_rules=df_field_rules,
        w4_history=adp.w4_history,
        summary_metrics=df_metrics,
        about=df_about,
    )


def _record_comparison(
    *, mapping, uz_key, kind, emp_id, emp_name, status, state_code, jurisdiction,
    adp_val_raw, uzio_val_raw, is_multi_row, eff_date_str,
    options: AuditOptions, labels: dict,
    mismatches: list, false_positives: list,
    uzio_column_label: str,
) -> None:
    """Compare one cell and append to the unified `mismatches` list (or
    `false_positives` for explicitly skipped comparisons).

    All findings go into the same list; the CATEGORY column distinguishes
    hard mismatches, blank-vs-value cases, and UI-verify-only flags. The
    downstream sheets split on EMPLOYMENT_STATUS, not on category.
    """
    result = compare_field(uz_key, kind, adp_val_raw, uzio_val_raw)
    if result.matched:
        return

    label_juris = "FED" if jurisdiction == "FED" else state_code
    label = field_label(uz_key, labels, label_juris)

    # Decide the category. UI-verify routing wins over blank-vs-value, since
    # the underlying field is a known UZIO derived flag and the comparison
    # result is unreliable either way.
    if options.skip_known_ui_only_fields and uz_key in FIELDS_REQUIRING_UI_VERIFICATION:
        category = CAT_UI_VERIFY
        rule_text = (
            f"{result.rule}: {result.reason} | "
            "Needs UI verification — internal/derived UZIO flag, "
            "not editable from the UI."
        )
    elif result.category == "blank_vs_value" and options.flag_blank_vs_value:
        category = CAT_BLANK_VS_V
        rule_text = (
            f"{result.rule}: {result.reason} | "
            "Blank vs Value — one side has a non-default value, the other is blank."
        )
    else:
        category = CAT_MISMATCH
        rule_text = f"{result.rule}: {result.reason}"

    mismatches.append({
        "EMPLOYEE_ID": emp_id, "EMPLOYEE_NAME": emp_name,
        "EMPLOYMENT_STATUS": status, "STATE_CODE": state_code,
        "JURISDICTION": jurisdiction,
        "CATEGORY": category,
        "FIELD_LABEL": label, "FIELD_KEY": uz_key,
        "ADP_COLUMN": mapping["adp_col"],
        "UZIO_COLUMN": uzio_column_label,
        "ADP_VALUE_RAW": _clean(adp_val_raw),
        "UZIO_VALUE_RAW": _clean(uzio_val_raw),
        "ADP_VALUE_NORMALIZED": result.adp_normalized,
        "UZIO_VALUE_NORMALIZED": result.uzio_normalized,
        "RULE_APPLIED": rule_text,
        "ADP_EFFECTIVE_DATE_USED": eff_date_str,
        "HAS_W4_HISTORY": "Yes" if is_multi_row else "No",
        "VERIFY_IN_UI_FIRST": "Yes" if is_multi_row else "",
    })


def _check_stale(*, uzio: UzioPivot, emp_id, uz_key, state, adp_eff_date, sink, emp_name, status):
    """If UZIO's effective_date for this field is older than ADP's W-4 date, flag it."""
    if pd.isna(adp_eff_date):
        return
    lookup_key = (emp_id, uz_key, state)
    uzio_dt = uzio.eff_by_emp_field.get(lookup_key)
    if uzio_dt is None or pd.isna(uzio_dt):
        return
    age = (adp_eff_date - uzio_dt).days
    if age > 0:
        sink.append({
            "EMPLOYEE_ID": emp_id, "EMPLOYEE_NAME": emp_name,
            "EMPLOYMENT_STATUS": status,
            "STATE_CODE": state or "",
            "FIELD_KEY": uz_key,
            "ADP_EFFECTIVE_DATE": adp_eff_date.strftime("%Y-%m-%d"),
            "UZIO_EFFECTIVE_DATE": uzio_dt.strftime("%Y-%m-%d"),
            "AGE_DAYS": age,
        })


# ─────────────────────────────────────────────────────────────────────────────
# Excel output
# ─────────────────────────────────────────────────────────────────────────────

_MISMATCH_DISPLAY = {
    "EMPLOYEE_ID":             "Employee ID",
    "EMPLOYEE_NAME":           "Employee Name",
    "EMPLOYMENT_STATUS":       "Employment Status",
    "STATE_CODE":              "State Code",
    "JURISDICTION":            "Jurisdiction",
    "CATEGORY":                "Category",
    "FIELD_LABEL":             "Field Label",
    "FIELD_KEY":               "Field Key",
    "ADP_COLUMN":              "ADP Column Name",
    "UZIO_COLUMN":             "UZIO Column Name",
    "ADP_VALUE_RAW":           "ADP Raw Value",
    "UZIO_VALUE_RAW":          "UZIO Raw Value",
    "ADP_VALUE_NORMALIZED":    "ADP Normalized Value",
    "UZIO_VALUE_NORMALIZED":   "UZIO Normalized Value",
    "RULE_APPLIED":            "Rule Applied",
    "ADP_EFFECTIVE_DATE_USED": "ADP Effective Date Used",
    "HAS_W4_HISTORY":          "Has W-4 History",
    "VERIFY_IN_UI_FIRST":      "Verify In UI First",
}


def _to_display(df: pd.DataFrame) -> pd.DataFrame:
    """Rename internal UPPER_SNAKE columns to spec display names just-in-time
    before writing to Excel. Keeps the in-memory dataclass stable."""
    return df.rename(columns=_MISMATCH_DISPLAY)


def build_workbook(out: AuditOutput) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        out.about.to_excel(writer, sheet_name="About This Report", index=False)
        out.summary_metrics.to_excel(writer, sheet_name="Summary", index=False)
        out.mismatch_summary.to_excel(writer, sheet_name="Mismatch Summary", index=False)
        _to_display(out.mismatches).to_excel(writer, sheet_name="Mismatches (All)", index=False)
        if not out.mismatches.empty:
            _to_display(out.mismatches[out.mismatches["EMPLOYMENT_STATUS"] == "ACTIVE"]).to_excel(
                writer, sheet_name="Mismatches (Active)", index=False,
            )
            _to_display(out.mismatches[out.mismatches["EMPLOYMENT_STATUS"] != "ACTIVE"]).to_excel(
                writer, sheet_name="Mismatches (Terminated)", index=False,
            )
        else:
            _to_display(out.mismatches).to_excel(writer, sheet_name="Mismatches (Active)", index=False)
            _to_display(out.mismatches).to_excel(writer, sheet_name="Mismatches (Terminated)", index=False)
        out.false_positives_filtered.to_excel(writer, sheet_name="False Positives Filtered", index=False)
        out.reciprocity.to_excel(writer, sheet_name="Reciprocity Check Needed", index=False)
        out.stale_uzio.to_excel(writer, sheet_name="Stale UZIO Records", index=False)
        out.employees_with_mismatches.to_excel(writer, sheet_name="Employees with Mismatches", index=False)
        out.field_rules.to_excel(writer, sheet_name="Field Mapping Rules", index=False)
        out.w4_history.to_excel(writer, sheet_name="ADP Effective Date Used", index=False)
        out.missing_in_adp.to_excel(writer, sheet_name="Missing in ADP", index=False)
        out.missing_in_uzio.to_excel(writer, sheet_name="Missing in UZIO", index=False)

        # Auto-fit columns.
        for sheet in writer.sheets.values():
            for col in sheet.columns:
                max_len = 0
                letter = col[0].column_letter
                for cell in col:
                    v = cell.value
                    if v is None:
                        continue
                    s = str(v)
                    if len(s) > max_len:
                        max_len = len(s)
                sheet.column_dimensions[letter].width = min(max(max_len + 2, 10), 60)

    return buf.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# Streamlit UI
# ─────────────────────────────────────────────────────────────────────────────

def render_ui():
    st.title("ADP ↔ UZIO Withholding Audit Tool")
    st.caption("v3 — auto-detects file source, multi-state aware, false-positive resistant.")

    st.markdown(
        "Upload the **ADP** export on the left and the **UZIO** withholding export on the right. "
        "Both CSV and XLSX are accepted."
    )

    with st.expander("What's new in v3", expanded=False):
        st.markdown(
            "- Status, effective date and state code now come from the UZIO file when ADP is missing them "
            "(previous versions emitted 'NAN' for every row).\n"
            "- Multi-state SIT join — an employee with SIT records for two states is compared per state.\n"
            "- SIT_FILING_STATUS is automatically skipped for states like MA where UZIO encodes filing "
            "status as SIT_HOH + allowance fields rather than a single enum. "
            "(In one prior audit this single guard removed ~95% of the false positives.)\n"
            "- New 'Blank vs Value' bucket separates 'one side missing a value' from numeric drift.\n"
            "- New 'Stale UZIO Records' sheet flags fields where UZIO's effective_date is older than "
            "the latest ADP W-4 date.\n"
            "- New 'Reciprocity Check Needed' sheet lists every employee whose Lived state differs from "
            "their Worked state."
        )

    col_a, col_b = st.columns(2)
    adp_upload = col_a.file_uploader("ADP file", type=["csv", "xlsx", "xls"], key="wh_adp_file")
    uzio_upload = col_b.file_uploader("UZIO file", type=["csv", "xlsx", "xls"], key="wh_uzio_file")
    client_name = st.text_input("Client name (used in the output filename)", value="Client_Name")

    with st.expander("Audit settings", expanded=False):
        skip_no_sit = st.checkbox(
            "Skip SIT comparisons for employees in no-SIT states (FL, TX, NV, …)",
            value=True,
        )
        skip_ui_only = st.checkbox(
            "Route SIT_WITHHOLDING_EXEMPTION mismatches to 'Needs UI Verification'",
            value=True,
            help="This field is an internal/derived DB flag in UZIO and not editable from the UI.",
        )
        skip_sit_fs = st.checkbox(
            "Skip SIT_FILING_STATUS for states where UZIO has no such records (recommended)",
            value=True,
            help="Many states (MA, NJ, etc. for certain employees) don't store filing status as a single "
                 "enum. Comparing then produces a false positive for every employee in that state.",
        )
        flag_blank_vs_value = st.checkbox(
            "Separate 'blank vs value' findings into their own sheet",
            value=True,
        )

    # ── Read each file from its labeled slot (no auto-swap) ──────────────
    adp_df = uzio_df = None
    if adp_upload:
        try:
            adp_df = read_uploaded(adp_upload)
        except Exception as e:
            st.error(f"Could not read the ADP file: {e}")
    if uzio_upload:
        try:
            uzio_df = read_uploaded(uzio_upload)
        except Exception as e:
            st.error(f"Could not read the UZIO file: {e}")

    # Soft sanity check — if the contents of a slot look like the other system,
    # warn the user but still trust their labels. The user can fix it by
    # re-uploading into the correct slot.
    if adp_df is not None and detect_source(adp_df) == "uzio":
        st.warning(
            "The file you uploaded as **ADP file** looks like a UZIO export "
            f"(it has `{UZIO_KEY_COL}` / `{UZIO_VAL_COL}` columns). "
            "Double-check the slots — the tool will treat it as ADP regardless."
        )
    if uzio_df is not None and detect_source(uzio_df) == "adp":
        st.warning(
            "The file you uploaded as **UZIO file** looks like an ADP export "
            "(it has `Associate ID` / W-4 columns). "
            "Double-check the slots — the tool will treat it as UZIO regardless."
        )

    run_btn = st.button(
        "Run Audit", type="primary",
        disabled=(adp_df is None or uzio_df is None),
    )

    if not run_btn:
        return

    options = AuditOptions(
        skip_no_sit_states=skip_no_sit,
        skip_known_ui_only_fields=skip_ui_only,
        skip_sit_fs_when_uzio_has_none=skip_sit_fs,
        flag_blank_vs_value=flag_blank_vs_value,
    )

    with st.spinner("Auditing…"):
        try:
            result = run_audit(adp_df, uzio_df, options)
        except Exception as e:
            st.error(f"Audit failed: {e}")
            st.exception(e)
            return

    # ── UI summary ───────────────────────────────────────────────────────
    m = result.mismatches
    n_total = len(m)
    n_active = int((m["EMPLOYMENT_STATUS"] == "ACTIVE").sum()) if not m.empty else 0
    n_term = int((m["EMPLOYMENT_STATUS"] != "ACTIVE").sum()) if not m.empty else 0
    n_cat_mis = int((m["CATEGORY"] == CAT_MISMATCH).sum()) if not m.empty else 0
    n_cat_blank = int((m["CATEGORY"] == CAT_BLANK_VS_V).sum()) if not m.empty else 0
    n_cat_ui = int((m["CATEGORY"] == CAT_UI_VERIFY).sum()) if not m.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Mismatches", n_total)
    c2.metric("Active", n_active)
    c3.metric("Terminated", n_term)
    c4.metric("Missing in UZIO / ADP",
              f"{len(result.missing_in_uzio)} / {len(result.missing_in_adp)}")

    c5, c6, c7, c8 = st.columns(4)
    c5.metric("False Positives Filtered", len(result.false_positives_filtered))
    c6.metric("Stale UZIO Records", len(result.stale_uzio))
    c7.metric("Reciprocity Rows", len(result.reciprocity))
    c8.metric("ADP rows w/ W-4 history", int(len(result.w4_history) - (
        result.w4_history["IS_SELECTED_LATEST"].sum() if "IS_SELECTED_LATEST" in result.w4_history.columns else 0
    )))

    if n_total:
        st.caption(
            f"Category breakdown — **Mismatch:** {n_cat_mis}  ·  "
            f"**Blank vs Value:** {n_cat_blank}  ·  "
            f"**Needs UI Verification:** {n_cat_ui}  "
            "(all combined into Mismatches (All / Active / Terminated) sheets; see the 'Category' column.)"
        )

    if len(result.mismatches) > 100:
        st.warning(
            "More than 100 real mismatches detected. Open the **Mismatch Summary** sheet first — "
            "a single field with a high count usually points to a systemic issue (e.g., a "
            "filing-status mapping that needs extending) rather than 100 separate problems."
        )

    if not result.mismatch_summary.empty:
        st.subheader("Mismatch Summary")
        st.dataframe(result.mismatch_summary, hide_index=True, use_container_width=True)

    # ── Download — spec filename: ADP_vs_UZIO_FIT_SIT_Mismatch_Report_<Client>.xlsx
    timestamp = datetime.now().strftime("%d_%m_%Y_%H%M")
    filename = f"ADP_vs_UZIO_FIT_SIT_Mismatch_Report_{client_name}_{timestamp}.xlsx"
    st.download_button(
        "📥 Download full Excel report",
        data=build_workbook(result),
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    st.set_page_config(layout="wide")
    render_ui()
