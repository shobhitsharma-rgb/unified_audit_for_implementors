import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
from utils.audit_utils import (
    read_uzio_raw_file, generate_uzio_template,
    HOURLY_ONLY_JOB_TITLES, is_hourly_only_job_title,
    norm_colname, norm_blank, try_parse_date, ensure_unique_columns, safe_val, normalize_space_and_case,
    as_float_or_none, find_col, get_identity_match_map, norm_ssn_canonical, detect_duplicate_ssns, norm_id
)

# =========================================================
# Data_Audit_Tool (Streamlit)
# - User uploads Raw Uzio Export (.xlsm) and Raw ADP Export (.xlsx)
# - Hardcoded column mappings
# =========================================================

APP_TITLE = "Census ADP and Uzio Data Review Tool"

# Hardcoded Mapping: Internal Standard Name -> ADP Column Name
ADP_FIELD_MAP = {
    'Employee ID': 'Associate ID',
    'First Name': 'Legal First Name',
    'Last Name': 'Legal Last Name',
    'Middle Initial': 'Legal Middle Name',
    'Suffix': 'Generation Suffix Code',
    'Employment Status': 'Position Status',
    'Employment Type': 'Worker Category Description',
    'Hire Date': 'Hire/Rehire Date',
    'Original Hire Date': 'Hire Date',
    'Termination Date': 'Termination Date',
    'Termination Reason': 'Termination Reason Description',
    'Pay Type': 'Regular Pay Rate Description',
    'Annual Salary': 'Annual Salary',
    'Hourly Pay Rate': 'Regular Pay Rate Amount',
    'Working Hours': 'Standard Hours',
    'Job Title': 'Job Title Description',
    'Department': 'Department Description',
    'Work Email': 'Work Contact: Work Email',
    'Personal Email': 'Personal Contact: Personal Email',
    'SSN': 'Tax ID (SSN)',
    'DOB': 'Birth Date',
    'Gender': 'Gender / Sex (Self-ID)',
    'Tobacco User': 'Tobacco User',
    'FLSA Classification': 'FLSA Description',
    'Address Line 1': 'Primary Address: Address Line 1',
    'Address Line 2': 'Primary Address: Address Line 2',
    'City': 'Primary Address: City',
    'Zip': 'Primary Address: Zip / Postal Code',
    'State': 'Primary Address: State / Territory Code',
    'Mailing Address Line 1': 'Legal / Preferred Address: Address Line 1',
    'Mailing Address Line 2': 'Legal / Preferred Address: Address Line 2',
    'Mailing City': 'Legal / Preferred Address: City',
    'Mailing Zip': 'Legal / Preferred Address: Zip / Postal Code',
    'Mailing State': 'Legal / Preferred Address: State / Territory Code',
    'Reports To ID': 'Reports To Associate ID',
    'Protected Veteran Status': 'Protected Veteran Status',
    'EEO Job Category': 'EEOC Job Classification',
    'Ethnicity': 'Race Description',
    'SOC Code': 'SOC Code',
    'Work Location': 'Location Description'
}

# ---------- Helpers ----------
# (redunant helpers removed, using utils.audit_utils)

def digits_only(x):
    x = norm_blank(x)
    if x == "":
        return ""
    # If pandas read it as a float, it might look like '9048729456.0'
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return re.sub(r"\D", "", s)

def norm_ssn_9digits(x):
    # ONLY CHANGE: SSN compare as 9 digits (pad leading zeros if Excel dropped them)
    d = digits_only(x)
    if d == "":
        return ""
    if len(d) < 9:
        return d.zfill(9)
    if len(d) > 9:
        return d[-9:]
    return d

def norm_zip_first5(x):
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, (int, np.integer)):
        s = str(int(x))
    elif isinstance(x, (float, np.floating)) and float(x).is_integer():
        s = str(int(x))
    else:
        s = re.sub(r"[^\d]", "", str(x).strip())
    if s == "":
        return ""
    if 0 < len(s) < 5:
        s = s.zfill(5)
    return s[:5]

NUMERIC_KEYWORDS = {"salary", "rate", "hours", "amount"}
DATE_KEYWORDS = {"date", "dob", "birth", "doh", "hire"}
SSN_KEYWORDS = {"ssn", "tax id"}
ZIP_KEYWORDS = {"zip", "zipcode", "postal"}
GENDER_KEYWORDS = {"gender"}
PHONE_KEYWORDS = {"phone"}
MIDDLE_INITIAL_KEYWORDS = {"middle initial"}  # ONLY CHANGE: treat as initial vs full middle name
JOB_TITLE_KEYWORDS = {"job title", "position title"}
VETERAN_KEYWORDS = {"veteran"}
EMPLOYMENT_TYPE_KEYWORDS = {"employment type"}

# Valid Uzio Employment Type values
VALID_EMPLOYMENT_TYPES = {"full time", "part time", "other", "seasonal"}

JOB_TITLE_MAPPINGS = {
    "admin": "administrator",
    "management": "manager",
    "dsp owner": "owner"
}

def norm_gender(x):
    x = norm_blank(x)
    if x == "":
        return ""
    s = str(x).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().casefold()
    if "female" in s or "woman" in s:
        return "female"
    if "male" in s or "man" in s:
        return "male"
    return s

def norm_middle_initial(x):
    # ONLY CHANGE: compare middle initial to the first letter of ADP middle name
    x = norm_blank(x)
    if x == "":
        return ""
    s = str(x).strip()
    m = re.search(r"[A-Za-z]", s)
    return (m.group(0).casefold() if m else "")

def norm_job_title(x):
    x = norm_blank(x)
    if x == "":
        return ""
    s = str(x).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().casefold()
    return JOB_TITLE_MAPPINGS.get(s, s)

def norm_veteran_status(x):
    x = norm_blank(x)
    if x == "":
        return ""
    s = str(x).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().casefold()
    
    # Normalize "Decline to answer" variants
    if "decline to self-identify" in s or "decline to answer" in s:
        return "decline to answer"

    # "i am not a protected veteran" -> "not a protected veteran"
    if "not a protected veteran" in s:
        return "not a protected veteran"
    
    # "identify as a protected veteran", "protected veteran" (without 'not') -> "protected veteran"
    if "protected veteran" in s and "not" not in s:
        return "protected veteran"
        
    return s

def norm_value(x, field_name: str):
    f = norm_colname(field_name).lower()
    x = norm_blank(x)
    if x == "":
        return ""

    if any(k in f for k in MIDDLE_INITIAL_KEYWORDS):  # ONLY CHANGE
        return norm_middle_initial(x)

    if any(k in f for k in GENDER_KEYWORDS):
        return norm_gender(x)

    if any(k in f for k in VETERAN_KEYWORDS):
        return norm_veteran_status(x)

    if any(k in f for k in JOB_TITLE_KEYWORDS):
        return norm_job_title(x)

    if any(k in f for k in SSN_KEYWORDS):  # ONLY CHANGE: use 9-digit padded SSN
        return norm_ssn_9digits(x)

    if any(k in f for k in PHONE_KEYWORDS):
        return digits_only(x)

    if any(k in f for k in ZIP_KEYWORDS):
        return norm_zip_first5(x)

    if any(k in f for k in DATE_KEYWORDS):
        return try_parse_date(x)

    if any(k in f for k in NUMERIC_KEYWORDS):
        if isinstance(x, (int, float, np.integer, np.floating)):
            return float(x)
        if isinstance(x, str):
            s = x.strip().replace(",", "").replace("$", "")
            try:
                return float(s)
            except Exception:
                return re.sub(r"\s+", " ", x.strip()).casefold()

    if isinstance(x, str):
        return re.sub(r"\s+", " ", x.strip()).casefold()

    return str(x).casefold()

def norm_emp_key_series(s: pd.Series) -> pd.Series:
    s2 = s.astype(object).where(~s.isna(), "")
    def _fix(v):
        v = str(v).strip()
        v = v.replace("\u00A0", " ")
        if re.fullmatch(r"\d+\.0+", v):
            v = v.split(".")[0]
        return v
    return s2.map(_fix)

# ---------- Rule helpers ----------
def is_termination_reason_field(field_name: str) -> bool:
    return "termination reason" in norm_colname(field_name).casefold()

def is_employment_status_field(field_name: str) -> bool:
    return "employment status" in norm_colname(field_name).casefold()

def status_contains_any(s: str, needles) -> bool:
    s = ("" if s is None else str(s)).casefold()
    return any(n in s for n in needles)

def uzio_is_active(uz_norm: str) -> bool:
    s = ("" if uz_norm is None else str(uz_norm)).casefold()
    return s == "active" or s.startswith("active")

def uzio_is_terminated(uz_norm: str) -> bool:
    s = ("" if uz_norm is None else str(uz_norm)).casefold()
    return s == "terminated" or s.startswith("terminated")

ALLOWED_TERM_REASONS = {
    "quit without notice",
    "no reason given",
    "misconduct",
    "abandoned job",
    "advancement (better job with higher pay)",
    "no-show (never started employment)",
    "performance",
    "personal",
    "scheduling conflicts (schedules don't work)",
    "attendance",
    "reduction in force",
    "reorganization",
    "mutual agreement",
    "import created action",
    "advancement",
    "no-show",
    "management",
    "layoff"
}

def normalize_reason_text(x) -> str:
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).replace("\u00A0", " ")
    s = s.replace("’", "'").replace("“", '"').replace("”", '"')
    s = re.sub(r"\s+", " ", s).strip()
    s = s.strip('"').strip("'")
    return s.casefold()

def normalize_paytype_text(x) -> str:
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s.casefold()

def paytype_bucket(paytype_norm: str) -> str:
    s = ("" if paytype_norm is None else str(paytype_norm)).casefold()
    if "hour" in s:
        return "hourly"
    if "salary" in s or "salaried" in s:
        return "salaried"
    return ""

def is_annual_salary_field(field_name: str) -> bool:
    return "annual salary" in norm_colname(field_name).casefold()

def is_hourly_rate_field(field_name: str) -> bool:
    f = norm_colname(field_name).casefold()
    return ("hourly pay rate" in f) or ("hourly rate" in f)

# ---------- Guardrail: prevent ACTIVE/TERMINATED/RETIRED values leaking into non-status fields ----------
EMP_STATUS_TOKENS = {"active", "terminated", "retired"}

def field_allows_emp_status_value(field_name: str) -> bool:
    f = norm_colname(field_name).casefold()
    return (f == "status") or ("employment status" in f)

def cleanse_uzio_value_for_field(field_name: str, uz_val):
    if norm_blank(uz_val) == "":
        return uz_val
    s = str(uz_val).strip().casefold()
    if (s in EMP_STATUS_TOKENS) and (not field_allows_emp_status_value(field_name)):
        return ""
    return uz_val

# ---------- Pay Type equivalence (UZIO Salaried == ADP Salary) ----------
def is_pay_type_field(field_name: str) -> bool:
    f = norm_colname(field_name).casefold()
    return f == "pay type" or ("pay type" in f)

# ---------- Employment Type (Full Time / Part Time / Other / Seasonal) ----------
def is_employment_type_field(field_name: str) -> bool:
    f = norm_colname(field_name).casefold()
    return "employment type" in f

def normalize_employment_type(x) -> str:
    """Normalize to one of: full time, part time, other, seasonal, or raw value."""
    s = norm_blank(x)
    if s == "":
        return ""
    s = str(s).replace("\u00A0", " ")
    s = re.sub(r"\s+", " ", s).strip().casefold()
    # Map common ADP variants -> Uzio canonical
    if s in {"full time", "fulltime", "full-time", "ft"}:
        return "full time"
    if s in {"part time", "parttime", "part-time", "pt"}:
        return "part time"
    if s in {"seasonal", "temporary", "temp"}:
        return "seasonal"
    if s in {"other"}:
        return "other"
    return s

def normalize_paytype_for_compare(x) -> str:
    s = normalize_paytype_text(x)
    if s in {"salary", "salaried"}:
        return "salaried"
    if s in {"hourly", "hour"}:
        return "hourly"
    return s

def deduplicate_adp(df: pd.DataFrame, key_col: str) -> pd.DataFrame:
    """Refactored Deduplication: Safe against Pandas `.apply` column drops"""
    # Identify special columns (normalized)
    col_map = {c: c.lower() for c in df.columns}
    
    status_col = next((c for c, l in col_map.items() if "position status" in l), None)
    term_date_col = next((c for c, l in col_map.items() if "termination date" in l), None)
    start_date_col = next((c for c, l in col_map.items() if "position start date" in l), None)
    loc_desc_col = next((c for c, l in col_map.items() if "work location description" in l), None)
    license_id_col = next((c for c, l in col_map.items() if "license/certification id" in l), None)
    
    # If we can't find status col, fallback to basic drop_duplicates
    if not status_col:
        return df.drop_duplicates(subset=[key_col], keep="first")
        
    def pick_best_idx(group):
        if len(group) <= 1:
            return group.index[0]
        
        # Helper to parse date for sorting
        def get_date_val(row, col):
            if not col or pd.isna(row[col]):
                return pd.Timestamp.min
            val = str(row[col]).strip()
            if not val:
                return pd.Timestamp.min
            try:
                return pd.to_datetime(val)
            except:
                return pd.Timestamp.min

        group_work = group.copy()
        group_work['__norm_status'] = group_work[status_col].astype(str).str.lower().str.strip()
        
        # Add license check
        if license_id_col:
            group_work['__has_license'] = group_work[license_id_col].apply(lambda x: 1 if norm_blank(x) != "" else 0)
        else:
            group_work['__has_license'] = 0

        actives = group_work[group_work['__norm_status'] == 'active']
        terms = group_work[group_work['__norm_status'] == 'terminated']
        others = group_work[(group_work['__norm_status'] != 'active') & (group_work['__norm_status'] != 'terminated')]
        
        # Logic 1: If Actives exist, prioritize them
        if not actives.empty:
            actives['__sort_date'] = actives.apply(lambda r: get_date_val(r, start_date_col), axis=1)

            if loc_desc_col:
                actives['__has_loc'] = actives[loc_desc_col].apply(lambda x: 1 if norm_blank(x) != "" else 0)
                best_active = actives.sort_values(by=['__has_loc', '__has_license', '__sort_date'], ascending=[False, False, False])
            else:
                best_active = actives.sort_values(by=['__has_license', '__sort_date'], ascending=[False, False])
            
            return best_active.index[0]

        # Logic 2: Terminated
        if not terms.empty:
            terms['__sort_date'] = pd.Timestamp.min
            
            use_start_date = False
            if term_date_col:
                terms['__term_dt_val'] = terms[term_date_col].apply(norm_blank)
                has_blank = (terms['__term_dt_val'] == "").any()
                has_val = (terms['__term_dt_val'] != "").any()
                
                if has_blank and has_val:
                    use_start_date = True
            else:
                use_start_date = True

            if use_start_date:
                 terms['__sort_date'] = terms.apply(lambda r: get_date_val(r, start_date_col), axis=1)
            elif term_date_col:
                 terms['__sort_date'] = terms.apply(lambda r: get_date_val(r, term_date_col), axis=1)
            
            best_term = terms.sort_values(by=['__has_license', '__sort_date'], ascending=[False, False])
            return best_term.index[0]

        # Fallback (Others, e.g. Leave)
        if not others.empty:
             others['__sort_date'] = others.apply(lambda r: get_date_val(r, start_date_col), axis=1)
             best_other = others.sort_values(by=['__has_license', '__sort_date'], ascending=[False, False])
             return best_other.index[0]

        return group_work.index[0]

    # Find the best index for each group
    best_indices = df.groupby(key_col, group_keys=False).apply(pick_best_idx)
    
    # Safely extract rows using exact indices to perfectly preserve all columns
    deduped = df.loc[best_indices].copy()
    
    return deduped.reset_index(drop=True)

# ---------- Core compare ----------
def compute_audit_dataframes(uzio_file, adp_file):
    """Pure-compute split of run_comparison: returns the DataFrames the Excel writer
    would have written, without the BytesIO/openpyxl step. The standalone tool calls
    this and then writes Excel; the ADP Consolidated Audit calls this and embeds the
    DataFrames into the chief workbook. Logic must stay identical to run_comparison."""
    # 1. Read UZIO Raw
    uzio = read_uzio_raw_file(uzio_file)
    if uzio is None:
        raise ValueError("Failed to read Uzio file.")

    # 2. Read ADP Raw
    try:
        if adp_file.name.lower().endswith('.csv'):
             try:
                 adp = pd.read_csv(adp_file, dtype=str)
             except UnicodeDecodeError:
                 adp_file.seek(0)
                 adp = pd.read_csv(adp_file, dtype=str, encoding='latin1')
        else:
             adp = pd.read_excel(adp_file, dtype=str)
    except Exception as e:
        raise ValueError(f"Failed to read ADP file: {e}")

    # Normalize ADP columns & Ensure Uniqueness
    adp = ensure_unique_columns(adp)
    adp.columns = [norm_colname(c) for c in adp.columns]

    # 3. Apply Mapping Strategy
    # UZIO columns are already standard (e.g. 'First Name')
    # ADP columns need to be looked up via ADP_FIELD_MAP
    
    # Verify Keys exist
    UZIO_KEY = 'Employee ID'
    if UZIO_KEY not in uzio.columns:
         raise ValueError(f"Required column '{UZIO_KEY}' not found in Uzio file.")
    
    ADP_KEY = norm_colname(ADP_FIELD_MAP.get('Employee ID', 'Associate ID'))
    if ADP_KEY not in adp.columns:
         raise ValueError(f"Required column '{ADP_KEY}' not found in ADP file.")

    # Normalize Keys
    uzio[UZIO_KEY] = norm_emp_key_series(uzio[UZIO_KEY])
    adp[ADP_KEY] = norm_emp_key_series(adp[ADP_KEY])


    # Apply the new deduplication
    adp = deduplicate_adp(adp, ADP_KEY)
    
    # Old simple drop (keep unique) - technically redundant but safe as backup for Uzio
    uzio = uzio.drop_duplicates(subset=[UZIO_KEY], keep="first").copy()
    # adp = adp.drop_duplicates(subset=[ADP_KEY], keep="first").copy() # Replaced by above

    # 0. Normalize IDs in the worksets
    uzio[UZIO_KEY] = uzio[UZIO_KEY].apply(norm_id)
    adp[ADP_KEY] = adp[ADP_KEY].apply(norm_id)

    uzio_keys = set(uzio[UZIO_KEY].replace("", pd.NA).dropna())
    adp_keys = set(adp[ADP_KEY].replace("", pd.NA).dropna())
    all_keys = sorted(uzio_keys.union(adp_keys))

    uzio_idx = uzio.set_index(UZIO_KEY, drop=False)
    adp_idx = adp.set_index(ADP_KEY, drop=False)
    # Ensure indices are properly handled as normalized strings
    uzio_idx.index = uzio_idx.index.map(str)
    adp_idx.index = adp_idx.index.map(str)

    uz_to_adp = {k: norm_colname(v) for k, v in ADP_FIELD_MAP.items()}
    mapped_fields = [f for f in ADP_FIELD_MAP.keys() if f != UZIO_KEY] # Internal Standard Keys

    # Identify fields missing in ADP
    
    # ---------------- Salaried Driver Exceptions ----------------
    salaried_drivers_audit = []
    adp_pay_type_col = norm_colname(ADP_FIELD_MAP.get('Pay Type', ''))
    adp_job_title_col = norm_colname(ADP_FIELD_MAP.get('Job Title', ''))
    adp_flsa_col = norm_colname(ADP_FIELD_MAP.get('FLSA Classification', ''))
    adp_emp_status_col = norm_colname(ADP_FIELD_MAP.get('Employment Status', ''))
    uzio_emp_status_col_str = 'Employment Status'
    uzio_pay_type_col_str = 'Pay Type'
    uzio_flsa_col_str = 'FLSA Classification'

    if adp_pay_type_col in adp_idx.columns and adp_job_title_col in adp_idx.columns:
        for emp_id, row in adp_idx.iterrows():
            pay_val = str(row[adp_pay_type_col]).strip().lower()
            if "salary" in pay_val or "salaried" in pay_val:
                jt_raw = row[adp_job_title_col]
                jt_val = str(jt_raw).strip().lower() if pd.notna(jt_raw) else ""
                if jt_val and jt_val != "nan":
                    if is_hourly_only_job_title(jt_val):
                        # --- ADP values ---
                        adp_pay_type_val = str(row[adp_pay_type_col]).strip()
                        adp_emp_status_val = ""
                        if adp_emp_status_col and adp_emp_status_col in adp_idx.columns:
                            adp_emp_status_val = str(norm_blank(row.get(adp_emp_status_col, "")) or "").strip()
                        adp_flsa_val = ""
                        if adp_flsa_col and adp_flsa_col in adp_idx.columns:
                            adp_flsa_val = str(norm_blank(row.get(adp_flsa_col, "")) or "").strip()

                        # --- Uzio values ---
                        uz_exists = emp_id in uzio_idx.index
                        uz_pay_type_val = ""
                        uz_emp_status_val = ""
                        uz_flsa_val = ""
                        if uz_exists:
                            if uzio_pay_type_col_str in uzio_idx.columns:
                                uz_pay_type_val = str(norm_blank(uzio_idx.at[emp_id, uzio_pay_type_col_str]) or "").strip()
                            if uzio_emp_status_col_str in uzio_idx.columns:
                                uz_emp_status_val = str(norm_blank(uzio_idx.at[emp_id, uzio_emp_status_col_str]) or "").strip()
                            if uzio_flsa_col_str in uzio_idx.columns:
                                uz_flsa_val = str(norm_blank(uzio_idx.at[emp_id, uzio_flsa_col_str]) or "").strip()

                        # --- Build smart Comment ---
                        comment_parts = [
                            f"ADP lists this employee as '{str(jt_raw).strip()}' with Pay Type '{adp_pay_type_val}'.",
                            "Uzio requires this job title to be Hourly/Non-Exempt — a Salaried assignment will cause a conflict.",
                        ]
                        if uz_emp_status_val:
                            comment_parts.append(f"Uzio status: {uz_emp_status_val}.")
                        if adp_emp_status_val:
                            comment_parts.append(f"ADP status: {adp_emp_status_val}.")
                        if uz_flsa_val:
                            comment_parts.append(f"Uzio FLSA: {uz_flsa_val}.")
                        if adp_flsa_val:
                            comment_parts.append(f"ADP FLSA: {adp_flsa_val}.")
                        if not uz_exists:
                            comment_parts.append("Employee NOT found in Uzio — will need to be added as Hourly.")
                        comment = " ".join(comment_parts)

                        salaried_drivers_audit.append({
                            'Employee ID': emp_id,
                            'Job Title (ADP)': str(row[adp_job_title_col]).strip(),
                            'Pay Type (ADP)': adp_pay_type_val,
                            'Pay Type (Uzio)': uz_pay_type_val if uz_pay_type_val else "Not in Uzio",
                            'Employment Status (ADP)': adp_emp_status_val if adp_emp_status_val else "Blank",
                            'Employment Status (Uzio)': uz_emp_status_val if uz_emp_status_val else "Not in Uzio" if not uz_exists else "Blank",
                            'FLSA Classification (ADP)': adp_flsa_val if adp_flsa_val else "Blank",
                            'FLSA Classification (Uzio)': uz_flsa_val if uz_flsa_val else "Blank" if uz_exists else "Not in Uzio",
                            'Comment': comment
                        })
    df_salaried_drivers = pd.DataFrame(salaried_drivers_audit)
    # We check if the mapped ADP column exists in adp df
    mapping_missing_adp_col = [] # List of fields

    # Employment Status column (UZIO) - It is 'Employment Status' standard
    uzio_employment_status_col = 'Employment Status'
    
    def get_uzio_employment_status(emp_id: str) -> str:
        if uzio_employment_status_col not in uzio_idx.columns:
             return ""
        if emp_id in uzio_idx.index:
            v = uzio_idx.at[emp_id, uzio_employment_status_col]
            return "" if norm_blank(v) == "" else str(v)
        return ""

    # ADP Employment Status lookup (always show ADP status regardless of Uzio presence)
    def get_adp_employment_status(emp_id: str) -> str:
        if adp_emp_status_col and adp_emp_status_col in adp_idx.columns:
            if emp_id in adp_idx.index:
                v = adp_idx.at[emp_id, adp_emp_status_col]
                return "" if norm_blank(v) == "" else str(v)
        return ""

    # Pay Type mapping (prefer ADP)
    UZIO_PAYTYPE_COL = 'Pay Type'
    ADP_PAYTYPE_COL = norm_colname(ADP_FIELD_MAP.get('Pay Type', ''))

    def get_employee_pay_type(emp_id: str, adp_exists: bool, uz_exists: bool) -> str:
        if ADP_PAYTYPE_COL and adp_exists and (ADP_PAYTYPE_COL in adp_idx.columns):
            v = adp_idx.at[emp_id, ADP_PAYTYPE_COL]
            if norm_blank(v) != "":
                return str(v)
        if UZIO_PAYTYPE_COL and uz_exists and (UZIO_PAYTYPE_COL in uzio_idx.columns):
            v = uzio_idx.at[emp_id, UZIO_PAYTYPE_COL]
            if norm_blank(v) != "":
                return str(v)
        return ""

    # ---------- FLSA Classification column (Uzio) ----------
    # Standard name is 'FLSA Classification'
    uzio_flsa_col = 'FLSA Classification'

    # Also locate employee name columns in Uzio for context in FLSA report
    # Locate ADP columns for name to use as context
    adp_fname_col = None
    for c in adp.columns:
        cl = norm_colname(c).casefold()
        if cl in {"legal first name", "first name", "firstname"}:
            adp_fname_col = c
            break
    if adp_fname_col is None:
        for c in adp.columns:
            cl = norm_colname(c).casefold()
            if "first" in cl and "name" in cl:
                adp_fname_col = c
                break

    adp_lname_col = None
    for c in adp.columns:
        cl = norm_colname(c).casefold()
        if cl in {"legal last name", "last name", "lastname"}:
            adp_lname_col = c
            break
    if adp_lname_col is None:
        for c in adp.columns:
            cl = norm_colname(c).casefold()
            if "last" in cl and "name" in cl:
                adp_lname_col = c
                break

    uzio_ssn_col = 'SSN'
    adp_ssn_col = next((c for c in adp.columns if "Tax ID (SSN)" in c or "SSN" in c), None)

    # 1. Resolve Identity Match Map (UZIO_ID -> ADP_ID)
    uz_to_adp_id_map = get_identity_match_map(
        uzio, adp, 
        uzio_id_col=UZIO_KEY, 
        vendor_id_col=ADP_KEY,
        uzio_ssn_col=uzio_ssn_col,
        vendor_ssn_col=adp_ssn_col
    )

    # 1.1 Data Quality Check: Duplicate SSNs
    uz_dupe_ssns = detect_duplicate_ssns(uzio, UZIO_KEY, uzio_ssn_col)
    adp_dupe_ssns = detect_duplicate_ssns(adp, ADP_KEY, adp_ssn_col)
    
    dupe_ssn_rows = []
    for ssn, ids in uz_dupe_ssns.items():
        dupe_ssn_rows.append({
            "Source": "Uzio",
            "SSN": ssn,
            "Employee IDs": ", ".join(ids),
            "Issue": "Duplicate SSN found for multiple IDs in Uzio"
        })
    for ssn, ids in adp_dupe_ssns.items():
        dupe_ssn_rows.append({
            "Source": "ADP",
            "SSN": ssn,
            "Employee IDs": ", ".join(ids),
            "Issue": "Duplicate SSN found for multiple IDs in ADP"
        })
    df_dupe_ssns = pd.DataFrame(dupe_ssn_rows)

    # 2. Collect all Uzio primary keys and track ADP processed keys
    adp_keys_processed = set()
    rows = []
    uzio_fname_col = 'First Name'
    uzio_lname_col = 'Last Name'

    # 3. Main Loop: Iterate through all Uzio employees
    for uz_id in sorted(uzio_keys):
        # find the associated adp_id (might be the same OR different via identity match)
        adp_id = uz_to_adp_id_map.get(uz_id, uz_id)
        
        uz_exists = True # because we are in uzio_keys
        adp_exists = adp_id in adp_idx.index
        
        if adp_exists:
            adp_keys_processed.add(adp_id)

        uz_emp_status = get_uzio_employment_status(uz_id)
        adp_id_for_status = adp_id if adp_id else uz_id # fallback for missing case
        adp_emp_status_val = get_adp_employment_status(adp_id_for_status) if adp_exists else ""
        emp_paytype = get_employee_pay_type(uz_id if uz_exists else adp_id, adp_exists=adp_exists, uz_exists=uz_exists)
        emp_pay_bucket = paytype_bucket(normalize_paytype_text(emp_paytype))

        # --- Determine Employee Name ---
        fname = ""
        lname = ""
        if uz_exists:
            if uzio_fname_col in uzio_idx.columns:
                fname = str(norm_blank(uzio_idx.at[uz_id, uzio_fname_col]) or "")
            if uzio_lname_col in uzio_idx.columns:
                lname = str(norm_blank(uzio_idx.at[uz_id, uzio_lname_col]) or "")
        
        if (fname == "" and lname == "") and adp_exists:
            if adp_fname_col and adp_fname_col in adp_idx.columns:
                fname = str(norm_blank(adp_idx.at[adp_id, adp_fname_col]) or "")
            if adp_lname_col and adp_lname_col in adp_idx.columns:
                lname = str(norm_blank(adp_idx.at[adp_id, adp_lname_col]) or "")

        emp_name = f"{fname} {lname}".strip()
        
        # Check for ID Mismatch case (Same identity, different IDs)
        if adp_exists and uz_id != adp_id:
            rows.append({
                "Employee ID": uz_id,
                "Employee Name": emp_name,
                "Employment Status": uz_emp_status,
                "Employment Status (ADP)": adp_emp_status_val,
                "Pay Type": emp_paytype,
                "Field": "Employee ID Correlation",
                "UZIO_Value": uz_id,
                "ADP_Value": adp_id,
                "ADP_SourceOfTruth_Status": "Data Mismatch (Identity Match via SSN)"
            })

        for field in mapped_fields:
            adp_col = uz_to_adp.get(field, "")
            
            uz_col_missing = (field not in uzio.columns)
            adp_col_missing = (adp_col not in adp.columns)

            uz_val_raw = safe_val(uzio_idx, uz_id, field) if (uz_exists and not uz_col_missing) else ""
            uz_val = cleanse_uzio_value_for_field(field, uz_val_raw)

            adp_val = safe_val(adp_idx, adp_id, adp_col) if (adp_exists and not adp_col_missing) else ""

            # Check for ID Mismatch case (Same identity, different IDs)
            is_id_mismatch = (adp_exists and uz_id != adp_id)

            if not adp_exists:
                status = "Employee ID Not Found in ADP"
            elif adp_col_missing:
                status = "Column Missing in ADP Sheet"
            elif uz_col_missing:
                status = "Column Missing in Uzio Sheet"
            else:
                if is_pay_type_field(field):
                    uz_pt = normalize_paytype_for_compare(uz_val)
                    adp_pt = normalize_paytype_for_compare(adp_val)

                    if (uz_pt == adp_pt) or (uz_pt == "" and adp_pt == ""):
                        status = "Data Match"
                    elif uz_pt == "" and adp_pt != "":
                        status = "Value missing in Uzio (ADP has value)"
                    elif uz_pt != "" and adp_pt == "":
                        status = "Value missing in ADP (Uzio has value)"
                    else:
                        status = "Data Mismatch"
                else:
                    uz_n = norm_value(uz_val, field)
                    adp_n = norm_value(adp_val, field)

                    if is_employment_status_field(field) and adp_n != "":
                        adp_is_term_or_ret = status_contains_any(adp_n, ["terminated", "retired"])
                        
                        # Special Case: UZIO Active == ADP Leave -> Match
                        is_active_leave = (uzio_is_active(uz_n) and "leave" in adp_n)
                        
                        # Special Case: UZIO Terminated == ADP Deceased -> Match
                        is_term_deceased = (uzio_is_terminated(uz_n) and "deceased" in adp_n)

                        if is_active_leave or is_term_deceased:
                            status = "Data Match"
                        elif (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                             status = "Data Match"
                        elif uzio_is_terminated(uz_n) and adp_is_term_or_ret:
                             # Both terminated/retired but strings diff -> Match
                             status = "Data Match"
                        else:
                            # MISMATCH / MISSING LOGIC per User Request
                            # 1. Active in Uzio
                            if uzio_is_active(uz_n):
                                status = "Active in Uzio"
                            # 2. Terminated in Uzio
                            elif uzio_is_terminated(uz_n):
                                status = "Terminated in Uzio"
                            # 3. Active in ADP (Uzio Blank)
                            elif uz_n == "" and not adp_is_term_or_ret:
                                status = "Active in ADP"
                            # 4. Terminated in ADP (Uzio Blank)
                            elif uz_n == "" and adp_is_term_or_ret:
                                status = "Terminated in ADP"
                            # Fallback for other cases
                            elif uz_n == "" and adp_n != "":
                                status = f"Value missing in Uzio (ADP: {adp_val})"  # Generic fallback
                            elif uz_n != "" and adp_n == "":
                                status = "Value missing in ADP (Uzio has value)"
                            else:
                                status = "Data Mismatch"

                    elif is_termination_reason_field(field):
                        uz_reason = normalize_reason_text(uz_val)
                        adp_reason = normalize_reason_text(adp_val)

                        # Logic: If UZIO says "Other" but ADP gives a reason we know is okay, Match it.
                        is_other_match = (uz_reason == "other" and adp_reason in ALLOWED_TERM_REASONS)
                        
                        # Keyword-based Matching: 
                        # If ADP reason contains "voluntary", match Uzio "Voluntary Termination of Employment"
                        is_voluntary_match = (
                            uz_reason == "voluntary termination of employment" and 
                            "voluntary" in adp_reason
                        )
                        
                        # If ADP reason contains "involuntary", match Uzio "Involuntary Termination of Employment"
                        is_involuntary_match = (
                            uz_reason == "involuntary termination of employment" and 
                            "involuntary" in adp_reason
                        )
                        
                        # Special Case (from before): "Involuntary Termination" == "Layoff" (Layoff doesn't always say 'involuntary')
                        is_layoff_match = (
                            uz_reason == "involuntary termination of employment" and 
                            adp_reason == "layoff"
                        )

                        if is_other_match or is_voluntary_match or is_involuntary_match or is_layoff_match:
                            status = "Data Match"
                        else:
                            if (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                                status = "Data Match"
                            elif uz_n == "" and adp_n != "":
                                status = "Value missing in Uzio (ADP has value)"
                            elif uz_n != "" and adp_n == "":
                                status = "Value missing in ADP (Uzio has value)"
                            else:
                                status = "Data Mismatch"
                    elif is_employment_type_field(field):
                        uz_et = normalize_employment_type(uz_val)
                        adp_et = normalize_employment_type(adp_val)
                        if (uz_et == adp_et) or (uz_et == "" and adp_et == ""):
                            status = "Data Match"
                        elif uz_et == "" and adp_et != "":
                            status = "Value missing in Uzio (ADP has value)"
                        elif uz_et != "" and adp_et == "":
                            status = "Value missing in ADP (Uzio has value)"
                        else:
                            status = "Data Mismatch"
                    else:
                        if (uz_n == adp_n) or (uz_n == "" and adp_n == ""):
                            status = "Data Match"
                        elif uz_n == "" and adp_n != "":
                            status = "Value missing in Uzio (ADP has value)"
                        elif uz_n != "" and adp_n == "":
                            status = "Value missing in ADP (Uzio has value)"
                        else:
                            status = "Data Mismatch"

                        if status in ["Value missing in Uzio (ADP has value)", "Data Mismatch"]:
                            if emp_pay_bucket == "hourly" and is_annual_salary_field(field):
                                status = "Data Match"
                            elif emp_pay_bucket == "salaried" and is_hourly_rate_field(field):
                                status = "Data Match"

            rows.append({
                "Employee ID": uz_id,
                "Employee Name": emp_name,
                "Employment Status": uz_emp_status,
                "Employment Status (ADP)": adp_emp_status_val,
                "Pay Type": emp_paytype,
                "Field": field,
                "UZIO_Value": uz_val_raw,
                "ADP_Value": adp_val,
                "ADP_SourceOfTruth_Status": status
            })

    # 4. Final Loop: Remaining ADP employees not in Uzio (even via identity)
    remaining_adp_ids = set(adp_keys) - adp_keys_processed
    for adp_id in sorted(remaining_adp_ids):
        adp_emp_status_val = get_adp_employment_status(adp_id)
        emp_paytype = get_employee_pay_type(adp_id, adp_exists=True, uz_exists=False)
        emp_pay_bucket = paytype_bucket(normalize_paytype_text(emp_paytype))

        fname = str(norm_blank(adp_idx.at[adp_id, adp_fname_col]) or "") if adp_fname_col else ""
        lname = str(norm_blank(adp_idx.at[adp_id, adp_lname_col]) or "") if adp_lname_col else ""
        emp_name = f"{fname} {lname}".strip()

        for field in mapped_fields:
            adp_col = uz_to_adp.get(field, "")
            adp_val = safe_val(adp_idx, adp_id, adp_col) if adp_col in adp.columns else ""

            rows.append({
                "Employee ID": adp_id,
                "Employee Name": emp_name,
                "Employment Status": "",
                "Employment Status (ADP)": adp_emp_status_val,
                "Pay Type": emp_paytype,
                "Field": field,
                "UZIO_Value": "",
                "ADP_Value": adp_val,
                "ADP_SourceOfTruth_Status": "Employee ID Not Found in Uzio"
            })

    comparison_detail = pd.DataFrame(rows)[[
        "Employee ID", "Employee Name", "Employment Status", "Employment Status (ADP)", "Pay Type",
        "Field", "UZIO_Value", "ADP_Value", "ADP_SourceOfTruth_Status"
    ]]

    mismatches_only = comparison_detail[comparison_detail["ADP_SourceOfTruth_Status"] != "Data Match"].copy()

    # ---------- FLSA Compliance Issues (4th sheet) ----------
    flsa_rows = []
    if uzio_flsa_col is not None:
        for emp_id in uzio_keys:
            if emp_id not in uzio_idx.index:
                continue
            # Get Pay Type from Uzio
            uz_pay_raw = ""
            if UZIO_PAYTYPE_COL and UZIO_PAYTYPE_COL in uzio_idx.columns:
                uz_pay_raw = uzio_idx.at[emp_id, UZIO_PAYTYPE_COL]
            pay_type_val = normalize_paytype_text(uz_pay_raw)
            pay_bucket = paytype_bucket(pay_type_val)

            # Get FLSA Classification from Uzio
            flsa_raw = uzio_idx.at[emp_id, uzio_flsa_col] if uzio_flsa_col in uzio_idx.columns else ""
            flsa_norm = normalize_paytype_text(flsa_raw)  # reuse: lowercases & strips

            # Get ADP values for context
            adp_id = uz_to_adp_id_map.get(emp_id, emp_id)
            adp_exists = adp_id in adp_idx.index
            
            adp_pay_type = ""
            adp_flsa = ""
            adp_job = ""
            adp_dept = ""
            uzio_job = ""
            
            if adp_exists:
                if adp_pay_type_col in adp_idx.columns:
                    adp_pay_type = str(norm_blank(adp_idx.at[adp_id, adp_pay_type_col]) or "").strip()
                if adp_flsa_col in adp_idx.columns:
                    adp_flsa = str(norm_blank(adp_idx.at[adp_id, adp_flsa_col]) or "").strip()
                if adp_job_title_col in adp_idx.columns:
                    adp_job = str(norm_blank(adp_idx.at[adp_id, adp_job_title_col]) or "").strip()
                if 'Department Description' in adp_idx.columns:
                    adp_dept = str(norm_blank(adp_idx.at[adp_id, 'Department Description']) or "").strip()
            
            if 'Job Title' in uzio_idx.columns:
                uzio_job = str(norm_blank(uzio_idx.at[emp_id, 'Job Title']) or "").strip()

            # Detect Issues
            all_issues = []
            
            # 1. Internal Uzio Inconsistency
            if pay_bucket == "hourly" and "exempt" in flsa_norm and "non" not in flsa_norm:
                all_issues.append("Hourly employee classified as Exempt (Uzio Internal)")
            elif pay_bucket == "salaried" and ("non-exempt" in flsa_norm or "non exempt" in flsa_norm or "nonexempt" in flsa_norm):
                all_issues.append("Salaried employee classified as Non-Exempt (Uzio Internal)")

            # 2. Cross-system Mismatches
            if adp_exists:
                # Pay Type Mismatch
                uz_pt_canon = normalize_paytype_for_compare(uz_pay_raw)
                adp_pt_canon = normalize_paytype_for_compare(adp_pay_type)
                if uz_pt_canon != adp_pt_canon and adp_pt_canon != "":
                    all_issues.append(f"Pay Type Mismatch (Uzio: {uz_pay_raw} vs ADP: {adp_pay_type})")
                
                # FLSA Mismatch
                uz_flsa_canon = normalize_paytype_text(flsa_raw)
                adp_flsa_canon = normalize_paytype_text(adp_flsa)
                if uz_flsa_canon != adp_flsa_canon and adp_flsa_canon != "":
                    all_issues.append(f"FLSA Mismatch (Uzio: {flsa_raw} vs ADP: {adp_flsa})")

            if all_issues:
                # Get employee name for context
                fname = ""
                lname = ""
                if uzio_fname_col and uzio_fname_col in uzio_idx.columns:
                    fname = str(norm_blank(uzio_idx.at[emp_id, uzio_fname_col]) or "")
                if uzio_lname_col and uzio_lname_col in uzio_idx.columns:
                    lname = str(norm_blank(uzio_idx.at[emp_id, uzio_lname_col]) or "")
                emp_name = f"{fname} {lname}".strip()

                flsa_rows.append({
                    "Employee ID": emp_id,
                    "Employee Name": emp_name,
                    "Pay Type (Uzio)": str(norm_blank(uz_pay_raw) or ""),
                    "Pay Type (ADP)": adp_pay_type,
                    "FLSA Classification (Uzio)": str(norm_blank(flsa_raw) or ""),
                    "FLSA Classification (ADP)": adp_flsa,
                    "Job Title (Uzio)": uzio_job,
                    "Job Title (ADP)": adp_job,
                    "Department (ADP)": adp_dept,
                    "Issue": "; ".join(all_issues),
                })

    flsa_issues = pd.DataFrame(flsa_rows, columns=[
        "Employee ID", "Employee Name", 
        "Pay Type (Uzio)", "Pay Type (ADP)",
        "FLSA Classification (Uzio)", "FLSA Classification (ADP)",
        "Job Title (Uzio)", "Job Title (ADP)",
        "Department (ADP)",
        "Issue"
    ])

    # ---------- Data Quality Issues (00/00/0000 dates) ----------
    dq_rows = []
    
    for emp_id in adp_idx.index:
        # Check all columns for this row
        for col in adp.columns:
            val = adp_idx.at[emp_id, col]
            if pd.notna(val) and '00/00/0000' in str(val):
                fname = str(norm_blank(adp_idx.at[emp_id, adp_fname_col]) or "") if adp_fname_col else ""
                lname = str(norm_blank(adp_idx.at[emp_id, adp_lname_col]) or "") if adp_lname_col else ""
                emp_name = f"{fname} {lname}".strip()
                
                dq_rows.append({
                    "Employee ID": str(emp_id).strip(),
                    "Employee Name": emp_name,
                    "Column": col,
                    "Invalid Value Found": str(val)
                })
                
    dq_issues = pd.DataFrame(dq_rows, columns=[
        "Employee ID", "Employee Name", "Column", "Invalid Value Found"
    ])

    # ---------- Active Employees Missing in Uzio (5th sheet) ----------
    # Find employees in ADP but NOT in Uzio who are Active
    adp_only_keys = adp_keys - uzio_keys

    # Locate ADP columns for status, name, hire date
    adp_status_col = None
    for c in adp.columns:
        cl = norm_colname(c).casefold()
        if cl in {"position status", "employment status", "employee status"}:
            adp_status_col = c
            break
    if adp_status_col is None:
        for c in adp.columns:
            cl = norm_colname(c).casefold()
            if "status" in cl and ("position" in cl or "employment" in cl):
                adp_status_col = c
                break

    # Name columns already located above for Data Quality check.
    # We can reuse adp_fname_col and adp_lname_col

    adp_hire_col = None
    for c in adp.columns:
        cl = norm_colname(c).casefold()
        if "hire" in cl and "date" in cl:
            adp_hire_col = c
            break

    active_missing_rows = []
    for emp_id in sorted(adp_only_keys):
        if emp_id not in adp_idx.index:
            continue
        # Check employment status in ADP
        status_val = ""
        if adp_status_col and adp_status_col in adp_idx.columns:
            status_val = str(norm_blank(adp_idx.at[emp_id, adp_status_col]) or "")
        status_lower = status_val.strip().lower()

        # Only include Active / Leave employees
        if "active" not in status_lower and "leave" not in status_lower:
            continue

        fname = ""
        lname = ""
        if adp_fname_col and adp_fname_col in adp_idx.columns:
            fname = str(norm_blank(adp_idx.at[emp_id, adp_fname_col]) or "")
        if adp_lname_col and adp_lname_col in adp_idx.columns:
            lname = str(norm_blank(adp_idx.at[emp_id, adp_lname_col]) or "")
        emp_name = f"{fname} {lname}".strip()

        hire_date = ""
        if adp_hire_col and adp_hire_col in adp_idx.columns:
            hire_date = str(norm_blank(adp_idx.at[emp_id, adp_hire_col]) or "")

        active_missing_rows.append({
            "Employee ID": emp_id,
            "Employee Name": emp_name,
            "Employment Status (ADP)": status_val,
            "Date of Hire (ADP)": hire_date,
        })

    active_missing_in_uzio = pd.DataFrame(active_missing_rows, columns=[
        "Employee ID", "Employee Name",
        "Employment Status (ADP)", "Date of Hire (ADP)"
    ])

    terminated_missing_rows = []
    for emp_id in sorted(adp_only_keys):
        if emp_id not in adp_idx.index:
            continue
        # Check employment status in ADP
        status_val = ""
        if adp_status_col and adp_status_col in adp_idx.columns:
            status_val = str(norm_blank(adp_idx.at[emp_id, adp_status_col]) or "")
        status_lower = status_val.strip().lower()

        # Only include Terminated / Inactive / Quit / Resign / Retired employees
        is_term = False
        term_keywords = ["terminated", "retired", "inactive", "quit", "resign", "term"]
        for kw in term_keywords:
            if kw in status_lower:
                is_term = True
                break
        
        if not is_term:
            continue

        fname = ""
        lname = ""
        if adp_fname_col and adp_fname_col in adp_idx.columns:
            fname = str(norm_blank(adp_idx.at[emp_id, adp_fname_col]) or "")
        if adp_lname_col and adp_lname_col in adp_idx.columns:
            lname = str(norm_blank(adp_idx.at[emp_id, adp_lname_col]) or "")
        emp_name = f"{fname} {lname}".strip()

        hire_date = ""
        if adp_hire_col and adp_hire_col in adp_idx.columns:
            hire_date = str(norm_blank(adp_idx.at[emp_id, adp_hire_col]) or "")

        terminated_missing_rows.append({
            "Employee ID": emp_id,
            "Employee Name": emp_name,
            "Employment Status (ADP)": status_val,
            "Date of Hire (ADP)": hire_date,
        })

    terminated_missing_in_uzio = pd.DataFrame(terminated_missing_rows, columns=[
        "Employee ID", "Employee Name",
        "Employment Status (ADP)", "Date of Hire (ADP)"
    ])

    # ---------- Field Summary By Status ----------
    cols_needed = [
        "Data Match",
        "Data Mismatch",
        "Value missing in Uzio (ADP has value)",
        "Value missing in ADP (Uzio has value)",
        "Employee ID Not Found in Uzio",
        "Employee ID Not Found in ADP",
        "Column Missing in ADP Sheet",
        "Column Missing in Uzio Sheet",
    ]

    pivot = comparison_detail.pivot_table(
        index="Field",
        columns="ADP_SourceOfTruth_Status",
        values="Employee ID",
        aggfunc="count",
        fill_value=0
    )

    for c in cols_needed:
        if c not in pivot.columns:
            pivot[c] = 0

    pivot["Total"] = pivot.sum(axis=1)
    pivot["Data Match"] = pivot["Data Match"].astype(int)

    field_summary_by_status = pivot.reset_index()[[
        "Field",
        "Total",
        "Data Match",
        "Data Mismatch",
        "Value missing in Uzio (ADP has value)",
        "Value missing in ADP (Uzio has value)",
        "Employee ID Not Found in Uzio",
        "Employee ID Not Found in ADP",
        "Column Missing in ADP Sheet",
        "Column Missing in Uzio Sheet"
    ]]

    # ---------- High Hourly Rate Anomalies (> $100/hr for hourly-only roles) ----------
    high_rate_rows = []
    adp_hourly_rate_col = norm_colname(ADP_FIELD_MAP.get('Hourly Pay Rate', ''))
    adp_job_col_hr = norm_colname(ADP_FIELD_MAP.get('Job Title', ''))
    adp_fname_hr = norm_colname(ADP_FIELD_MAP.get('First Name', ''))
    adp_lname_hr = norm_colname(ADP_FIELD_MAP.get('Last Name', ''))
    HOURLY_RATE_THRESHOLD = 100.0

    if adp_hourly_rate_col in adp_idx.columns and adp_job_col_hr in adp_idx.columns:
        for emp_id, row in adp_idx.iterrows():
            jt_raw = row.get(adp_job_col_hr, '')
            jt_str = str(jt_raw).strip().lower() if pd.notna(jt_raw) else ""
            if not jt_str or jt_str == "nan":
                continue
            if is_hourly_only_job_title(jt_str):
                rate_raw = row.get(adp_hourly_rate_col, '')
                try:
                    rate = float(str(rate_raw).replace('$', '').replace(',', '').strip())
                except (ValueError, TypeError):
                    continue
                if rate > HOURLY_RATE_THRESHOLD:
                    fname = str(norm_blank(row.get(adp_fname_hr, '')) or '').strip()
                    lname = str(norm_blank(row.get(adp_lname_hr, '')) or '').strip()
                    emp_name = f"{fname} {lname}".strip()
                    uz_rate = ""
                    if emp_id in uzio_idx.index and 'Hourly Pay Rate' in uzio_idx.columns:
                        uz_rate = str(norm_blank(uzio_idx.at[emp_id, 'Hourly Pay Rate']) or '').strip()
                    high_rate_rows.append({
                        'Employee ID': emp_id,
                        'Employee Name': emp_name,
                        'Job Title (ADP)': str(jt_raw).strip(),
                        'Hourly Pay Rate (ADP)': f"${rate:.2f}",
                        'Hourly Pay Rate (Uzio)': f"${float(uz_rate):.2f}" if uz_rate else "Not in Uzio",
                        'Comment': f"Hourly rate ${rate:.2f}/hr exceeds the ${HOURLY_RATE_THRESHOLD:.0f}/hr threshold for a '{str(jt_raw).strip()}' role. Please verify this is not a data entry error."
                    })
    df_high_rate = pd.DataFrame(high_rate_rows)

    # ---------- Hourly = 0 Hours validation (Check Uzio only) ----------
    hourly_zero_hours_rows = []
    if 'Pay Type' in uzio_idx.columns and 'Working Hours' in uzio_idx.columns:
        for emp_id, row in uzio_idx.iterrows():
            uz_pay_raw = str(norm_blank(row.get('Pay Type', '')) or "")
            emp_pay_bucket = paytype_bucket(normalize_paytype_text(uz_pay_raw))
            if emp_pay_bucket == "hourly":
                wh_raw = row.get('Working Hours', '')
                try:
                    wh_val = float(str(wh_raw).replace(",", "").strip()) if str(wh_raw).strip() else 0.0
                except Exception:
                    wh_val = 0.0
                
                if wh_val > 0:
                    fname = str(norm_blank(row.get('First Name', '')) or "")
                    lname = str(norm_blank(row.get('Last Name', '')) or "")
                    emp_name = f"{fname} {lname}".strip()
                    hourly_zero_hours_rows.append({
                        "Employee ID": str(emp_id),
                        "Employee Name": emp_name,
                        "Pay Type (Uzio)": uz_pay_raw,
                        "Working Hours (Uzio)": str(wh_raw),
                        "Issue": f"Hourly employee has {wh_raw} working hours. Must be 0."
                    })
    df_hourly_zero_hours = pd.DataFrame(hourly_zero_hours_rows)

    # ---------- Summary metrics ----------
    summary = pd.DataFrame({
        "Metric": [
            "Employees in UZIO sheet",
            "Employees in ADP sheet",
            "Employees present in both",
            "Employees missing in ADP (UZIO only)",
            "Employees missing in UZIO (ADP only)",
            "Mapped fields total (from mapping sheet)",
            "Mapped fields with ADP column missing",
            "Total comparison rows (employees x mapped fields)",
            "Total NOT OK rows",
            "FLSA Compliance Issues",
            "Active in ADP but Missing in Uzio",
            "Terminated in ADP but Missing in Uzio",
            "Data Quality Issues (00/00/0000)",
            "Duplicate SSN Warnings",
            "Salaried Hourly-Only Exceptions",
            "High Hourly Rate Anomalies (>$100/hr)",
            "Hourly Zero Hours Exceptions"
        ],
        "Value": [
            len(uzio_keys),
            len(adp_keys),
            len(uzio_keys.intersection(adp_keys)),
            len(uzio_keys - adp_keys),
            len(adp_keys - uzio_keys),
            len(mapped_fields),
            len(mapping_missing_adp_col),
            comparison_detail.shape[0],
            mismatches_only.shape[0],
            len(flsa_rows),
            len(active_missing_rows),
            len(terminated_missing_rows),
            len(dq_rows),
            len(dupe_ssn_rows),
            len(df_salaried_drivers),
            len(df_high_rate),
            len(df_hourly_zero_hours)
        ]
    })

    return {
        "Summary": summary,
        "Field_Summary_By_Status": field_summary_by_status,
        "Comparison_Detail_AllFields": comparison_detail,
        "FLSA_Compliance_Issues": flsa_issues,
        "Data_Quality_Issues": dq_issues,
        "Active_Missing_In_Uzio": active_missing_in_uzio,
        "Terminated_Missing_In_Uzio": terminated_missing_in_uzio,
        "Duplicate_SSN_Check": df_dupe_ssns,
        "Salaried_Driver_Exceptions": df_salaried_drivers,
        "High_Hourly_Rate_Anomalies": df_high_rate,
        "Hourly_Zero_Hours_Exceptions": df_hourly_zero_hours,
    }


def run_comparison(uzio_file, adp_file) -> bytes:
    dfs = compute_audit_dataframes(uzio_file, adp_file)
    summary = dfs["Summary"]
    field_summary_by_status = dfs["Field_Summary_By_Status"]
    comparison_detail = dfs["Comparison_Detail_AllFields"]
    flsa_issues = dfs["FLSA_Compliance_Issues"]
    dq_issues = dfs["Data_Quality_Issues"]
    active_missing_in_uzio = dfs["Active_Missing_In_Uzio"]
    terminated_missing_in_uzio = dfs["Terminated_Missing_In_Uzio"]
    df_dupe_ssns = dfs["Duplicate_SSN_Check"]
    df_salaried_drivers = dfs["Salaried_Driver_Exceptions"]
    df_high_rate = dfs["High_Hourly_Rate_Anomalies"]
    df_hourly_zero_hours = dfs["Hourly_Zero_Hours_Exceptions"]

    # ---------- Export report ----------
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        field_summary_by_status.to_excel(writer, sheet_name="Field_Summary_By_Status", index=False)
        comparison_detail.to_excel(writer, sheet_name="Comparison_Detail_AllFields", index=False)
        flsa_issues.to_excel(writer, sheet_name="FLSA_Compliance_Issues", index=False)
        dq_issues.to_excel(writer, sheet_name="Data_Quality_Issues", index=False)
        active_missing_in_uzio.to_excel(writer, sheet_name="Active_Missing_In_Uzio", index=False)
        terminated_missing_in_uzio.to_excel(writer, sheet_name="Terminated_Missing_In_Uzio", index=False)
        if not df_dupe_ssns.empty:
            df_dupe_ssns.to_excel(writer, sheet_name="Duplicate_SSN_Check", index=False)
        if not df_salaried_drivers.empty:
            df_salaried_drivers.to_excel(writer, sheet_name="Salaried_Driver_Exceptions", index=False)
        if not df_high_rate.empty:
            df_high_rate.to_excel(writer, sheet_name="High_Hourly_Rate_Anomalies", index=False)
        if not df_hourly_zero_hours.empty:
            df_hourly_zero_hours.to_excel(writer, sheet_name="Hourly_Zero_Hours_Exceptions", index=False)

    return out.getvalue()

# ---------- Minimal UI ----------
def render_ui():
    st.title(APP_TITLE)
    st.markdown("""
    **Instructions**:
    1. Upload **Standard Uzio Census Template** (.xlsm).
    2. Upload **ADP Census Template** (.xlsx).
    3. Make sure the ADP Census template does not have licence details and emergency details for that we have seperate tool.
    
    
    **Output Reports**:
    - **Comparison**: Discrepancies between Uzio and ADP.
    - **FLSA_Compliance_Issues**: Invalid Pay Type/FLSA Classification.
    - **Active_Missing_In_Uzio**: Active employees in ADP not found in Uzio.
    - **Data_Quality_Issues**: Identifies dates with '00/00/0000'.
    - **Salaried_Driver_Exceptions**: Employees mapped as salaried drivers, which are incompatible.
    """)

    client_name = st.text_input("Client Name", value="Client", key="adp_census_client")

    uzio_file = st.file_uploader("Upload Uzio Census Export (.xlsm)", type=["xlsm"])
    adp_file = st.file_uploader("Upload ADP Census Export (.csv or .xlsx)", type=["csv", "xlsx"])

    if st.button("Run Audit", type="primary", disabled=(not uzio_file or not adp_file)):
        try:
            with st.spinner("Running audit..."):
                # run_comparison now expects (uzio_file, adp_file) per my audit_utils logic
                # But wait, did I update run_comparison in census_audit_app.py?
                # I need to check if run_comparison signature was updated!
                # I suspect I updated it in Step 594?
                # I need to check run_comparison signature first!
                out_excel = run_comparison(uzio_file, adp_file)
            
            st.success("Report generated.")
            
            timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
            out_filename = f"{client_name}_Uzio_ADP_Census_Audit_Report_{timestamp}.xlsx"
            
            st.download_button(
                label="Download Report (.xlsx)",
                data=out_excel,
                file_name=out_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )
        except Exception as e:
            st.error(f"Failed: {e}")

if __name__ == "__main__":
    st.set_page_config(page_title=APP_TITLE, layout="centered", initial_sidebar_state="collapsed")
    render_ui()
