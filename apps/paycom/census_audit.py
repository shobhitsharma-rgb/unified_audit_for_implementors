# app.py
import io
import re
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
from utils.audit_utils import (
    read_uzio_raw_file,
    HOURLY_ONLY_JOB_TITLES, is_hourly_only_job_title,
    norm_colname, norm_blank, try_parse_date, ensure_unique_columns, safe_val, normalize_space_and_case,
    norm_key_series, as_float_or_none, find_col, get_identity_match_map, norm_ssn_canonical, detect_duplicate_ssns, norm_id
)

# =========================================================
# Paycom vs UZIO – Census Audit Tool
# - User uploads Raw Uzio Export (.xlsm) and Raw Paycom Export (.csv)
# - Hardcoded mappings
# =========================================================

APP_TITLE = "Paycom Uzio Census Audit Tool"

# Hardcoded Mapping: Internal Standard Name -> Paycom Column Name
PAYCOM_FIELD_MAP = {
    'Employee ID': 'Employee_Code',
    'First Name': 'Legal_Firstname',
    'Last Name': 'Legal_Lastname',
    'Middle Initial': 'Legal_Middle_Name',
    'Suffix': 'Legal_Employee_Suffix',
    'Employment Status': 'Employee_Status',
    'Employment Type': 'DOL_Status',
    'Hire Date': 'Most_Recent_Hire_Date',
    'Original Hire Date': 'Hire_Date',
    'Termination Date': 'Termination_Date',
    'Termination Reason': 'Termination_Reason',
    'Pay Type': 'Pay_Type',
    'Annual Salary': 'Annual_Salary',
    'Hourly Pay Rate': 'Rate_1',
    'Working Hours': 'Scheduled_Pay_Period_Hours',
    'Job Title': 'Position',
    'Department': 'Department_Desc',
    'Work Email': 'Work_Email',
    'Personal Email': 'Personal_Email',
    'Phone Number': 'Primary_Phone',
    'SSN': 'SS_Number',
    'DOB': 'Birth_Date_(MM/DD/YYYY)',
    'Gender': 'Gender',
    'Tobacco User': 'Tobacco_User',
    'FLSA Classification': 'Exempt_Status',
    'Address Line 1': 'Primary_Address_Line_1',
    'Address Line 2': 'Primary_Address_Line_2',
    'City': 'Primary_City/Municipality',
    'Zip': 'Primary_Zip/Postal_Code',
    'State': 'Primary_State/Province',
    'Mailing Address Line 1': 'Mailing_Address_Line_1',
    'Mailing Address Line 2': 'Mailing_Address_Line_2',
    'Mailing City': 'Mailing_City/Municipality',
    'Mailing Zip': 'Mailing_Zip/Postal_Code',
    'Mailing State': 'Mailing_State/Province',
    'License Number': 'DriversLicense',
    'License Expiration Date': 'DLExpirationDate',
    'Work Location': 'Work_Location',
    'Reports To ID': 'Supervisor_Primary_Code',
    'Ethnicity': 'EEO1_Ethnicity',
    'SOC Code': 'SOC_Code',
    'EEO Job Category': 'EEO1_Category'
}

# ---------- Helpers ----------
# (redunant helpers removed, using utils.audit_utils)

def normalize_employment_type(x):
    s = normalize_space_and_case(x)
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    if s in {"full time", "fulltime", "ft"}:
        return "full time"
    if s in {"part time", "parttime", "pt"}:
        return "part time"
    if s in {"seasonal", "temporary", "temp"}:
        return "seasonal"
    return s

def normalize_suffix(x):
    s = normalize_space_and_case(x)
    s = re.sub(r"[^a-z0-9]", "", s)  # remove punctuation/spaces
    return s

def normalize_phone(x):
    s = norm_blank(x)
    if s == "":
        return ""
    # If pandas read it as a float, it might look like '9048729456.0'
    s_str = str(s).strip()
    if s_str.endswith(".0"):
        s_str = s_str[:-2]
    # remove all non-digits
    digits = re.sub(r"[^0-9]", "", s_str)
    # if 11 digits and starts with 1, remove leading 1
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits

def first_alpha_char(x):
    s = norm_blank(x)
    if s == "":
        return ""
    txt = str(s).strip()
    m = re.search(r"[A-Za-z]", txt)
    return m.group(0).casefold() if m else ""

def normalize_middle_initial(uzio_val, paycom_val):
    # UZIO has 'M', Paycom has 'MICHELLE' => OK if first letter matches
    u = first_alpha_char(uzio_val)
    p = first_alpha_char(paycom_val)
    return u != "" and p != "" and u == p

def canonical_pay_type(x):
    s = normalize_space_and_case(x)
    if s == "":
        return ""
    if "hour" in s:
        return "hourly"
    if "salar" in s or "salary" in s:
        return "salaried"
    return s

def canonical_employment_status(x):
    # Paycom "On Leave" treated as "Active"
    s = normalize_space_and_case(x)
    if s == "":
        return ""
    if "on leave" in s:
        return "active"
    if s in {"active", "activated"}:
        return "active"
    return s

def termination_reason_equal(uzio_val, paycom_val):
    uz = normalize_space_and_case(uzio_val)
    pc = normalize_space_and_case(paycom_val)

    if uz == "" and pc == "":
        return True

    # UZIO "Other" is acceptable for any Paycom reason
    if uz == "other":
        return True

    # If either has involuntary, both must have involuntary
    if ("involuntary" in uz) or ("involuntary" in pc):
        return ("involuntary" in uz) and ("involuntary" in pc)

    # If either has voluntary, both must have voluntary
    if ("voluntary" in uz) or ("voluntary" in pc):
        return ("voluntary" in uz) and ("voluntary" in pc)

    return uz == pc

def resolve_sheet_name(xls: pd.ExcelFile, candidates):
    existing_norm = {norm_colname(s).casefold(): s for s in xls.sheet_names}
    for c in candidates:
        k = norm_colname(c).casefold()
        if k in existing_norm:
            return existing_norm[k]
    return None

def resolve_paycom_col_label(label: str, paycom_cols_all) -> str:
    if label is None:
        return ""
    raw = str(label).strip()
    raw = raw.replace("’", "'").replace("“", '"').replace("”", '"')
    raw = raw.strip().strip(",")
    if raw == "":
        return ""

    pay_norm = {norm_colname(c).casefold(): c for c in paycom_cols_all}

    direct = norm_colname(raw).casefold()
    if direct in pay_norm:
        return pay_norm[direct]

    parts = re.split(r"\(|\)|\bor\b|/|,|;", raw, flags=re.IGNORECASE)
    parts = [norm_colname(p) for p in parts if norm_colname(p)]

    extra = []
    for p in parts:
        extra.extend([norm_colname(x) for x in re.split(r"\s[-–]\s", p) if norm_colname(x)])
    parts = parts + extra

    for p in parts:
        k = norm_colname(p).casefold()
        if k in pay_norm:
            return pay_norm[k]

    for k_norm, actual in pay_norm.items():
        if k_norm and (k_norm in direct or direct in k_norm):
            return actual

    return ""

def read_mapping_sheet(xls: pd.ExcelFile, sheet_name: str, paycom_cols_all: list) -> pd.DataFrame:
    m = pd.read_excel(xls, sheet_name=sheet_name, dtype=object)
    m.columns = [norm_colname(c) for c in m.columns]

    uz_col_name = None
    pc_col_name = None
    for c in m.columns:
        if norm_colname(c).casefold() in {"uzio coloumn", "uzio column"}:
            uz_col_name = c
        if norm_colname(c).casefold() in {"paycom coloumn", "paycom column"}:
            pc_col_name = c

    if uz_col_name is None or pc_col_name is None:
        raise ValueError(f"'{sheet_name}' must contain columns: 'UZIO Column' and 'Paycom Column'.")

    m[uz_col_name] = m[uz_col_name].map(norm_colname)
    m[pc_col_name] = m[pc_col_name].map(norm_colname)

    m = m.dropna(subset=[uz_col_name, pc_col_name]).copy()
    m = m[(m[uz_col_name] != "") & (m[pc_col_name] != "")]
    m = m.drop_duplicates(subset=[uz_col_name], keep="first").copy()

    m["UZIO_Column"] = m[uz_col_name]
    m["PAYCOM_Label"] = m[pc_col_name]
    m["PAYCOM_Resolved_Column"] = m["PAYCOM_Label"].map(lambda x: resolve_paycom_col_label(x, paycom_cols_all))

    # exclude Employee ID/Employee Code from comparisons (key only)
    m["_uz_norm"] = m["UZIO_Column"].map(lambda x: norm_colname(x).casefold())
    m = m[~m["_uz_norm"].isin({"employee id", "employee", "employee_code", "employee code"})].copy()
    m.drop(columns=["_uz_norm"], inplace=True)

    return m

def should_ignore_field_for_paytype(field_name: str, pay_type_canon: str) -> bool:
    """
    Pay-type based ignore rules (as per your requirement):
      - HOURLY employees: ignore annual salary fields
      - SALARIED employees: ignore hourly pay rate AND working hours per week
    """
    f = norm_colname(field_name).casefold()
    pt = (pay_type_canon or "").casefold()

    if pt == "hourly":
        if "annual salary" in f:
            return True

    if pt == "salaried":
        # ignore Hourly Pay Rate (covers: "Hourly Pay Rate", "Hourly Rate", etc.)
        if ("hourly" in f and "rate" in f):
            return True

        # ignore Working Hours per Week (covers: "Working Hours per Week(Digits)", "Hours per Week", etc.)
        if ("hours per week" in f) or ("working hours" in f):
            return True

    return False

def normalized_compare(field_name: str, uzio_val, paycom_val) -> bool:
    f = norm_colname(field_name).casefold()

    if "termination reason" in f:
        return termination_reason_equal(uzio_val, paycom_val)

    if "employment status" in f:
        return canonical_employment_status(uzio_val) == canonical_employment_status(paycom_val)

    if "pay type" in f:
        return canonical_pay_type(uzio_val) == canonical_pay_type(paycom_val)

    if "employment type" in f:
        return normalize_employment_type(uzio_val) == normalize_employment_type(paycom_val)

    if ("middle" in f) and ("initial" in f):
        if normalize_middle_initial(uzio_val, paycom_val):
            return True
        return first_alpha_char(uzio_val) == first_alpha_char(paycom_val)

    if "suffix" in f:
        return normalize_suffix(uzio_val) == normalize_suffix(paycom_val)

    if "ssn" in f:
        # Normalize SSN: digits only, remove leading zeros
        u = re.sub(r"\D", "", str(uzio_val)).lstrip("0")
        p = re.sub(r"\D", "", str(paycom_val)).lstrip("0")
        return u == p

    if "phone" in f:
        # Normalize Phone: digits only, remove leading zeros
        u = normalize_phone(uzio_val).lstrip("0")
        p = normalize_phone(paycom_val).lstrip("0")
        return u == p

    if "zip" in f:
        # Normalize Zip: digits only (simple), remove leading zeros
        u = re.sub(r"\D", "", str(uzio_val)).lstrip("0")
        p = re.sub(r"\D", "", str(paycom_val)).lstrip("0")
        return u == p

    # Date-ish fields (including DOH)
    if any(k in f for k in ["date", "dob", "birth", "effective", "doh", "hire", "termination"]):
        return try_parse_date(uzio_val) == try_parse_date(paycom_val)

    # Numeric-ish fields
    if any(k in f for k in ["salary", "rate", "hours", "amount", "percent", "percentage", "digits"]):
        fa = as_float_or_none(uzio_val)
        fb = as_float_or_none(paycom_val)
        if fa is not None and fb is not None:
            return abs(fa - fb) <= 1e-9
        return normalize_space_and_case(uzio_val) == normalize_space_and_case(paycom_val)

    if "license" in f:
        # Standardize License Number: remove leading zeros
        u = str(uzio_val).strip().lstrip("0")
        p = str(paycom_val).strip().lstrip("0")
        return u == p

    return normalize_space_and_case(uzio_val) == normalize_space_and_case(paycom_val)

# ---------- Core comparison ----------
def run_comparison(uzio_file, paycom_file) -> bytes:
    # 1. Read UZIO Raw
    uzio = read_uzio_raw_file(uzio_file)
    if uzio is None:
        raise ValueError("Failed to read Uzio file.")

    # 2. Read Paycom Raw
    try:
        # Determine encoding or engine
        if paycom_file.name.lower().endswith('.csv'):
             try:
                 paycom = pd.read_csv(paycom_file, dtype=str)
             except UnicodeDecodeError:
                 paycom_file.seek(0)
                 paycom = pd.read_csv(paycom_file, dtype=str, encoding='latin1')
        else:
             paycom = pd.read_excel(paycom_file, dtype=str)
    except Exception as e:
        raise ValueError(f"Failed to read Paycom file: {e}")

    # Column Normalization & De-duplication
    uzio = ensure_unique_columns(uzio)
    uzio.columns = [norm_colname(c) for c in uzio.columns]
    uzio = uzio.reset_index(drop=True)

    paycom = ensure_unique_columns(paycom)
    paycom.columns = [norm_colname(c) for c in paycom.columns]
    paycom = paycom.reset_index(drop=True)

    # Verify Keys
    UZIO_KEY = 'Employee ID'
    if UZIO_KEY not in uzio.columns:
        raise ValueError(f"Required column '{UZIO_KEY}' not found in Uzio file.")

    PAYCOM_KEY = norm_colname(PAYCOM_FIELD_MAP.get('Employee ID', 'Employee_Code'))
    if PAYCOM_KEY not in paycom.columns:
        raise ValueError(f"Required column '{PAYCOM_KEY}' not found in Paycom file.")

    # 0. Normalize IDs in the worksets
    uzio[UZIO_KEY] = uzio[UZIO_KEY].apply(norm_id)
    paycom[PAYCOM_KEY] = paycom[PAYCOM_KEY].apply(norm_id)

    # normalize display IDs (keep source formatting for display if needed, but match using norm_id)
    uzio_orig_keys = uzio[UZIO_KEY].astype(str).str.strip()
    paycom_orig_keys = paycom[PAYCOM_KEY].astype(str).str.strip()

    display_id_map = {}
    for norm_val, orig_val in zip(uzio[UZIO_KEY], uzio_orig_keys):
        n = str(norm_val).strip()
        o = str(orig_val).strip()
        if n and (n not in display_id_map or len(o) > len(display_id_map[n])):
            display_id_map[n] = o
    for norm_val, orig_val in zip(paycom[PAYCOM_KEY], paycom_orig_keys):
        n = str(norm_val).strip()
        o = str(orig_val).strip()
        if n and (n not in display_id_map or len(o) > len(display_id_map[n])):
            display_id_map[n] = o

    # Prepare mapping iteration
    # Iterate over PAYCOM_FIELD_MAP keys (Internal Standard Names usually match Uzio columns)
    # Filter out Keys
    mapped_fields = [f for f in PAYCOM_FIELD_MAP.keys() if f != UZIO_KEY]

    # employment status context map (prefer UZIO)
    uzio_emp_status_col = 'Employment Status'
    paycom_emp_status_col = norm_colname(PAYCOM_FIELD_MAP.get('Employment Status', ''))

    uzio_status_map = {}
    if uzio_emp_status_col is not None:
        tmp = uzio[[UZIO_KEY, uzio_emp_status_col]].copy()
        tmp[uzio_emp_status_col] = tmp[uzio_emp_status_col].map(norm_blank)
        tmp = tmp[tmp[UZIO_KEY] != ""]
        for _, r in tmp.iterrows():
            eid = str(r[UZIO_KEY]).strip()
            v = r[uzio_emp_status_col]
            if eid and norm_blank(v) != "" and eid not in uzio_status_map:
                uzio_status_map[eid] = str(v)

    paycom_status_map = {}
    if paycom_emp_status_col in paycom.columns:
        tmp = paycom[[PAYCOM_KEY, paycom_emp_status_col]].copy()
        tmp[paycom_emp_status_col] = tmp[paycom_emp_status_col].map(norm_blank)
        tmp = tmp[tmp[PAYCOM_KEY] != ""]
        for _, r in tmp.iterrows():
            eid = str(r[PAYCOM_KEY]).strip()
            v = r[paycom_emp_status_col]
            if eid and norm_blank(v) != "" and eid not in paycom_status_map:
                paycom_status_map[eid] = str(v)

    def get_emp_status(eid: str) -> str:
        eid = (eid or "").strip()
        if eid in uzio_status_map:
            return str(uzio_status_map[eid])
        if eid in paycom_status_map:
            return str(paycom_status_map[eid])
        return ""

    # pay type map (prefer UZIO)
    uzio_pay_type_col = 'Pay Type'
    paycom_pay_type_col = norm_colname(PAYCOM_FIELD_MAP.get('Pay Type', ''))

    pay_type_map = {}
    if uzio_pay_type_col in uzio.columns:
        tmp = uzio[[UZIO_KEY, uzio_pay_type_col]].copy()
        tmp[uzio_pay_type_col] = tmp[uzio_pay_type_col].map(norm_blank)
        tmp = tmp[tmp[UZIO_KEY] != ""]
        for _, r in tmp.iterrows():
            eid = str(r[UZIO_KEY]).strip()
            v = r[uzio_pay_type_col]
            if eid and norm_blank(v) != "" and eid not in pay_type_map:
                pay_type_map[eid] = canonical_pay_type(v)

    if paycom_pay_type_col in paycom.columns:
        tmp = paycom[[PAYCOM_KEY, paycom_pay_type_col]].copy()
        tmp[paycom_pay_type_col] = tmp[paycom_pay_type_col].map(norm_blank)
        tmp = tmp[tmp[PAYCOM_KEY] != ""]
        for _, r in tmp.iterrows():
            eid = str(r[PAYCOM_KEY]).strip()
            v = r[paycom_pay_type_col]
            if eid and norm_blank(v) != "" and eid not in pay_type_map:
                pay_type_map[eid] = canonical_pay_type(v)

    # index maps (keep first occurrence per employee)
    uzio_idx = {}
    for i, eid in uzio[UZIO_KEY].items():
        e = str(eid).strip()
        if e and e not in uzio_idx:
            uzio_idx[e] = i

    paycom_idx = {}
    for i, eid in paycom[PAYCOM_KEY].items():
        e = str(eid).strip()
        if e and e not in paycom_idx:
            paycom_idx[e] = i

    # ---------- FLSA Classification column (Uzio) ----------
    uzio_flsa_col = 'FLSA Classification'

    # Also locate employee name columns in Uzio for context in FLSA report
    uzio_fname_col = 'First Name'
    uzio_lname_col = 'Last Name'

    uzio_ssn_col = 'SSN'
    paycom_ssn_col = next((c for c in paycom.columns if "Tax ID (SSN)" in c or "SSN" in c), None)

    # 1. Resolve Identity Match Map (UZIO_ID -> PAYCOM_ID)
    uz_to_pc_id_map = get_identity_match_map(
        uzio, paycom, 
        uzio_id_col=UZIO_KEY, 
        vendor_id_col=PAYCOM_KEY,
        uzio_ssn_col=uzio_ssn_col,
        vendor_ssn_col=paycom_ssn_col
    )

    # 1.1 Data Quality Check: Duplicate SSNs
    uz_dupe_ssns = detect_duplicate_ssns(uzio, UZIO_KEY, uzio_ssn_col)
    pc_dupe_ssns = detect_duplicate_ssns(paycom, PAYCOM_KEY, paycom_ssn_col)
    
    dupe_ssn_rows = []
    for ssn, ids in uz_dupe_ssns.items():
        dupe_ssn_rows.append({
            "Source": "Uzio",
            "SSN": ssn,
            "Employee IDs": ", ".join(ids),
            "Issue": "Duplicate SSN found for multiple IDs in Uzio"
        })
    for ssn, ids in pc_dupe_ssns.items():
        dupe_ssn_rows.append({
            "Source": "Paycom",
            "SSN": ssn,
            "Employee IDs": ", ".join(ids),
            "Issue": "Duplicate SSN found for multiple IDs in Paycom"
        })
    df_dupe_ssns = pd.DataFrame(dupe_ssn_rows)

    pc_keys_processed = set()
    rows = []
    uzio_keys = sorted(set(uzio_idx.keys()))
    paycom_keys = sorted(set(paycom_idx.keys()))

    # 2. Main Loop: Iterate through all Uzio employees
    for uz_id in uzio_keys:
        pc_id = uz_to_pc_id_map.get(uz_id, uz_id)
        
        u_i = uzio_idx.get(uz_id)
        p_i = paycom_idx.get(pc_id) if pc_id else None
        
        if p_i is not None:
            pc_keys_processed.add(pc_id)

        emp_status_context = get_emp_status(uz_id)
        # Fix Paycom ID lookup
        pc_id_for_lookup = pc_id if pc_id else uz_id
        paycom_emp_status_val = str(paycom_status_map.get(pc_id_for_lookup, "")) 
        emp_pay_type = pay_type_map.get(uz_id, "")

        # --- Determine Employee Name ---
        fname = ""
        lname = ""
        if u_i is not None:
            if uzio_fname_col in uzio.columns:
                fname = str(norm_blank(uzio.at[u_i, uzio_fname_col]) or "")
            if uzio_lname_col in uzio.columns:
                lname = str(norm_blank(uzio.at[u_i, uzio_lname_col]) or "")
        
        if (fname == "" and lname == "") and p_i is not None:
             # Fallback to Paycom name cols
             pc_fname_col = next((c for c in paycom.columns if "First Name" in c), None)
             pc_lname_col = next((c for c in paycom.columns if "Last Name" in c), None)
             if pc_fname_col: fname = str(norm_blank(paycom.at[p_i, pc_fname_col]) or "")
             if pc_lname_col: lname = str(norm_blank(paycom.at[p_i, pc_lname_col]) or "")
        
        emp_name = f"{fname} {lname}".strip()

        # Check for ID Mismatch case (Same identity, different IDs)
        if pc_id and uz_id != pc_id:
            rows.append({
                "Employee ID": uz_id,
                "Employee Name": emp_name,
                "Field": "Employee ID Correlation",
                "Employment Status": emp_status_context,
                "Employment Status (Paycom)": paycom_emp_status_val,
                "UZIO_Value": uz_id,
                "PAYCOM_Value": pc_id,
                "PAYCOM_SourceOfTruth_Status": "Data Mismatch (Identity Match via SSN)"
            })

        for field in mapped_fields:
            uz_field = field 
            pc_col_raw = PAYCOM_FIELD_MAP.get(field)
            if not pc_col_raw: continue
            pc_col = norm_colname(pc_col_raw)

            uz_missing_col = (uz_field not in uzio.columns)
            pc_missing_col = (pc_col not in paycom.columns)

            uz_val = safe_val(uzio, u_i, uz_field) if not uz_missing_col else ""
            pc_val = safe_val(paycom, p_i, pc_col) if (p_i is not None and not pc_missing_col) else ""

            # Decide status
            if p_i is None and (pc_id is None or pc_id not in paycom_idx):
                status = "Employee ID Not Found in Paycom"
            elif pc_missing_col:
                status = "Column Missing in Paycom Sheet"
            elif uz_missing_col:
                status = "Column Missing in Uzio Sheet"
            else:
                if should_ignore_field_for_paytype(uz_field, emp_pay_type):
                    status = "Data Match"
                else:
                    same = normalized_compare(uz_field, uz_val, pc_val)
                    if same:
                        status = "Data Match"
                    else:
                        uz_b = norm_blank(uz_val)
                        pc_b = norm_blank(pc_val)
                        
                        f_case = norm_colname(uz_field).casefold()
                        if "employment status" in f_case and pc_b != "":
                            uz_stat = canonical_employment_status(uz_b)
                            pc_stat = canonical_employment_status(pc_b)
                            if "term" in uz_stat and "inactive" in pc_b.lower(): status = "Data Match"
                            elif "active" in uz_stat: status = "Active in Uzio"
                            elif "term" in uz_stat: status = "Terminated in Uzio"
                            elif uz_b == "" and "active" in pc_stat: status = "Active in Paycom"
                            elif uz_b == "" and ("term" in pc_stat or "retire" in pc_stat): status = "Terminated in Paycom"
                            else: status = "Data Mismatch"
                        else:
                            if (uz_b == "" or uz_b is None) and (pc_b != "" and pc_b is not None):
                                status = "Value missing in Uzio (Paycom has value)"
                            elif (uz_b != "" and uz_b is not None) and (pc_b == "" or pc_b is None):
                                status = "Value missing in Paycom (Uzio has value)"
                            else:
                                status = "Data Mismatch"

            rows.append({
                "Employee ID": uz_id,
                "Employee Name": emp_name,
                "Field": uz_field,
                "Employment Status": emp_status_context,
                "Employment Status (Paycom)": paycom_emp_status_val,
                "UZIO_Value": uz_val,
                "PAYCOM_Value": pc_val,
                "PAYCOM_SourceOfTruth_Status": status,
            })

    # 3. Final Loop: Remaining Paycom employees not in Uzio (even via identity)
    remaining_pc_ids = set(paycom_keys) - pc_keys_processed
    for pc_id in sorted(remaining_pc_ids):
        p_i = paycom_idx.get(pc_id)
        emp_status_context = get_emp_status(pc_id)
        paycom_emp_status_val = str(paycom_status_map.get(pc_id, ""))
        emp_pay_type = pay_type_map.get(pc_id, "")

        # Determine Name for Paycom-only records
        fname = ""
        lname = ""
        pc_fname_col = next((c for c in paycom.columns if "First Name" in c), None)
        pc_lname_col = next((c for c in paycom.columns if "Last Name" in c), None)
        if pc_fname_col: fname = str(norm_blank(paycom.at[p_i, pc_fname_col]) or "")
        if pc_lname_col: lname = str(norm_blank(paycom.at[p_i, pc_lname_col]) or "")
        emp_name = f"{fname} {lname}".strip()

        for field in mapped_fields:
            uz_field = field
            pc_col_raw = PAYCOM_FIELD_MAP.get(field)
            if not pc_col_raw: continue
            pc_col = norm_colname(pc_col_raw)
            pc_val = safe_val(paycom, p_i, pc_col) if pc_col in paycom.columns else ""

            rows.append({
                "Employee ID": pc_id,
                "Employee Name": emp_name,
                "Field": uz_field,
                "Employment Status": emp_status_context,
                "Employment Status (Paycom)": paycom_emp_status_val,
                "UZIO_Value": "",
                "PAYCOM_Value": pc_val,
                "PAYCOM_SourceOfTruth_Status": "Employee ID Not Found in Uzio"
            })

    comparison_detail = pd.DataFrame(rows, columns=[
        "Employee ID", "Employee Name", "Field", "Employment Status", "Employment Status (Paycom)",
        "UZIO_Value", "PAYCOM_Value", "PAYCOM_SourceOfTruth_Status",
    ])

    # ---------------- Salaried Driver Exceptions ----------------
    salaried_drivers_pc = []
    pc_pay_type_col = norm_colname(PAYCOM_FIELD_MAP.get('Pay Type', ''))
    pc_job_title_col = norm_colname(PAYCOM_FIELD_MAP.get('Job Title', ''))
    pc_flsa_col = norm_colname(PAYCOM_FIELD_MAP.get('FLSA Classification', ''))
    uzio_flsa_col_str = 'FLSA Classification'
    uzio_pay_type_col_str = 'Pay Type'
    uzio_emp_status_col_str = 'Employment Status'

    if pc_pay_type_col in paycom.columns and pc_job_title_col in paycom.columns:
        for idx_label, row in paycom.iterrows():
            pay_val = str(row[pc_pay_type_col]).strip().lower()
            if "salary" in pay_val or "salaried" in pay_val:
                jt_raw = row[pc_job_title_col]
                jt_val = str(jt_raw).strip().lower() if pd.notna(jt_raw) else ""
                if jt_val and jt_val != "nan":
                    if is_hourly_only_job_title(jt_val):
                        emp_id = str(row[PAYCOM_KEY]).strip()

                        # --- Paycom values ---
                        pc_pay_type_val = str(row[pc_pay_type_col]).strip()
                        pc_emp_status_val = str(paycom_status_map.get(emp_id, "Not Found")).strip()
                        pc_flsa_val = ""
                        if pc_flsa_col and pc_flsa_col in paycom.columns:
                            pc_flsa_val = str(norm_blank(row.get(pc_flsa_col, "")) or "").strip()

                        # --- Uzio values (look up via index) ---
                        uz_i = uzio_idx.get(emp_id)
                        uz_pay_type_val = ""
                        uz_emp_status_val = ""
                        uz_flsa_val = ""
                        if uz_i is not None:
                            if uzio_pay_type_col_str in uzio.columns:
                                uz_pay_type_val = str(norm_blank(safe_val(uzio, uz_i, uzio_pay_type_col_str)) or "").strip()
                            if uzio_emp_status_col_str in uzio.columns:
                                uz_emp_status_val = str(norm_blank(safe_val(uzio, uz_i, uzio_emp_status_col_str)) or "").strip()
                            if uzio_flsa_col_str in uzio.columns:
                                uz_flsa_val = str(norm_blank(safe_val(uzio, uz_i, uzio_flsa_col_str)) or "").strip()

                        # --- Build smart Comment ---
                        comment_parts = [
                            f"Paycom lists this employee as '{str(jt_raw).strip()}' with Pay Type '{pc_pay_type_val}'.",
                            "Uzio requires this job title to be Hourly/Non-Exempt — a Salaried assignment will cause a conflict.",
                        ]
                        if uz_emp_status_val:
                            comment_parts.append(f"Uzio status: {uz_emp_status_val}.")
                        if pc_emp_status_val and pc_emp_status_val != "Not Found":
                            comment_parts.append(f"Paycom status: {pc_emp_status_val}.")
                        if uz_flsa_val:
                            comment_parts.append(f"Uzio FLSA: {uz_flsa_val}.")
                        if pc_flsa_val:
                            comment_parts.append(f"Paycom FLSA: {pc_flsa_val}.")
                        if not uz_emp_status_val and uz_i is None:
                            comment_parts.append("Employee NOT found in Uzio — will need to be added as Hourly.")
                        comment = " ".join(comment_parts)

                        salaried_drivers_pc.append({
                            'Employee ID': emp_id,
                            'Job Title (Paycom)': str(row[pc_job_title_col]).strip(),
                            'Pay Type (Paycom)': pc_pay_type_val,
                            'Pay Type (Uzio)': uz_pay_type_val if uz_pay_type_val else "Not in Uzio",
                            'Employment Status (Paycom)': pc_emp_status_val,
                            'Employment Status (Uzio)': uz_emp_status_val if uz_emp_status_val else "Not in Uzio",
                            'FLSA Classification (Paycom)': pc_flsa_val if pc_flsa_val else "Blank",
                            'FLSA Classification (Uzio)': uz_flsa_val if uz_flsa_val else "Blank" if uz_i is not None else "Not in Uzio",
                            'Comment': comment
                        })
    df_salaried_drivers_pc = pd.DataFrame(salaried_drivers_pc)

    # ---------- FLSA Compliance Issues (4th sheet) ----------
    flsa_rows = []
    if uzio_flsa_col is not None:
        for eid, u_i in uzio_idx.items():
            # Get Pay Type from Uzio (raw value for display, canonical for comparison)
            pay_raw = ""
            if uzio_pay_type_col in uzio.columns:
                pay_raw = str(norm_blank(safe_val(uzio, u_i, uzio_pay_type_col)) or "").strip()
            pay_canon = pay_type_map.get(eid, "")

            # Get FLSA Classification from Uzio
            flsa_raw = ""
            if uzio_flsa_col in uzio.columns:
                flsa_raw = safe_val(uzio, u_i, uzio_flsa_col)
            flsa_norm = normalize_space_and_case(flsa_raw)

            # Get Paycom values for context
            pc_id = uz_to_pc_id_map.get(eid, eid)
            p_i = paycom_idx.get(pc_id)
            
            pc_pay_type = ""
            pc_flsa = ""
            pc_job = ""
            pc_dept = ""
            uzio_job = ""
            
            if p_i is not None:
                if pc_pay_type_col in paycom.columns:
                    pc_pay_type = str(norm_blank(safe_val(paycom, p_i, pc_pay_type_col)) or "").strip()
                if pc_flsa_col in paycom.columns:
                    pc_flsa = str(norm_blank(safe_val(paycom, p_i, pc_flsa_col)) or "").strip()
                if pc_job_title_col in paycom.columns:
                    pc_job = str(norm_blank(safe_val(paycom, p_i, pc_job_title_col)) or "").strip()
                if 'Department_Desc' in paycom.columns:
                    pc_dept = str(norm_blank(safe_val(paycom, p_i, 'Department_Desc')) or "").strip()
            
            if 'Job Title' in uzio.columns:
                uzio_job = str(norm_blank(safe_val(uzio, u_i, 'Job Title')) or "").strip()

            # Detect Issues
            all_issues = []
            
            # 1. Internal Uzio Inconsistency
            if pay_canon == "hourly" and "exempt" in flsa_norm and "non" not in flsa_norm:
                all_issues.append("Hourly employee classified as Exempt (Uzio Internal)")
            elif pay_canon == "salaried" and ("non-exempt" in flsa_norm or "non exempt" in flsa_norm or "nonexempt" in flsa_norm):
                all_issues.append("Salaried employee classified as Non-Exempt (Uzio Internal)")

            # 2. Cross-system Mismatches
            if p_i is not None:
                # Pay Type Mismatch
                pc_pt_canon = canonical_pay_type(pc_pay_type)
                if pay_canon != pc_pt_canon and pc_pt_canon != "":
                    all_issues.append(f"Pay Type Mismatch (Uzio: {pay_raw} vs Paycom: {pc_pay_type})")
                
                # FLSA Mismatch
                pc_flsa_canon = normalize_space_and_case(pc_flsa)
                uz_flsa_canon = normalize_space_and_case(flsa_raw)
                if uz_flsa_canon != pc_flsa_canon and pc_flsa_canon != "":
                    all_issues.append(f"FLSA Mismatch (Uzio: {flsa_raw} vs Paycom: {pc_flsa})")

            if all_issues:
                # Get employee name for context
                fname = ""
                lname = ""
                if uzio_fname_col and uzio_fname_col in uzio.columns:
                    fname = str(norm_blank(uzio.loc[u_i, uzio_fname_col]) or "")
                if uzio_lname_col and uzio_lname_col in uzio.columns:
                    lname = str(norm_blank(uzio.loc[u_i, uzio_lname_col]) or "")
                emp_name = f"{fname} {lname}".strip()

                flsa_rows.append({
                    "Employee ID": display_id_map.get(eid, eid),
                    "Employee Name": emp_name,
                    "Pay Type (Uzio)": pay_raw,
                    "Pay Type (Paycom)": pc_pay_type,
                    "FLSA Classification (Uzio)": str(norm_blank(flsa_raw) or ""),
                    "FLSA Classification (Paycom)": pc_flsa,
                    "Job Title (Uzio)": uzio_job,
                    "Job Title (Paycom)": pc_job,
                    "Department (Paycom)": pc_dept,
                    "Issue": "; ".join(all_issues),
                })

    flsa_issues = pd.DataFrame(flsa_rows, columns=[
        "Employee ID", "Employee Name", 
        "Pay Type (Uzio)", "Pay Type (Paycom)",
        "FLSA Classification (Uzio)", "FLSA Classification (Paycom)",
        "Job Title (Uzio)", "Job Title (Paycom)",
        "Department (Paycom)",
        "Issue"
    ])

    # ---------- Data Quality Issues (00/00/0000 dates) ----------
    dq_rows = []
    
    # Locate Emp ID and Name columns again for context
    pc_fname_col = find_col(paycom.columns, "Legal_Firstname", "Legal Firstname", "First Name", "FirstName")
    if pc_fname_col is None:
        for c in paycom.columns:
            cl = norm_colname(c).casefold()
            if "first" in cl and "name" in cl:
                pc_fname_col = c
                break
    pc_lname_col = find_col(paycom.columns, "Legal_Lastname", "Legal Lastname", "Last Name", "LastName")
    if pc_lname_col is None:
        for c in paycom.columns:
            cl = norm_colname(c).casefold()
            if "last" in cl and "name" in cl:
                pc_lname_col = c
                break

    for eid in paycom_idx.keys():
        p_i = paycom_idx.get(eid)
        if p_i is not None:
            # Check all columns for this row
            for col in paycom.columns:
                val = safe_val(paycom, p_i, col)
                if pd.notna(val) and '00/00/0000' in str(val):
                    fname = str(norm_blank(paycom.loc[p_i, pc_fname_col]) or "") if pc_fname_col else ""
                    lname = str(norm_blank(paycom.loc[p_i, pc_lname_col]) or "") if pc_lname_col else ""
                    emp_name = f"{fname} {lname}".strip()
                    pc_raw_id = paycom_orig_keys.loc[p_i] if p_i in paycom_orig_keys.index else eid
                    
                    dq_rows.append({
                        "Employee ID": str(pc_raw_id).strip(),
                        "Employee Name": emp_name,
                        "Column": col,
                        "Invalid Value Found": str(val)
                    })
                    
    dq_issues = pd.DataFrame(dq_rows, columns=[
        "Employee ID", "Employee Name", "Column", "Invalid Value Found"
    ])

    # ---------- Active Employees Missing in Uzio (5th sheet) ----------
    # Find employees in Paycom but NOT in Uzio who are Active / On Leave
    paycom_only_emps = set(paycom_idx.keys()) - set(uzio_idx.keys())

    # (Already located above for DQ checks, but redeclared here if needed, safe to keep as is)
    if pc_fname_col is None:
        for c in paycom.columns:
            cl = norm_colname(c).casefold()
            if "first" in cl and "name" in cl:
                pc_fname_col = c
                break
    if pc_lname_col is None:
        for c in paycom.columns:
            cl = norm_colname(c).casefold()
            if "last" in cl and "name" in cl:
                pc_lname_col = c
                break
    pc_hire_col = find_col(paycom.columns, "Most_Recent_Hire_Date", "Most Recent Hire Date",
                           "Hire_Date", "Hire Date")
    if pc_hire_col is None:
        for c in paycom.columns:
            cl = norm_colname(c).casefold()
            if "hire" in cl and "date" in cl:
                pc_hire_col = c
                break

    active_missing_rows = []
    for eid in sorted(paycom_only_emps):
        p_i = paycom_idx[eid]
        # Check employment status
        status_val = ""
        if paycom_emp_status_col and paycom_emp_status_col in paycom.columns:
            status_val = str(norm_blank(safe_val(paycom, p_i, paycom_emp_status_col)) or "")
        status_lower = status_val.strip().lower()

        # Only include Active / On Leave employees
        if status_lower not in {"active", "on leave", "leave", "activated"}:
            continue

        fname = ""
        lname = ""
        if pc_fname_col and pc_fname_col in paycom.columns:
            fname = str(norm_blank(paycom.loc[p_i, pc_fname_col]) or "")
        if pc_lname_col and pc_lname_col in paycom.columns:
            lname = str(norm_blank(paycom.loc[p_i, pc_lname_col]) or "")
        emp_name = f"{fname} {lname}".strip()

        hire_date = ""
        if pc_hire_col and pc_hire_col in paycom.columns:
            hire_date = str(norm_blank(paycom.loc[p_i, pc_hire_col]) or "")

        active_missing_rows.append({
            "Employee ID": display_id_map.get(eid, eid),
            "Employee Name": emp_name,
            "Employment Status (Paycom)": status_val,
            "Date of Hire (Paycom)": hire_date,
        })

    active_missing_in_uzio = pd.DataFrame(active_missing_rows, columns=[
        "Employee ID", "Employee Name",
        "Employment Status (Paycom)", "Date of Hire (Paycom)"
    ])

    terminated_missing_rows = []
    for eid in sorted(paycom_only_emps):
        p_i = paycom_idx[eid]
        # Check employment status
        status_val = ""
        if paycom_emp_status_col and paycom_emp_status_col in paycom.columns:
            status_val = str(norm_blank(safe_val(paycom, p_i, paycom_emp_status_col)) or "")
        status_lower = status_val.strip().lower()

        # Only include Terminated / Inactive employees
        is_term = False
        term_keywords = ["term", "inactive", "quit", "resign", "retired"]
        for kw in term_keywords:
            if kw in status_lower:
                is_term = True
                break
        
        if not is_term:
            continue

        fname = ""
        lname = ""
        if pc_fname_col and pc_fname_col in paycom.columns:
            fname = str(norm_blank(paycom.loc[p_i, pc_fname_col]) or "")
        if pc_lname_col and pc_lname_col in paycom.columns:
            lname = str(norm_blank(paycom.loc[p_i, pc_lname_col]) or "")
        emp_name = f"{fname} {lname}".strip()

        hire_date = ""
        if pc_hire_col and pc_hire_col in paycom.columns:
            hire_date = str(norm_blank(paycom.loc[p_i, pc_hire_col]) or "")

        terminated_missing_rows.append({
            "Employee ID": display_id_map.get(eid, eid),
            "Employee Name": emp_name,
            "Employment Status (Paycom)": status_val,
            "Date of Hire (Paycom)": hire_date,
        })

    terminated_missing_in_uzio = pd.DataFrame(terminated_missing_rows, columns=[
        "Employee ID", "Employee Name",
        "Employment Status (Paycom)", "Date of Hire (Paycom)"
    ])

    # Field summary
    statuses = [
        "Data Match",
        "Data Mismatch",
        "Value missing in Uzio (Paycom has value)",
        "Value missing in Paycom (Uzio has value)",
        "Employee ID Not Found in Uzio",
        "Employee ID Not Found in Paycom",
        "Column Missing in Paycom Sheet",
        "Column Missing in Uzio Sheet",
    ]

    if not comparison_detail.empty:
        field_summary_by_status = (
            comparison_detail.pivot_table(
                index="Field",
                columns="PAYCOM_SourceOfTruth_Status",
                values="Employee ID",
                aggfunc="count",
                fill_value=0,
            )
            .reindex(columns=statuses, fill_value=0)
            .reset_index()
        )
        field_summary_by_status["Total"] = field_summary_by_status[statuses].sum(axis=1)
    else:
        field_summary_by_status = pd.DataFrame(columns=["Field"] + statuses + ["Total"])

    # Summary
    uzio_emps = set(uzio[UZIO_KEY].dropna().map(str))
    paycom_emps = set(paycom[PAYCOM_KEY].dropna().map(str))

    # ---------- High Hourly Rate Anomalies (> $100/hr for hourly-only roles) ----------
    pc_hourly_rate_col = norm_colname(PAYCOM_FIELD_MAP.get('Hourly Pay Rate', ''))
    pc_job_col_hr = norm_colname(PAYCOM_FIELD_MAP.get('Job Title', ''))
    pc_fname_hr = norm_colname(PAYCOM_FIELD_MAP.get('First Name', ''))
    pc_lname_hr = norm_colname(PAYCOM_FIELD_MAP.get('Last Name', ''))
    HOURLY_RATE_THRESHOLD = 100.0
    high_rate_rows_pc = []

    if pc_hourly_rate_col in paycom.columns and pc_job_col_hr in paycom.columns:
        for idx_label, row in paycom.iterrows():
            jt_raw = row.get(pc_job_col_hr, '')
            jt_str = str(jt_raw).strip().lower() if pd.notna(jt_raw) else ""
            if not jt_str or jt_str == "nan":
                continue
            if is_hourly_only_job_title(jt_str):
                rate_raw = row.get(pc_hourly_rate_col, '')
                try:
                    rate = float(str(rate_raw).replace('$', '').replace(',', '').strip())
                except (ValueError, TypeError):
                    continue
                if rate > HOURLY_RATE_THRESHOLD:
                    emp_id = str(row[PAYCOM_KEY]).strip()
                    fname = str(norm_blank(row.get(pc_fname_hr, '')) or '').strip()
                    lname = str(norm_blank(row.get(pc_lname_hr, '')) or '').strip()
                    emp_name = f"{fname} {lname}".strip()
                    uz_i = uzio_idx.get(emp_id)
                    uz_rate = ""
                    if uz_i is not None and 'Hourly Pay Rate' in uzio.columns:
                        uz_rate = str(norm_blank(safe_val(uzio, uz_i, 'Hourly Pay Rate')) or '').strip()
                    high_rate_rows_pc.append({
                        'Employee ID': emp_id,
                        'Employee Name': emp_name,
                        'Job Title (Paycom)': str(jt_raw).strip(),
                        'Hourly Pay Rate (Paycom)': f"${rate:.2f}",
                        'Hourly Pay Rate (Uzio)': f"${float(uz_rate):.2f}" if uz_rate else "Not in Uzio",
                        'Comment': f"Hourly rate ${rate:.2f}/hr exceeds the ${HOURLY_RATE_THRESHOLD:.0f}/hr threshold for a '{str(jt_raw).strip()}' role. Please verify this is not a data entry error."
                    })
    df_high_rate_pc = pd.DataFrame(high_rate_rows_pc)

    # ---------- Hourly = 0 Hours validation (Check Uzio only) ----------
    hourly_zero_hours_rows = []
    if 'Pay Type' in uzio.columns and 'Working Hours' in uzio.columns:
        for u_i, row in uzio.iterrows():
            uz_pay_raw = str(norm_blank(row.get('Pay Type', '')) or "")
            # Assume canonical_pay_type returns 'hourly'
            if "hour" in str(uz_pay_raw).lower():
                wh_raw = row.get('Working Hours', '')
                try:
                    wh_val = float(str(wh_raw).replace(",", "").strip()) if str(wh_raw).strip() else 0.0
                except Exception:
                    wh_val = 0.0
                
                if wh_val > 0:
                    emp_id = str(row.get(UZIO_KEY, '')).strip()
                    fname = str(norm_blank(row.get('First Name', '')) or "")
                    lname = str(norm_blank(row.get('Last Name', '')) or "")
                    emp_name = f"{fname} {lname}".strip()
                    hourly_zero_hours_rows.append({
                        "Employee ID": emp_id,
                        "Employee Name": emp_name,
                        "Pay Type (Uzio)": uz_pay_raw,
                        "Working Hours (Uzio)": str(wh_raw),
                        "Issue": f"Hourly employee has {wh_raw} working hours. Must be 0."
                    })
    df_hourly_zero_hours = pd.DataFrame(hourly_zero_hours_rows)

    summary = pd.DataFrame(
        {
            "Metric": [
                "Total UZIO Employees",
                "Total PAYCOM Employees",
                "Employees in both",
                "Employees only in UZIO",
                "Employees only in PAYCOM",
                "Total UZIO Records",
                "Total PAYCOM Records",
                "Fields Compared",
                "Total Comparisons (field-level rows)",
                "FLSA Compliance Issues",
                "Active in Paycom but Missing in Uzio",
                "Terminated in Paycom but Missing in Uzio",
                "Data Quality Issues (00/00/0000)",
                "Duplicate SSN Warnings",
                "Salaried Hourly-Only Exceptions",
                "High Hourly Rate Anomalies (>$100/hr)",
                "Hourly Zero Hours Exceptions",
            ],
            "Value": [
                len(uzio_emps),
                len(paycom_emps),
                len(uzio_emps & paycom_emps),
                len(uzio_emps - paycom_emps),
                len(paycom_emps - uzio_emps),
                int(len(uzio)),
                int(len(paycom)),
                int(len(PAYCOM_FIELD_MAP)),
                int(comparison_detail.shape[0]),
                len(flsa_rows),
                len(active_missing_rows),
                len(terminated_missing_rows),
                len(dq_rows),
                len(dupe_ssn_rows),
                len(df_salaried_drivers_pc),
                len(df_high_rate_pc),
                len(df_hourly_zero_hours),
            ],
        }
    )

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
        if not df_salaried_drivers_pc.empty:
            df_salaried_drivers_pc.to_excel(writer, sheet_name="Salaried_Driver_Exceptions", index=False)
        if not df_high_rate_pc.empty:
            df_high_rate_pc.to_excel(writer, sheet_name="High_Hourly_Rate_Anomalies", index=False)
        if not df_hourly_zero_hours.empty:
            df_hourly_zero_hours.to_excel(writer, sheet_name="Hourly_Zero_Hours_Exceptions", index=False)

    return out.getvalue()

# ---------- UI ----------
def render_ui():
    st.title(APP_TITLE)
    st.markdown("""
    **Instructions**:
    1. Upload **Uzio Census Export** (.xlsm).
    2. Upload **Paycom Census Export** (.csv or .xlsx).
    
    **Output Reports**:
    - **Comparison**: Discrepancies between Uzio and Paycom.
    - **FLSA_Compliance_Issues**: Flags employees where 'FLSA Status' does not match their assigned 'Pay Type' constraints.
    - **Active_Missing_In_Uzio**: Active employees found in Paycom but genuinely missing from the Uzio census entirely.
    - **Data_Quality_Issues**: Identifies unexpected placeholder dates such as '00/00/0000'.
    - **Salaried_Driver_Exceptions**: Employees mapped as salaried drivers, which are incompatible.
    """)

    uzio_file = st.file_uploader("Upload Uzio Census Export (.xlsm)", type=["xlsm"])
    paycom_file = st.file_uploader("Upload Paycom Census Export (.csv or .xlsx)", type=["csv", "xlsx"])
    
    client_name = st.text_input("Client Name", value="Client", key="paycom_census_client")

    if st.button("Run Audit", type="primary", disabled=(not uzio_file or not paycom_file)):
        try:
            with st.spinner("Running audit..."):
                out_excel = run_comparison(uzio_file, paycom_file)
            st.success("Audit Complete!")
            timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
            filename = f"{client_name}_Uzio_Paycom_Census_Audit_Report_{timestamp}.xlsx"

            st.download_button(
                label="Download Audit Report",
                data=out_excel,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error during audit: {e}")
            # Add error logging logic here if requested
            print(f"ERROR: {e}")

if __name__ == "__main__":
    st.set_page_config(page_title=APP_TITLE, layout="centered", initial_sidebar_state="collapsed")
    render_ui()
