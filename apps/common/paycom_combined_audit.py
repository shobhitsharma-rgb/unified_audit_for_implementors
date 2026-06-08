import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime, date
from utils.audit_utils import (
    norm_col, norm_colname, norm_blank, try_parse_date, clean_money_val, 
    get_identity_match_map, norm_ssn_canonical, detect_duplicate_ssns, norm_id
)

# =========================================================
# Paycom Consolidated Audit Tool (Census, Payment, Emergency)
# =========================================================

APP_TITLE = "Paycom - Consolidated Audit (Census/Payment/Emergency)"

# --- Status Constants ---
STATUS_MATCH = "Data Match"
STATUS_MISMATCH = "Data Mismatch"
STATUS_VAL_MISSING_UZIO = "Value missing in Uzio"
STATUS_VAL_MISSING_PAYCOM = "Value missing in Paycom"
STATUS_MISSING_UZIO = "Employee ID Not Found in Uzio"
STATUS_MISSING_PAYCOM = "Employee ID Not Found in Paycom"

def norm_str(x):
    if pd.isna(x) or x is None:
        return ""
    return str(x).strip()


def norm_phone(x):
    """Normalize phone to just digits."""
    if pd.isna(x): return ""
    digits = re.sub(r"\D", "", str(x))
    if digits.startswith("1") and len(digits) == 11:
        digits = digits[1:]
    return digits

def norm_money(x):
    """Parse money/float safely."""
    if pd.isna(x) or x is None:
        return 0.0
    s = str(x).replace(",", "").replace("$", "").strip()
    try:
        return float(s)
    except:
        return 0.0

def norm_relation(x):
    """Normalize relationship (uppercase, strip)."""
    return str(x).strip().upper()

# --- Uzio Master Reader ---
def read_uzio_master(file):
    """
    Reads the Uzio Master CSV which has Category labels in Row 1 
    and actual Headers in Row 2.
    """
    # Read first two rows to build combined headers
    df_headers = pd.read_csv(io.StringIO(file.getvalue().decode('utf-8', errors='replace')), nrows=2, header=None)
    
    row1 = df_headers.iloc[0].ffill().tolist() # Categories
    row2 = df_headers.iloc[1].fillna('').tolist() # Headers
    
    # Combined headers for easy lookup
    combined_cols = []
    for c, h in zip(row1, row2):
        if h:
            combined_cols.append(f"{c}|{h}")
        else:
            combined_cols.append(c)
            
    # Read full data
    file.seek(0)
    df = pd.read_csv(file, skiprows=2, header=None, dtype=str)
    df.columns = combined_cols
    return df


def company_name_from_uzio_master(df):
    """Return the client/company name carried in a parsed Uzio Master Report, or
    "" if absent. The report's 'Company Information|Company Name' column repeats
    the same company on every employee row, so take the first non-empty value."""
    if df is None or df.empty:
        return ""
    col = next((c for c in df.columns
                if str(c).strip().lower().endswith("company name")), None)
    if col is None:
        return ""
    s = df[col].dropna().astype(str).str.strip()
    s = s[s != ""]
    return s.iloc[0] if len(s) else ""

# --- Field Mappings ---
# --- Field Mappings ---
PAYCOM_CENSUS_MAP = {
    'Personal|First Name': 'Legal_Firstname',
    'Personal|Last Name': 'Legal_Lastname',
    'Personal|Middle Name': 'Legal_Middle_Name',
    'Personal|Suffix': 'Legal_Employee_Suffix',
    'Personal|SSN': 'SS_Number',
    'Personal|Date Of Birth': 'Birth_Date_(MM/DD/YYYY)',
    'Personal|Gender': 'Gender',
    'Job|Employee ID': 'Employee_Code',
    'Job|Date of Hire': 'Most_Recent_Hire_Date',
    'Job|Original DOH': 'Hire_Date',
    'Job|Status': 'Employee_Status',
    'Job|Employment Type': 'DOL_Status',
    'Job|Pay Type': 'Pay_Type',
    'Job|Annual Salary': 'Annual_Salary',
    'Job|Hourly Rate': 'Rate_1',
    'Job|Working Hours per Week': 'Scheduled_Pay_Period_Hours',
    'Job|Job Title': 'Position',
    'Job|Department': 'Department_Desc',
    'Personal|Work Email': 'Work_Email',
    'Home Address|Personal Email': 'Personal_Email',
    'Home Address|Phone': 'Primary_Phone',
    'Home Address|Address Line 1': 'Primary_Address_Line_1',
    'Home Address|Address Line 2': 'Primary_Address_Line_2',
    'Home Address|City': 'Primary_City/Municipality',
    'Home Address|Zip': 'Primary_Zip/Postal_Code',
    'Home Address|State': 'Primary_State/Province',
    'Additional Information|License Number': 'DriversLicense',
    'Additional Information|License Expiration Date': 'DLExpirationDate',
    'Job|Work Location': 'Work_Location',
    'Job|Reporting Manager': 'Supervisor_Primary_Code',
    'Job|Race/Ethnicity': 'EEO1_Ethnicity',
    'Job|EEO Job Category': 'SOC_Code',
    'Job|Termination Date': 'Termination_Date',
    'Job|Termination Reason': 'Termination_Reason',
    'Personal|Tobacco Usage': 'Tobacco_User',
    'Job|FLSA Classification': 'Exempt_Status',
    'Mailing Address|Address Line 1': 'Mailing_Address_Line_1',
    'Mailing Address|Address Line 2': 'Mailing_Address_Line_2',
    'Mailing Address|City': 'Mailing_City/Municipality',
    'Mailing Address|Zip': 'Mailing_Zip/Postal_Code',
    'Mailing Address|State': 'Mailing_State/Province'
}

# --- Normalization Helpers (Ported from census_audit.py) ---

def normalize_employment_type(x):
    s = norm_colname(x).casefold()
    s = s.replace("-", " ")
    s = re.sub(r"\s+", " ", s).strip()
    if s in {"full time", "fulltime", "ft"}:
        return "full time"
    if s in {"part time", "parttime", "pt"}:
        return "part time"
    if s in {"seasonal", "temporary", "temp"}:
        return "seasonal"
    return s

# --- Anomaly Detection Constants ---
HOURLY_ONLY_JOB_TITLES = {
    "driver", "delivery driver", "truck driver", "warehouse", 
    "warehouse worker", "material handler", "forklift operator"
}

def is_hourly_only_job_title(title):
    t = str(title or "").lower().strip()
    return any(keyword in t for keyword in HOURLY_ONLY_JOB_TITLES)

def normalize_suffix(x):
    s = norm_colname(x).casefold()
    s = re.sub(r"[^a-z0-9]", "", s)
    return s

def first_alpha_char(x):
    s = norm_blank(x)
    if not s: return ""
    txt = str(s).strip()
    m = re.search(r"[A-Za-z]", txt)
    return m.group(0).casefold() if m else ""

def normalize_middle_initial(uzio_val, paycom_val):
    u = first_alpha_char(uzio_val)
    p = first_alpha_char(paycom_val)
    return u != "" and p != "" and u == p

def canonical_pay_type(x):
    s = norm_colname(x).casefold()
    if not s: return ""
    if "hour" in s: return "hourly"
    if "salar" in s: return "salaried"
    return s

def canonical_employment_status(x):
    s = norm_colname(x).casefold()
    if not s: return ""
    if "on leave" in s: return "active"
    if s in {"active", "activated"}: return "active"
    if "term" in s or "inactive" in s or "quit" in s or "resign" in s: return "terminated"
    return s

def termination_reason_equal(uzio_val, paycom_val):
    uz = norm_colname(uzio_val).casefold()
    pc = norm_colname(paycom_val).casefold()
    if not uz and not pc: return True
    if uz == "other": return True
    if ("involuntary" in uz) or ("involuntary" in pc):
        return ("involuntary" in uz) and ("involuntary" in pc)
    if ("voluntary" in uz) or ("voluntary" in pc):
        return ("voluntary" in uz) and ("voluntary" in pc)
    return uz == pc

def should_ignore_field_for_paytype(field_name, pay_type_canon):
    f = norm_colname(field_name).casefold()
    pt = (pay_type_canon or "").casefold()
    if pt == "hourly":
        if "annual salary" in f: return True
    if pt == "salaried":
        if ("hourly" in f and "rate" in f) or ("hours per week" in f) or ("working hours" in f):
            return True
    return False

def normalized_compare(field_name, uzio_val, paycom_val):
    f = norm_colname(field_name).casefold()
    if "termination reason" in f:
        return termination_reason_equal(uzio_val, paycom_val)
    if "employment status" in f:
        u_stat = canonical_employment_status(uzio_val)
        p_stat = canonical_employment_status(paycom_val)
        return u_stat == p_stat
    if "pay type" in f:
        return canonical_pay_type(uzio_val) == canonical_pay_type(paycom_val)
    if "employment type" in f:
        return normalize_employment_type(uzio_val) == normalize_employment_type(paycom_val)
    if "middle" in f and "initial" in f:
        return normalize_middle_initial(uzio_val, paycom_val)
    if "suffix" in f:
        return normalize_suffix(uzio_val) == normalize_suffix(paycom_val)
    if "ssn" in f:
        u = re.sub(r"\D", "", str(uzio_val)).lstrip("0")
        p = re.sub(r"\D", "", str(paycom_val)).lstrip("0")
        return u == p
    if "phone" in f:
        u = norm_phone(uzio_val).lstrip("0")
        p = norm_phone(paycom_val).lstrip("0")
        return u == p
    if "zip" in f:
        u = re.sub(r"\D", "", str(uzio_val)).lstrip("0")
        p = re.sub(r"\D", "", str(paycom_val)).lstrip("0")
        return u == p
    if any(k in f for k in ["date", "dob", "birth", "effective", "doh", "hire", "termination"]):
        return try_parse_date(uzio_val) == try_parse_date(paycom_val)
    if any(k in f for k in ["salary", "rate", "hours", "amount", "percent"]):
        try:
            fa = float(str(uzio_val).replace(",","").replace("$","") or 0)
            fb = float(str(paycom_val).replace(",","").replace("$","") or 0)
            return abs(fa - fb) <= 1e-6
        except:
            return norm_colname(uzio_val).casefold() == norm_colname(paycom_val).casefold()
    if "license" in f:
        u = str(uzio_val).strip().lstrip("0")
        p = str(paycom_val).strip().lstrip("0")
        return u == p
    return norm_colname(uzio_val).casefold() == norm_colname(paycom_val).casefold()

# --- Payment Helpers (Ported from payment_audit.py) ---

_TYPE_CODE_MAP = {
    "22": "checking",
    "32": "savings",
    "1": "checking",
    "2": "checking",
}

def norm_digits(x):
    if pd.isna(x): return ""
    if isinstance(x, (float, int)):
        return str(int(x))
    return re.sub(r"\D", "", str(x))

def strip_type(t):
    if not t: return ""
    s = str(t).strip()
    if s.endswith(".0"): s = s[:-2]
    if s in _TYPE_CODE_MAP: return _TYPE_CODE_MAP[s]
    return s.lower().replace("account", "").replace("code: ", "").strip()

def _get_field_val(acc, field):
    mapping = {
        "Routing Number": "Routing",
        "Account Number": "Account",
        "Account Type": "Type",
        "Amount": "Amount",
        "Percent": "Percent"
    }
    val = acc.get(mapping.get(field, ""), "")
    return str(val) if val != "" else ""

def _compare_field(field, u_val, p_val, u_acc, p_acc):
    u_n = str(u_val).strip()
    p_n = str(p_val).strip()
    if not u_n and not p_n: return STATUS_MATCH
    if not u_n and p_n: return "Value missing in Uzio (Paycom has value)"
    if u_n and not p_n: return "Value missing in Paycom (Uzio has value)"
    if field == "Account Type":
        if strip_type(u_n) == strip_type(p_n): return STATUS_MATCH
        return STATUS_MISMATCH
    if field in ("Amount", "Percent"):
        try:
            diff = abs(float(u_n) - float(p_n))
            if diff < 0.01: return STATUS_MATCH
        except: pass
        return STATUS_MISMATCH
    return STATUS_MATCH if u_n == p_n else STATUS_MISMATCH

def run_census_audit(df_uzio, df_paycom, uz_to_pc_id_map=None):
    rows = []
    
    if uz_to_pc_id_map is None:
        u_id_col = "Job|Employee ID"
        p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
        uzio_ssn_col = 'Personal|SSN'
        paycom_ssn_col = next((c for c in df_paycom.columns if "SS_Number" in c or "SSN" in c), "SS_Number")
        
        uz_to_pc_id_map = get_identity_match_map(
            df_uzio, df_paycom, 
            uzio_id_col=u_id_col, 
            vendor_id_col=p_id_col,
            uzio_ssn_col=uzio_ssn_col,
            vendor_ssn_col=paycom_ssn_col
        )

    u_id_col = "Job|Employee ID"
    p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")

    # Pre-calculate status and pay type maps for context
    uzio_status_map = {}
    if "Job|Status" in df_uzio.columns:
        for _, r in df_uzio[[u_id_col, "Job|Status"]].dropna().iterrows():
            uzio_status_map[norm_id(r[u_id_col])] = str(r["Job|Status"])

    paycom_status_map = {}
    p_stat_col = next((c for c in df_paycom.columns if "Employee_Status" in c), "Employee_Status")
    if p_stat_col in df_paycom.columns:
        for _, r in df_paycom[[p_id_col, p_stat_col]].dropna().iterrows():
            paycom_status_map[norm_id(r[p_id_col])] = str(r[p_stat_col])

    pay_type_map = {}
    if "Job|Pay Type" in df_uzio.columns:
        for _, r in df_uzio[[u_id_col, "Job|Pay Type"]].dropna().iterrows():
            pay_type_map[norm_id(r[u_id_col])] = canonical_pay_type(r["Job|Pay Type"])

    u_map = {id: idx for idx, id in enumerate(df_uzio[u_id_col].map(norm_id))}
    p_map = {id: idx for idx, id in enumerate(df_paycom[p_id_col].map(norm_id))}
    
    pc_keys_processed = set()
    uzio_keys = sorted(set(u_map.keys()))
    paycom_keys = sorted(set(p_map.keys()))

    # 2. Main Loop: Iterate through all Uzio employees
    for uz_id in uzio_keys:
        if not uz_id or uz_id == "nan": continue
        pc_id = uz_to_pc_id_map.get(uz_id, uz_id)
        
        u_idx = u_map.get(uz_id)
        p_idx = p_map.get(pc_id) if pc_id in p_map else None
        
        if p_idx is not None:
            pc_keys_processed.add(pc_id)

        # Context
        u_status = uzio_status_map.get(uz_id, "")
        p_status = paycom_status_map.get(pc_id if pc_id else uz_id, "")
        emp_pay_type = pay_type_map.get(uz_id, "")
        is_terminated_context = "terminated" in canonical_employment_status(u_status)

        fname = norm_blank(df_uzio.at[u_idx, 'Personal|First Name'])
        lname = norm_blank(df_uzio.at[u_idx, 'Personal|Last Name'])
        name = f"{fname} {lname}".strip()

        # Check for ID Mismatch case (Same identity, different IDs)
        if pc_id and uz_id != pc_id:
            rows.append({
                "Employee ID": uz_id,
                "Employee Name": name,
                "Section": "Census",
                "Field": "Employee ID Correlation",
                "Uzio Value": uz_id,
                "Paycom Value": pc_id,
                "Status": "Data Mismatch (Identity Match via SSN)"
            })

        for u_col, p_col in PAYCOM_CENSUS_MAP.items():
            u_val = norm_str(df_uzio.at[u_idx, u_col]) if u_col in df_uzio.columns else ""
            p_val = ""
            if p_idx is not None and p_col in df_paycom.columns:
                p_val = norm_str(df_paycom.at[p_idx, p_col])
                if p_col == "Position" and p_val == "":
                    for alt in ["Business_Title", "Job_Title_Description"]:
                        if alt in df_paycom.columns:
                            alt_val = norm_str(df_paycom.at[p_idx, alt])
                            if alt_val:
                                p_val = alt_val
                                break

            # Decide status
            if p_idx is None:
                status = STATUS_MISSING_PAYCOM
            elif p_col not in df_paycom.columns:
                status = "Column Missing in Paycom Sheet"
            elif u_col not in df_uzio.columns:
                status = "Column Missing in Uzio Sheet"
            else:
                if should_ignore_field_for_paytype(u_col, emp_pay_type):
                    status = STATUS_MATCH
                else:
                    if normalized_compare(u_col, u_val, p_val):
                        status = STATUS_MATCH
                    else:
                        u_b = norm_blank(u_val)
                        p_b = norm_blank(p_val)
                        f_case = norm_colname(u_col).casefold()
                        if "employment status" in f_case and p_b != "":
                            uz_stat = canonical_employment_status(u_b)
                            pc_stat = canonical_employment_status(p_b)
                            if "terminated" in uz_stat and "terminated" in pc_stat: status = STATUS_MATCH
                            elif "active" in uz_stat: status = "Active in Uzio"
                            elif "terminated" in uz_stat: status = "Terminated in Uzio"
                            elif not u_b and "active" in pc_stat: status = "Active in Paycom"
                            else: status = STATUS_MISMATCH
                        elif not u_b and p_b: status = "Value missing in Uzio (Paycom has value)"
                        elif u_b and not p_b: status = "Value missing in Paycom (Uzio has value)"
                        else: status = STATUS_MISMATCH

            rows.append({
                "Employee ID": uz_id,
                "Employee Name": name,
                "Section": "Census",
                "Field": u_col.split("|")[-1],
                "Uzio Value": u_val,
                "Paycom Value": p_val,
                "Status": status
            })

    # 3. Final Loop: Remaining Paycom employees not in Uzio (even via identity)
    remaining_pc_ids = set(paycom_keys) - pc_keys_processed
    for pc_id in sorted(remaining_pc_ids):
        if not pc_id or pc_id == "nan": continue
        p_idx = p_map.get(pc_id)
        name = f"{norm_blank(df_paycom.at[p_idx, 'Legal_Firstname'])} {norm_blank(df_paycom.at[p_idx, 'Legal_Lastname'])}".strip()
        p_status = paycom_status_map.get(pc_id, "")

        for u_col, p_col in PAYCOM_CENSUS_MAP.items():
            p_val = norm_str(df_paycom.at[p_idx, p_col]) if p_col in df_paycom.columns else ""
            rows.append({
                "Employee ID": pc_id,
                "Employee Name": name,
                "Section": "Census",
                "Field": u_col.split("|")[-1],
                "Uzio Value": "",
                "Paycom Value": p_val,
                "Status": STATUS_MISSING_UZIO
            })
    return pd.DataFrame(rows)

# --- Anomaly Extraction Functions ---

def get_salaried_driver_exceptions(df_uzio, df_paycom):
    exceptions = []
    p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
    p_jt_col = next((c for c in df_paycom.columns if "Position" in c), "Position")
    p_pt_col = next((c for c in df_paycom.columns if "Pay_Type" in c), "Pay_Type")
    
    for idx, row in df_paycom.iterrows():
        jt = str(row.get(p_jt_col, "")).strip()
        pt = str(row.get(p_pt_col, "")).strip().lower()
        if "salary" in pt and is_hourly_only_job_title(jt):
            eid = norm_id(row.get(p_id_col))
            exceptions.append({
                "Employee ID": eid,
                "Employee Name": f"{row.get('Legal_Firstname', '')} {row.get('Legal_Lastname', '')}".strip(),
                "Job Title (Paycom)": jt,
                "Pay Type (Paycom)": pt.capitalize(),
                "Issue": "Salaried employee in hourly-only role"
            })
    return pd.DataFrame(exceptions)

def get_flsa_compliance_issues(df_uzio):
    issues = []
    u_id_col = "Job|Employee ID"
    u_pt_col = "Job|Pay Type"
    u_flsa_col = "Job|FLSA Classification"
    
    if all(c in df_uzio.columns for c in [u_id_col, u_pt_col, u_flsa_col]):
        for idx, row in df_uzio.iterrows():
            pt = canonical_pay_type(row.get(u_pt_col))
            flsa = str(row.get(u_flsa_col, "")).lower()
            issue = ""
            if pt == "hourly" and "exempt" in flsa and "non" not in flsa:
                issue = "Hourly employee classified as Exempt"
            elif pt == "salaried" and "non-exempt" in flsa:
                issue = "Salaried employee classified as Non-Exempt"
            
            if issue:
                issues.append({
                    "Employee ID": norm_id(row.get(u_id_col)),
                    "Employee Name": f"{row.get('Personal|First Name', '')} {row.get('Personal|Last Name', '')}".strip(),
                    "Pay Type": pt.capitalize(),
                    "FLSA Classification": row.get(u_flsa_col),
                    "Issue": issue
                })
    return pd.DataFrame(issues)

def get_active_missing_in_uzio(df_uzio, df_paycom):
    missing = []
    u_id_col = "Job|Employee ID"
    p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
    p_stat_col = next((c for c in df_paycom.columns if "Employee_Status" in c), "Employee_Status")
    
    # Use identity map to find which Paycom IDs are actually represented in Uzio
    uzio_ssn_col = 'Personal|SSN'
    paycom_ssn_col = next((c for c in df_paycom.columns if "SS_Number" in c or "SSN" in c), "SS_Number")
    
    uz_to_pc_map = get_identity_match_map(
        df_uzio, df_paycom,
        uzio_id_col=u_id_col,
        vendor_id_col=p_id_col,
        uzio_ssn_col=uzio_ssn_col,
        vendor_ssn_col=paycom_ssn_col
    )
    uzio_id_set = set(df_uzio[u_id_col].map(norm_id)) if u_id_col in df_uzio.columns else set()
    pc_ids_in_uzio = set(uz_to_pc_map.values()) | uzio_id_set

    for idx, row in df_paycom.iterrows():
        eid = norm_id(row.get(p_id_col))
        if eid and eid not in pc_ids_in_uzio:
            stat = canonical_employment_status(row.get(p_stat_col))
            if stat == "active":
                missing.append({
                    "Employee ID": eid,
                    "Employee Name": f"{row.get('Legal_Firstname', '')} {row.get('Legal_Lastname', '')}".strip(),
                    "Status (Paycom)": row.get(p_stat_col),
                    "Hire Date": row.get('Most_Recent_Hire_Date', '')
                })
    return pd.DataFrame(missing)

def get_terminated_missing_in_uzio(df_uzio, df_paycom):
    missing = []
    u_id_col = "Job|Employee ID"
    p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
    p_stat_col = next((c for c in df_paycom.columns if "Employee_Status" in c), "Employee_Status")
    
    # Use identity map to find which Paycom IDs are actually represented in Uzio
    uzio_ssn_col = 'Personal|SSN'
    paycom_ssn_col = next((c for c in df_paycom.columns if "SS_Number" in c or "SSN" in c), "SS_Number")
    
    uz_to_pc_map = get_identity_match_map(
        df_uzio, df_paycom,
        uzio_id_col=u_id_col,
        vendor_id_col=p_id_col,
        uzio_ssn_col=uzio_ssn_col,
        vendor_ssn_col=paycom_ssn_col
    )
    uzio_id_set = set(df_uzio[u_id_col].map(norm_id)) if u_id_col in df_uzio.columns else set()
    pc_ids_in_uzio = set(uz_to_pc_map.values()) | uzio_id_set

    for idx, row in df_paycom.iterrows():
        eid = norm_id(row.get(p_id_col))
        if eid and eid not in pc_ids_in_uzio:
            stat = canonical_employment_status(row.get(p_stat_col))
            if stat == "terminated":
                missing.append({
                    "Employee ID": eid,
                    "Employee Name": f"{row.get('Legal_Firstname', '')} {row.get('Legal_Lastname', '')}".strip(),
                    "Status (Paycom)": row.get(p_stat_col),
                    "Hire Date": row.get('Most_Recent_Hire_Date', '')
                })
    return pd.DataFrame(missing)

def get_data_quality_issues(df_paycom):
    issues = []
    p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
    for idx, row in df_paycom.iterrows():
        eid = norm_id(row.get(p_id_col))
        for col in df_paycom.columns:
            val = str(row[col])
            if "00/00/0000" in val:
                issues.append({
                    "Employee ID": eid,
                    "Employee Name": f"{row.get('Legal_Firstname', '')} {row.get('Legal_Lastname', '')}".strip(),
                    "Column": col,
                    "Issue": "Invalid Date Placeholder (00/00/0000)"
                })
    return pd.DataFrame(issues)

def get_high_rate_anomalies(df_paycom):
    anomalies = []
    p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
    p_rate_col = next((c for c in df_paycom.columns if "Rate_1" in c), "Rate_1")
    p_jt_col = next((c for c in df_paycom.columns if "Position" in c), "Position")
    
    for idx, row in df_paycom.iterrows():
        rate = norm_money(row.get(p_rate_col))
        jt = str(row.get(p_jt_col, "")).strip()
        if rate > 100.0 and is_hourly_only_job_title(jt):
            anomalies.append({
                "Employee ID": norm_id(row.get(p_id_col)),
                "Employee Name": f"{row.get('Legal_Firstname', '')} {row.get('Legal_Lastname', '')}".strip(),
                "Job Title": jt,
                "Rate": f"${rate:.2f}/hr",
                "Issue": "Hourly rate > $100/hr"
            })
    return pd.DataFrame(anomalies)

def run_payment_audit(df_uzio, df_paycom, uz_to_pc_id_map=None):
    rows = []
    
    # Pre-process identity map for reverse lookup
    pc_to_uz_id_map = {}
    if uz_to_pc_id_map:
        pc_to_uz_id_map = {pc: uz for uz, pc in uz_to_pc_id_map.items() if pc}

    u_id_col = "Job|Employee ID"
    p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
    
    # Process Uzio Accounts
    u_map = {}
    uzio_emp_names = {}
    for idx, row in df_uzio.iterrows():
        eid = norm_id(row.get(u_id_col))
        if not eid: continue
        
        name = f"{norm_blank(row.get('Personal|First Name'))} {norm_blank(row.get('Personal|Last Name'))}".strip()
        uzio_emp_names[eid] = name
        
        acc = {
            "Routing": norm_digits(row.get("Payment Method|Routing Number")).lstrip("0"),
            "Account": norm_digits(row.get("Payment Method|Account Number")).lstrip("0"),
            "Type": norm_str(row.get("Payment Method|Account Type")),
            "Percent": norm_money(row.get("Payment Method|Paycheck Percentage")),
            "Amount": norm_money(row.get("Payment Method|Paycheck Amount")),
            "Name": name
        }
        if acc["Routing"] or acc["Account"]:
            if eid not in u_map: u_map[eid] = []
            if acc not in u_map[eid]: u_map[eid].append(acc)

    # Process Paycom Accounts (Unpivot)
    p_map = {}
    for idx, row in df_paycom.iterrows():
        eid = norm_id(row.get(p_id_col))
        if not eid: continue
        
        # Translate Paycom ID to Uzio ID if mapping exists
        mapped_eid = pc_to_uz_id_map.get(eid, eid)
        
        accs = []
        total_dist_pct = 0.0
        # Distributions 1-8
        for i in range(1, 9):
            prefix = f"Dist_{i}_"
            d_acc = norm_digits(row.get(f"{prefix}Acct_Code")).lstrip("0")
            d_rout = norm_digits(row.get(f"{prefix}Rout_Code")).lstrip("0")
            
            raw_amt_val = row.get(f"{prefix}Amount")
            d_amt = norm_money(raw_amt_val)
            d_pct = 0.0
            
            if f"{prefix}Percent" in df_paycom.columns:
                d_pct = norm_money(row.get(f"{prefix}Percent"))
            
            if d_pct == 0.0:
                raw_str = str(raw_amt_val).strip()
                if "%" in raw_str:
                    try: d_pct = float(raw_str.replace("%", "").replace(",", "").strip())
                    except: pass
                    d_amt = 0.0
                elif d_amt != 0.0 and 0.01 < abs(d_amt) <= 1.0:
                    d_pct = round(d_amt * 100, 4)
                    d_amt = 0.0
            
            total_dist_pct += d_pct
            if d_acc or d_rout:
                acc_type = row.get(f"{prefix}Type_Code")
                accs.append({
                    "Routing": d_rout, "Account": d_acc, 
                    "Type": str(acc_type) if acc_type is not None else "",
                    "Percent": d_pct, "Amount": d_amt, "IsNet": False
                })
        
        # Net Account
        n_acc = norm_digits(row.get("Net_Acct_Code")).lstrip("0")
        n_rout = norm_digits(row.get("Net_Rout_Code")).lstrip("0")
        if n_acc or n_rout:
            n_type = row.get("Net_Type_Code")
            n_pct = 100.0 - total_dist_pct if total_dist_pct > 0 else (100.0 if not accs else 0.0)
            accs.append({
                "Routing": n_rout, "Account": n_acc,
                "Type": str(n_type) if n_type is not None else "",
                "Percent": max(0, n_pct), "Amount": 0.0, "IsNet": True
            })
        if accs: 
            # Translate Paycom ID to Uzio ID if mapping exists
            mapped_eid = pc_to_uz_id_map.get(eid, eid)
            p_map[mapped_eid] = accs

    FIELDS = ["Routing Number", "Account Number", "Account Type", "Amount", "Percent"]
    all_ids = set(u_map.keys()) | set(p_map.keys()) | set(uzio_emp_names.keys())
    
    for eid in sorted(all_ids):
        uas = u_map.get(eid, [])
        pas = p_map.get(eid, [])
        name = uzio_emp_names.get(eid, "")
        
        if not uas and pas:
            is_in_uzio = eid in uzio_emp_names
            for p in pas:
                for f in FIELDS:
                    rows.append({
                        "Employee ID": eid, "Employee Name": name, "Section": "Payment",
                        "Field": f, "Uzio Value": "", "Paycom Value": _get_field_val(p, f),
                        "Status": "Value missing in Uzio" if is_in_uzio else STATUS_MISSING_UZIO
                    })
            continue

        if uas and not pas:
            for u in uas:
                for f in FIELDS:
                    rows.append({
                        "Employee ID": eid, "Employee Name": name, "Section": "Payment",
                        "Field": f, "Uzio Value": _get_field_val(u, f), "Paycom Value": "",
                        "Status": STATUS_MISSING_PAYCOM
                    })
            continue

        # 3-Pass Matching
        p_pending = list(pas)
        u_pending = list(uas)
        matched = [] # (u, p)

        # Pass 1: Exact Routing + Account
        for u in list(u_pending):
            match = next((p for p in p_pending if u["Routing"] == p["Routing"] and u["Account"] == p["Account"]), None)
            if match:
                matched.append((u, match))
                u_pending.remove(u)
                p_pending.remove(match)

        # Pass 2: Routing + Type
        for u in list(u_pending):
            u_t = strip_type(u["Type"])
            match = next((p for p in p_pending if u["Routing"] == p["Routing"] and u_t == strip_type(p["Type"])), None)
            if match:
                matched.append((u, match))
                u_pending.remove(u)
                p_pending.remove(match)

        # Pass 3: Routing Only
        for u in list(u_pending):
            match = next((p for p in p_pending if u["Routing"] == p["Routing"]), None)
            if match:
                matched.append((u, match))
                u_pending.remove(u)
                p_pending.remove(match)

        # Record results
        for u, p in matched:
            for f in FIELDS:
                u_v = _get_field_val(u, f)
                p_v = _get_field_val(p, f)
                rows.append({
                    "Employee ID": eid, "Employee Name": name, "Section": "Payment",
                    "Field": f, "Uzio Value": u_v, "Paycom Value": p_v,
                    "Status": _compare_field(f, u_v, p_v, u, p)
                })
        
        for u in u_pending:
            for f in FIELDS:
                rows.append({
                    "Employee ID": eid, "Employee Name": name, "Section": "Payment",
                    "Field": f, "Uzio Value": _get_field_val(u, f), "Paycom Value": "Not Found",
                    "Status": "Value missing in Paycom"
                })
        for p in p_pending:
            for f in FIELDS:
                rows.append({
                    "Employee ID": eid, "Employee Name": name, "Section": "Payment",
                    "Field": f, "Uzio Value": "Not Found", "Paycom Value": _get_field_val(p, f),
                    "Status": "Value missing in Uzio"
                })

    return pd.DataFrame(rows)

def run_emergency_audit(df_uzio, df_paycom, uz_to_pc_id_map=None):
    rows = []
    
    # Pre-process identity map for reverse lookup
    pc_to_uz_id_map = {}
    if uz_to_pc_id_map:
        pc_to_uz_id_map = {pc: uz for uz, pc in uz_to_pc_id_map.items() if pc}

    u_id_col = "Job|Employee ID"
    p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
    
    # 1. Map Columns — discover every Emergency_N_* slot present in the Paycom
    # export (Advanced Report Writer dumps up to N emergency contacts per
    # employee as Emergency_1_*, Emergency_2_*, ...). Earlier versions only
    # read Emergency_1_*, silently dropping additional contacts.
    import re as _re
    p_slot_indices = sorted({
        int(m.group(1))
        for c in df_paycom.columns
        for m in [_re.match(r"Emergency_(\d+)_Contact$", str(c))]
        if m
    })
    p_slot_cols = []
    for n in p_slot_indices:
        p_slot_cols.append({
            "Name":  next((c for c in df_paycom.columns if c == f"Emergency_{n}_Contact"), None),
            "Rel":   next((c for c in df_paycom.columns if c == f"Emergency_{n}_Relationship"), None),
            "Phone": next((c for c in df_paycom.columns if c == f"Emergency_{n}_Phone"), None),
            "Lang":  next((c for c in df_paycom.columns if c == f"Emergency_{n}_Language"), None),
        })

    # 2. Process Uzio
    u_data = {}
    uzio_emp_names = {}
    for idx, row in df_uzio.iterrows():
        eid = norm_id(row.get(u_id_col))
        if not eid: continue
        
        name_emp = f"{norm_blank(row.get('Personal|First Name'))} {norm_blank(row.get('Personal|Last Name'))}".strip()
        uzio_emp_names[eid] = name_emp
        
        contact = {
            "Name": norm_str(row.get("Emergency Contact|Name")),
            "Relation": norm_relation(row.get("Emergency Contact|Relationship")),
            "Phone": norm_phone(row.get("Emergency Contact|Phone")),
            "RawPhone": norm_str(row.get("Emergency Contact|Phone")),
            "Language": ""
        }
        if contact["Name"] or contact["Phone"]:
            if eid not in u_data: u_data[eid] = []
            u_data[eid].append(contact)

    # 3. Process Paycom
    p_data = {}
    for idx, row in df_paycom.iterrows():
        eid = norm_id(row.get(p_id_col))
        if not eid: continue
        
        # Translate Paycom ID to Uzio ID if mapping exists
        mapped_eid = pc_to_uz_id_map.get(eid, eid)
        
        # Iterate every Emergency_N_* slot the file actually has, not just
        # Emergency_1_*. Each non-empty contact (Name OR Phone) is captured.
        for slot in p_slot_cols:
            name_col = slot["Name"]
            if not name_col:
                continue
            c_name = norm_str(row.get(name_col))
            c_phone_raw = row.get(slot["Phone"]) if slot["Phone"] else ""
            c_phone = norm_phone(c_phone_raw) if slot["Phone"] else ""
            if not c_name and not c_phone:
                continue
            contact = {
                "Name": c_name,
                "Relation": norm_relation(row.get(slot["Rel"])) if slot["Rel"] else "",
                "Phone": c_phone,
                "RawPhone": norm_str(c_phone_raw) if slot["Phone"] else "",
                "Language": norm_str(row.get(slot["Lang"])) if slot["Lang"] else "",
            }
            if mapped_eid not in p_data:
                p_data[mapped_eid] = []
            p_data[mapped_eid].append(contact)

    def compare_emergency(field, u_val, p_val):
        u_s = str(u_val).strip().lower()
        p_s = str(p_val).strip().lower()
        if u_s == p_s: return True
        if field == "Phone":
            u_p, p_p = norm_phone(u_val), norm_phone(p_val)
            if u_p == p_p: return True
            if u_p and p_p and (u_p in p_p or p_p in u_p): return True
        if field == "Relation":
            synonyms = [{"spouse", "husband", "wife"}, {"mother", "father", "parent"}]
            for group in synonyms:
                if u_s in group and p_s in group: return True
            if "child" in u_s and "child" in p_s: return True
        return False

    FIELDS = ["Name", "Relation", "Phone"]
    all_ids = set(u_data.keys()) | set(p_data.keys()) | set(uzio_emp_names.keys())
    
    for eid in sorted(all_ids):
        ucs = u_data.get(eid, [])
        pcs = p_data.get(eid, [])
        emp_name = uzio_emp_names.get(eid, "")

        if not ucs and pcs:
            for p in pcs:
                for f in FIELDS:
                    rows.append({
                        "Employee ID": eid, "Employee Name": emp_name, "Section": "Emergency",
                        "Field": f, "Uzio Value": "", "Paycom Value": p[f],
                        "Status": "Value missing in Uzio"
                    })
            continue
        if ucs and not pcs:
            for u in ucs:
                for f in FIELDS:
                    rows.append({
                        "Employee ID": eid, "Employee Name": emp_name, "Section": "Emergency",
                        "Field": f, "Uzio Value": u[f], "Paycom Value": "",
                        "Status": "Value missing in Paycom"
                    })
            continue

        # 2-Pass Match
        u_pending = list(ucs)
        p_pending = list(pcs)
        matched = []

        # Pass 1: Name Match
        for u in list(u_pending):
            match = next((p for p in p_pending if u["Name"].lower() == p["Name"].lower()), None)
            if match:
                matched.append((u, match))
                u_pending.remove(u)
                p_pending.remove(match)

        # Pass 2: Phone Match
        for u in list(u_pending):
            if not u["Phone"]: continue
            match = next((p for p in p_pending if u["Phone"] == p["Phone"]), None)
            if match:
                matched.append((u, match))
                u_pending.remove(u)
                p_pending.remove(match)

        # Record
        for u, p in matched:
            for f in FIELDS:
                u_v, p_v = u[f], p[f]
                rows.append({
                    "Employee ID": eid, "Employee Name": emp_name, "Section": "Emergency",
                    "Field": f, "Uzio Value": u["RawPhone"] if f=="Phone" else u_v,
                    "Paycom Value": p["RawPhone"] if f=="Phone" else p_v,
                    "Status": STATUS_MATCH if compare_emergency(f, u_v, p_v) else STATUS_MISMATCH
                })
            if p["Language"]:
                rows.append({
                    "Employee ID": eid, "Employee Name": emp_name, "Section": "Emergency",
                    "Field": "Language", "Uzio Value": "N/A", "Paycom Value": p["Language"],
                    "Status": "Info Only"
                })

        for u in u_pending:
            for f in FIELDS:
                rows.append({
                    "Employee ID": eid, "Employee Name": emp_name, "Section": "Emergency",
                    "Field": f, "Uzio Value": u[f], "Paycom Value": "", "Status": "Value missing in Paycom"
                })
        for p in p_pending:
            for f in FIELDS:
                rows.append({
                    "Employee ID": eid, "Employee Name": emp_name, "Section": "Emergency",
                    "Field": f, "Uzio Value": "", "Paycom Value": p[f], "Status": "Value missing in Uzio"
                })

    return pd.DataFrame(rows)

# --- UI Functions ---
def render_ui():
    st.title(APP_TITLE)
    st.markdown("""
    This tool performs a consolidated audit of Census, Payment, and Emergency contact data 
    using the **Uzio Master Custom Report** and a **Paycom Census Export**.
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        u_file = st.file_uploader("Upload Uzio Master Report (CSV)", type=["csv"], key="comb_u")
    with col2:
        p_file = st.file_uploader("Upload Paycom Census Export (Excel/CSV)", type=["xlsx", "csv"], key="comb_p")

    if st.button("Run Consolidated Audit", type="primary"):
        if not u_file or not p_file:
            st.error("Please upload both files.")
            return
            
        try:
            with st.spinner("Processing files..."):
                # Load Uzio
                df_uzio = read_uzio_master(u_file)

                # Client name comes straight from the Uzio report's
                # Company Information > Company Name column.
                client_name = company_name_from_uzio_master(df_uzio) or "Client"

                # Load Paycom
                if p_file.name.endswith(".csv"):
                    df_paycom = pd.read_csv(p_file, dtype=str)
                else:
                    df_paycom = pd.read_excel(p_file, dtype=str)
                
                # --- 1. Prepare ID Maps & Identity Match ---
                u_id_col = "Job|Employee ID"
                p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
                uzio_ssn_col = 'Personal|SSN'
                paycom_ssn_col = next((c for c in df_paycom.columns if "SS_Number" in c or "SSN" in c), "SS_Number")
                
                # Normalize IDs in dataframes first
                df_uzio[u_id_col] = df_uzio[u_id_col].apply(norm_id)
                df_paycom[p_id_col] = df_paycom[p_id_col].apply(norm_id)
                
                # Get the identity map (Uzio_ID -> Paycom_ID)
                uz_to_pc_id_map = get_identity_match_map(
                    df_uzio, df_paycom, 
                    uzio_id_col=u_id_col, 
                    vendor_id_col=p_id_col,
                    uzio_ssn_col=uzio_ssn_col,
                    vendor_ssn_col=paycom_ssn_col
                )

                # --- 2. Run Data Quality Checks ---
                df_uz_dupes = detect_duplicate_ssns(df_uzio, u_id_col, uzio_ssn_col)
                df_pc_dupes = detect_duplicate_ssns(df_paycom, p_id_col, paycom_ssn_col)
                
                dupe_rows = []
                for ssn, ids in df_uz_dupes.items():
                    dupe_rows.append({"Source": "Uzio", "SSN": ssn, "IDs": ", ".join(ids), "Issue": "Duplicate SSN"})
                for ssn, ids in df_pc_dupes.items():
                    dupe_rows.append({"Source": "Paycom", "SSN": ssn, "IDs": ", ".join(ids), "Issue": "Duplicate SSN"})
                df_dupe_ssn_check = pd.DataFrame(dupe_rows)

                # --- 3. Run Audits ---
                res_census = run_census_audit(df_uzio, df_paycom, uz_to_pc_id_map=uz_to_pc_id_map)
                res_payment = run_payment_audit(df_uzio, df_paycom, uz_to_pc_id_map=uz_to_pc_id_map)
                res_emergency = run_emergency_audit(df_uzio, df_paycom, uz_to_pc_id_map=uz_to_pc_id_map)
                
                # --- 4. Run Anomaly Reports ---
                df_salaried_drivers = get_salaried_driver_exceptions(df_uzio, df_paycom)
                df_flsa_issues = get_flsa_compliance_issues(df_uzio)
                df_active_missing = get_active_missing_in_uzio(df_uzio, df_paycom)
                df_terminated_missing = get_terminated_missing_in_uzio(df_uzio, df_paycom)
                df_dq_issues = get_data_quality_issues(df_paycom)
                df_high_rates = get_high_rate_anomalies(df_paycom)
                
                # --- Generate Summary Metrics ---
                p_id_col = next((c for c in df_paycom.columns if "Employee_Code" in c), "Employee_Code")
                uzio_ids = set(df_uzio["Job|Employee ID"].map(norm_id))
                pay_ids = set(df_paycom[p_id_col].map(norm_id))
                
                summary_data = [
                    {"Metric": "Employees in Uzio Master", "Value": len(uzio_ids)},
                    {"Metric": "Employees in Paycom Export", "Value": len(pay_ids)},
                    {"Metric": "Employees in Both", "Value": len(uzio_ids & pay_ids)},
                    {"Metric": "---", "Value": ""},
                    {"Metric": "Census Matches", "Value": len(res_census[res_census["Status"] == STATUS_MATCH])},
                    {"Metric": "Census Mismatches", "Value": len(res_census[res_census["Status"] == STATUS_MISMATCH])},
                    {"Metric": "Payment Matches", "Value": len(res_payment[res_payment["Status"] == STATUS_MATCH])},
                    {"Metric": "Payment Mismatches", "Value": len(res_payment[res_payment["Status"] == STATUS_MISMATCH])},
                    {"Metric": "Emergency Matches", "Value": len(res_emergency[res_emergency["Status"] == STATUS_MATCH])},
                    {"Metric": "Emergency Mismatches", "Value": len(res_emergency[res_emergency["Status"] == STATUS_MISMATCH])},
                    {"Metric": "---", "Value": ""},
                    {"Metric": "Salaried Driver Exceptions", "Value": len(df_salaried_drivers)},
                    {"Metric": "FLSA Compliance Issues", "Value": len(df_flsa_issues)},
                    {"Metric": "Active Employees Missing in Uzio", "Value": len(df_active_missing)},
                    {"Metric": "Terminated Employees Missing in Uzio", "Value": len(df_terminated_missing)},
                    {"Metric": "Data Quality Issues (00/00/0000)", "Value": len(df_dq_issues)},
                    {"Metric": "High Hourly Rate Anomalies (>$100)", "Value": len(df_high_rates)},
                    {"Metric": "Duplicate SSN Warnings", "Value": len(df_dupe_ssn_check)},
                ]
                df_summary = pd.DataFrame(summary_data)

                # Download
                out = io.BytesIO()
                with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
                    df_summary.to_excel(writer, sheet_name="Summary", index=False)
                    if not df_dupe_ssn_check.empty:
                        df_dupe_ssn_check.to_excel(writer, sheet_name="Duplicate_SSN_Check", index=False)
                    
                    # Consolidate Anomaly Reports into the Summary Sheet
                    start_row = len(df_summary) + 3
                    anomaly_groups = [
                        ("Salaried Driver Exceptions", df_salaried_drivers),
                        ("FLSA Compliance Issues", df_flsa_issues),
                        ("Active Employees Missing in Uzio", df_active_missing),
                        ("Terminated Employees Missing in Uzio", df_terminated_missing),
                        ("Data Quality Issues", df_dq_issues),
                        ("High Hourly Rate Anomalies", df_high_rates)
                    ]
                    
                    curr_row = start_row
                    workbook = writer.book
                    worksheet = writer.sheets['Summary']
                    header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
                    
                    for title, df_ano in anomaly_groups:
                        worksheet.write(curr_row, 0, title, header_fmt)
                        curr_row += 1
                        if not df_ano.empty:
                            df_ano.to_excel(writer, sheet_name="Summary", startrow=curr_row, index=False)
                            curr_row += len(df_ano) + 2
                        else:
                            worksheet.write(curr_row, 0, "No issues found.")
                            curr_row += 2

                    res_census.to_excel(writer, sheet_name="Census_Audit", index=False)
                    res_payment.to_excel(writer, sheet_name="Payment_Audit", index=False)
                    res_emergency.to_excel(writer, sheet_name="Emergency_Audit", index=False)
                    
                timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
                filename = f"{client_name}_Consolidated_Audit_Report_{timestamp}.xlsx"
                
                st.download_button(
                    label=f"Download {client_name} Consolidated Audit Report",
                    data=out.getvalue(),
                    file_name=filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)

if __name__ == "__main__":
    render_ui()
