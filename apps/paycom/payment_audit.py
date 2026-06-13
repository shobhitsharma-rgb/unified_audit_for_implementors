import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime, date

# =========================================================
# Paycom vs Uzio Payment Audit Tool
# =========================================================

APP_TITLE = "Paycom vs Uzio – Payment Audit Tool"

# --- Constants for Status (8 statuses, matching census_audit_app.py) ---
STATUS_MATCH = "Data Match"
STATUS_MISMATCH = "Data Mismatch"
STATUS_VAL_MISSING_UZIO = "Value missing in Uzio (Paycom has value)"
STATUS_VAL_MISSING_PAYCOM = "Value missing in Paycom (Uzio has value)"
STATUS_MISSING_UZIO = "Employee ID Not Found in Uzio"
STATUS_MISSING_PAYCOM = "Employee ID Not Found in Paycom"
STATUS_COL_MISSING_PAYCOM = "Column Missing in Paycom Sheet"
STATUS_COL_MISSING_UZIO = "Column Missing in Uzio Sheet"
# A difference that is ONLY a missing leading zero (e.g. Paycom 028248003 vs
# Uzio 28248003) is the Excel/Uzio numeric-coercion artifact, NOT a real
# data-entry mismatch. It is surfaced as its own status so reviewers can tell
# the benign export bug apart from a genuinely wrong account number.
STATUS_LEADING_ZERO = "Leading Zero Dropped (likely export artifact)"

def _is_leading_zero_drop(a, b):
    """True when a and b are identical apart from one or more leading zeros
    (and at least one digit remains after stripping). Both must be non-blank."""
    a, b = str(a).strip(), str(b).strip()
    if not a or not b or a == b:
        return False
    sa, sb = a.lstrip("0"), b.lstrip("0")
    return sa == sb and sa != ""

def norm_str(x):
    if x is None:
        return ""
    return str(x).strip()

def norm_digits(x):
    """Keep only digits, remove spaces/dashes. Preserves leading zeros."""
    if x is None:
        return ""
    if isinstance(x, (float, int)):
        if pd.isna(x):
            return ""
        # Handle float like 123.0 -> '123'
        return str(int(x))
    # For strings, just remove non-digits. This preserves leading zeros like '00123'.
    return re.sub(r"\D", "", str(x))

def norm_money(x):
    """Parse money/float safely."""
    if x is None:
        return 0.0
    if isinstance(x, (float, int)):
        return 0.0 if pd.isna(x) else float(x)
    s = str(x).replace(",", "").replace("$", "").strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except:
        return 0.0

def norm_id(x):
    """Normalize Employee ID: strip float .0 and pad to 4 digits if numeric."""
    if x is None: return ""
    s = str(x).strip()
    if s.endswith(".0"): 
        s = s[:-2]
    # Pad to 4 digits if it's a number and < 4 length (e.g. '1' -> '0001')
    if s.isdigit() and len(s) < 4:
        return s.zfill(4)
    return s

# Paycom uses numeric type codes; map them to human-readable names
_TYPE_CODE_MAP = {
    "22": "checking",
    "32": "savings",
    "1": "checking",   # Net_Type_Code sometimes uses 1 for checking
    "2": "checking",   # Sometimes Paycom uses 2 or 2.0 for checking
}

def strip_type(t):
    """Normalize account type string for comparison.
    Handles Paycom numeric codes: 22=Checking, 32=Savings."""
    if not t: return ""
    s = str(t).strip()
    # Remove trailing ".0" from float-read values like "22.0"
    if s.endswith(".0"):
        s = s[:-2]
    # Check if it's a known Paycom type code
    if s in _TYPE_CODE_MAP:
        return _TYPE_CODE_MAP[s]
    return s.lower().replace("account", "").replace("code: ", "").strip()

# ---------- UI ----------
def _inject_payment_styles():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;700;800&family=Inter:wght@400;600&display=swap');
        .pay-hero { background: linear-gradient(135deg,#0f172a 0%,#1e3a8a 100%); border-radius:16px; padding:26px 30px; margin-bottom:14px; }
        .pay-hero-title { color:#ffffff !important; font-family:'Manrope',sans-serif; font-size:28px; font-weight:800; letter-spacing:-0.02em; }
        .pay-hero-sub { color:#c7d2fe !important; font-family:'Inter',sans-serif; font-size:15px; margin-top:6px; line-height:1.5; }
        .step-pill { display:inline-block; background:#eef2ff; color:#3730a3; font-weight:700; border-radius:999px; padding:2px 10px; font-size:12px; }
        .metric-card { background:#ffffff; border:1px solid #e6e8ee; border-radius:14px; padding:16px 14px; text-align:center; box-shadow:0 1px 3px rgba(16,24,40,.06); }
        .metric-icon { font-size:20px; }
        .metric-num { font-family:'Manrope',sans-serif; font-size:30px; font-weight:800; color:#0f172a; line-height:1.15; }
        .metric-label { font-family:'Inter',sans-serif; font-size:12.5px; color:#475467; margin-top:2px; }
        .action-card { background:#fff7ed; border:1px solid #fed7aa; border-left:6px solid #ea580c; border-radius:12px; padding:18px 20px; margin:4px 0 10px; }
        .action-title { font-family:'Manrope',sans-serif; font-size:18px; font-weight:800; color:#9a3412 !important; margin-bottom:8px; }
        .action-body { font-family:'Inter',sans-serif; font-size:14.5px; color:#7c2d12 !important; line-height:1.65; }
        .action-body code { background:#fde68a; padding:1px 6px; border-radius:5px; color:#7c2d12; font-weight:700; }
        .ok-card { background:#ecfdf5; border:1px solid #a7f3d0; border-left:6px solid #059669; border-radius:12px; padding:18px 20px; margin:4px 0 10px; }
        .ok-title { font-family:'Manrope',sans-serif; font-size:18px; font-weight:800; color:#065f46 !important; }
        .ok-body { font-family:'Inter',sans-serif; font-size:14.5px; color:#065f46 !important; margin-top:4px; }
        .stButton button, .stDownloadButton button { border-radius:10px !important; font-weight:700 !important; }
        .stButton button[kind="primary"], .stDownloadButton button[kind="primary"] { background:linear-gradient(135deg,#1e3a8a,#3b82f6) !important; border:none !important; }
        .stButton button[kind="primary"] p, .stDownloadButton button[kind="primary"] p { color:#ffffff !important; }
        </style>
    """, unsafe_allow_html=True)


def render_ui():
    _inject_payment_styles()
    st.markdown(
        '<div class="pay-hero">'
        '<div class="pay-hero-title">&#128179; Paycom &#8594; Uzio &#183; Bank Account Check</div>'
        '<div class="pay-hero-sub">Upload the two files and we will compare every bank account in Uzio against '
        'Paycom, then tell you in plain English exactly what is wrong and how to fix it.</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<span class="step-pill">Step 1</span> &nbsp;Upload your Uzio file &nbsp;&nbsp;&nbsp;'
        '<span class="step-pill">Step 2</span> &nbsp;Upload your Paycom file &nbsp;&nbsp;&nbsp;'
        '<span class="step-pill">Step 3</span> &nbsp;Read the results &amp; download the fix list',
        unsafe_allow_html=True,
    )
    st.markdown("<br>", unsafe_allow_html=True)

    client_name = st.text_input("Client name", value="", placeholder="e.g. Chief Delivery",
                                help="Only used to name the downloaded report file.")

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**\U0001F4E4 Uzio file**")
        st.caption("The HR Report you exported from Uzio (.xlsx or .csv).")
        uzio_file = st.file_uploader("Uzio HR Report", type=["xlsx", "xlsm", "csv"],
                                     label_visibility="collapsed", key="pay_uzio")
    with c2:
        st.markdown("**\U0001F4E5 Paycom file**")
        st.caption("The Advanced Report Writer export from Paycom (.csv or .xlsx).")
        paycom_file = st.file_uploader("Paycom export", type=["xlsx", "csv"],
                                       label_visibility="collapsed", key="pay_paycom")

    run_btn = st.button("\U0001F50D  Check the bank accounts", type="primary",
                        disabled=(not uzio_file or not paycom_file), use_container_width=True)

    if not run_btn:
        return

    try:
        with st.spinner("Comparing every bank account..."):
            report_bytes, leading_zero_issues, stats = run_audit(uzio_file, paycom_file)
    except Exception as e:
        st.error(f"Sorry - something went wrong while reading the files: {e}")
        return

    client_label = client_name.strip() or "Client"
    st.markdown("---")
    st.markdown(f"## ✅ Done - here is what we found for **{client_label}**")

    cards = [
        ("\U0001F465", stats["employees_checked"], "Employees checked", "#3b82f6"),
        ("✅", stats["matched"], "Account details that match", "#16a34a"),
        ("\U0001F527", stats["leading_zero"], "Need re-import (lost a 0)", "#ea580c"),
        ("⚠️", stats["real_mismatch"], "Real mismatches to review", "#dc2626"),
    ]
    cols = st.columns(len(cards))
    for col, (icon, val, label, color) in zip(cols, cards):
        col.markdown(
            f'<div class="metric-card" style="border-top:4px solid {color}">'
            f'<div class="metric-icon">{icon}</div>'
            f'<div class="metric-num">{val}</div>'
            f'<div class="metric-label">{label}</div></div>',
            unsafe_allow_html=True,
        )
    st.markdown("<br>", unsafe_allow_html=True)

    if not leading_zero_issues.empty:
        n = len(leading_zero_issues)
        st.markdown(
            f'<div class="action-card">'
            f'<div class="action-title">\U0001F527 {n} bank account number(s) need a quick re-import</div>'
            f'<div class="action-body">'
            f'<b>What happened:</b> when this data was loaded into Uzio, {n} account number(s) lost their starting '
            f'zero &mdash; for example <code>028248003</code> turned into <code>28248003</code>.<br>'
            f'<b>Is it serious?</b> No. It is an Excel/Uzio number-formatting glitch, not a typing mistake. '
            f'The value in <b>Paycom is the correct one</b>.<br>'
            f'<b>✅ What to do:</b> re-import these account numbers from the Paycom file. The exact correct values '
            f'are listed below and in the <b>"Accounts To Re-Import"</b> tab of the downloaded report.'
            f'</div></div>',
            unsafe_allow_html=True,
        )
        st.caption('"Correct value (from Paycom)" is what the account should be. "Wrong value now in Uzio" is what is stored today.')
        with st.container(height=360, border=True):
            st.dataframe(leading_zero_issues, hide_index=True, use_container_width=True)

    if stats["real_mismatch"] > 0:
        st.markdown(f"#### ⚠️ {stats['real_mismatch']} value(s) genuinely differ - please review")
        st.caption("These are not leading-zero glitches: the Uzio and Paycom values are actually different, so someone "
                   "should confirm which one is right. Open the \"All Comparisons\" tab in the report and filter the "
                   "Status column to \"Data Mismatch\".")

    if leading_zero_issues.empty and stats["real_mismatch"] == 0:
        st.markdown(
            '<div class="ok-card"><div class="ok-title">\U0001F389 Everything matches!</div>'
            '<div class="ok-body">Every bank account in Uzio agrees with Paycom. There is nothing to fix.</div></div>',
            unsafe_allow_html=True,
        )

    if stats["missing_uzio"] or stats["missing_paycom"]:
        st.caption(f"ℹ️ Heads-up: {stats['missing_uzio']} value(s) are in Paycom but missing from Uzio, and "
                   f"{stats['missing_paycom']} are in Uzio but missing from Paycom. Details are in the report.")

    st.markdown("<br>", unsafe_allow_html=True)
    timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
    out_filename = f"{client_label}_Uzio_Paycom_Payment_Audit_Report_{timestamp}.xlsx"
    st.download_button(
        label="⬇️  Download the full report (Excel)",
        data=report_bytes,
        file_name=out_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )
    st.caption('The report opens with a "Start Here" tab that explains every other tab in plain English.')

# ---------- Core Audit Logic ----------
def _read_payment_file(file, header=0):
    """Read an .xlsx/.xlsm/.csv payment export. Header is the row index of column names."""
    name = (getattr(file, 'name', '') or '').lower()
    file.seek(0)
    if name.endswith('.csv'):
        try:
            return pd.read_csv(file, header=header, dtype=str)
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, header=header, dtype=str, encoding='latin1')
    return pd.read_excel(file, header=header, dtype=str)

def run_audit(uzio_file, paycom_file):
    # 1. Load Data
    # Uzio Raw: Skip first row (header=1)
    df_uzio = _read_payment_file(uzio_file, header=1)

    # Paycom: Payment Export (.xlsx or .csv)
    try:
        df_paycom = _read_payment_file(paycom_file, header=0)
    except Exception as e:
        raise ValueError(f"Error reading Paycom file: {e}")

    # 2. Process Uzio Data
    df_uzio.columns = [str(c).strip() for c in df_uzio.columns]
    
    # Mapping for Raw Uzio Export
    # Raw Columns: 'Company Name', 'Full Name', 'Employee ID', 'Payment Method', 
    # 'Paycheck Distribution', 'Routing Number', 'Account Type', 'Account Number', 
    # 'Paycheck Percentage', 'Paycheck Amount', 'Priority'
    
    # Map to internal keys expected by logic:
    # EmpID -> 'Employee ID'
    # Routing -> 'Routing Number'
    # Account -> 'Account Number'
    # Type -> 'Account Type'
    # Percent -> 'Paycheck Percentage'
    # Amount -> 'Paycheck Amount'
    # Name -> 'Full Name'
    
    col_map = {
        "EmpID": "Employee ID",
        "Routing": "Routing Number",
        "Account": "Account Number",
        "Type": "Account Type",
        "Percent": "Paycheck Percentage",
        "Amount": "Paycheck Amount",
        "Name": "Full Name"
    }

    uzio_map = {}
    uzio_emp_names = {}
    
    for idx, row in df_uzio.iterrows():
        # UZIO ID: Keep AS IS (do not pad/normalize) to avoid colliding "001" and "0001"
        emp_id = norm_str(row.get(col_map["EmpID"]))
        if not emp_id: continue
        
        name_str = norm_str(row.get(col_map["Name"]))
        if emp_id not in uzio_emp_names:
            uzio_emp_names[emp_id] = name_str
        
        acc = {
            "Routing": norm_digits(row.get(col_map["Routing"])),
            "Account": norm_digits(row.get(col_map["Account"])),
            "Type": norm_str(row.get(col_map["Type"])),
            "Percent": norm_money(row.get(col_map["Percent"])),
            "Amount": norm_money(row.get(col_map["Amount"])),
            "Name": name_str
        }
        
        if acc["Routing"] or acc["Account"]:
            if emp_id not in uzio_map:
                uzio_map[emp_id] = []
            
            # Deduplicate: Only add if this exact account isn't already listed for this employee
            if acc not in uzio_map[emp_id]:
                uzio_map[emp_id].append(acc)

    # 3. Process Paycom Data (Wide Format -> Unpivot)
    paycom_accounts = []
    paycom_status_map = {}  # resolved EmpID -> Paycom Employee_Status (Active/Terminated/...)
    df_paycom.columns = [str(c).strip() for c in df_paycom.columns]

    pc_empid_col = next((c for c in df_paycom.columns if "Employee_Code" in c or "Emp Code" in c), "Employee_Code")
    # Employee status column (exact match only — avoid Bonus_Status / Net_Status / DOL_Status etc.)
    pc_status_col = next((c for c in df_paycom.columns if c.strip().lower() in ("employee_status", "employee status")), "")
    
    # Identify Name Columns in Paycom for validation
    # Pria Paycom Cenus.xlsx has: Legal_Firstname, Legal_Lastname, Legal_Middle_Name
    # HR Report (Uzio) has: Full Name
    
    pc_first_col = next((c for c in df_paycom.columns if "Firstname" in c or "First Name" in c), "")
    pc_last_col = next((c for c in df_paycom.columns if "Lastname" in c or "Last Name" in c), "")
    
    # Helper to resolve Paycom ID with Name Validation
    def resolve_paycom_id(raw_id, raw_name_parts, uzio_data):
        if pd.isna(raw_id): return None
        s_id = str(raw_id).strip()
        if s_id.endswith(".0"): s_id = s_id[:-2]
        
        # 1. Direct Candidates (Exact + Padded)
        direct_to_try = [s_id]
        for w in [3, 4, 5]:
            padded = s_id.zfill(w)
            if padded not in direct_to_try:
                direct_to_try.append(padded)
        
        valid_direct = [c for c in direct_to_try if c in uzio_data]
        
        using_fixed = False
        if valid_direct:
            candidates = valid_direct
        else:
            # 2. Transposed Candidates (Only if no direct matches)
            # Often A0BZ in Paycom is AOBZ in Uzio
            transposed_to_try = []
            for c in direct_to_try:
                c_o = c.replace('0', 'O')
                if c_o != c and c_o not in transposed_to_try: transposed_to_try.append(c_o)
                c_zero = c.replace('O', '0')
                if c_zero != c and c_zero not in transposed_to_try: transposed_to_try.append(c_zero)
            
            valid_fixed = [c for c in transposed_to_try if c in uzio_data]
            candidates = valid_fixed
            using_fixed = True

        if not candidates:
            return s_id.zfill(4) # Fallback if absolutely no match found
            
        pc_lowers = [str(p).lower().strip() for p in raw_name_parts if p and not pd.isna(p)]
        
        # If we have only one direct match, no reason to doubt it
        if len(candidates) == 1 and not using_fixed:
            return candidates[0]
            
        # COLLISION OR FIX VERIFICATION:
        # Use Name Matching to decide between candidates or verify a "fixed" ID
        best_match = None
        best_score = -1
        
        for cand in candidates:
            u_name = uzio_data.get(cand, "").lower()
            score = 0
            if pc_lowers:
                for part in pc_lowers:
                    if part in u_name:
                        score += 5 # High score for name part match
                    elif len(part) > 3:
                        # Partial match for longer names
                        if u_name.startswith(part) or u_name.endswith(part):
                            score += 2
            
            # Favor exact length matches if scores are tied
            if len(cand) == len(s_id):
                score += 1

            if score > best_score:
                best_score = score
                best_match = cand
            elif score == best_score:
                # If tied, pick first one or stick with best_match
                pass

        # SAFETY CHECK for "fixed" (transposed) IDs:
        # If we are using a "fixed" ID, we require at least one name overlap to accept it,
        # otherwise it's safer to report it as the original ID (Missing in Uzio).
        if using_fixed and pc_lowers and best_score <= 1: # 1 is tiebreaker for length, 5+ is name match
             return s_id.zfill(4)
                
        return best_match if best_match else candidates[0]

    uzio_keys = set(uzio_map.keys())

    for idx, row in df_paycom.iterrows():
        raw_id = row.get(pc_empid_col)
        
        # Get Name parts for validation
        name_parts = []
        if pc_first_col: name_parts.append(row.get(pc_first_col))
        if pc_last_col: name_parts.append(row.get(pc_last_col))
        
        # Smart Resolve with Name
        emp_id = resolve_paycom_id(raw_id, name_parts, uzio_emp_names)

        if not emp_id: continue

        # Capture Paycom employee status (first non-blank wins per employee)
        if pc_status_col:
            status_val = norm_str(row.get(pc_status_col))
            if status_val and emp_id not in paycom_status_map:
                paycom_status_map[emp_id] = status_val

        # --- Extract Distributions (1 to 8) FIRST, so we can sum percents ---
        dist_entries = []
        total_dist_pct = 0.0
        total_dist_amt = 0.0

        for i in range(1, 9):
            prefix = f"Dist_{i}_"
            d_acc = norm_digits(row.get(f"{prefix}Acct_Code"))
            d_rout = norm_digits(row.get(f"{prefix}Rout_Code"))
            
            # Extract Amount/Percent always (even if no account, e.g. Check/Cash)
            raw_amt = row.get(f"{prefix}Amount")
            d_amt = norm_money(raw_amt)
            d_pct = 0.0

            # Check for a dedicated Percent column first
            pct_col = f"{prefix}Percent"
            if pct_col in df_paycom.columns:
                d_pct = norm_money(row.get(pct_col))

            # Detect percentage in Amount field:
            # logic: if existing d_pct is 0, check raw_amt for '%' symbol
            if d_pct == 0.0:
                raw_str = str(raw_amt).strip() if raw_amt is not None else ""
                
                # Case A: String contains explicit "%" (e.g. "25%" or "99%")
                if "%" in raw_str:
                    try:
                        d_pct = float(raw_str.replace("%", "").replace(",", "").strip())
                    except:
                        d_pct = 0.0
                    d_amt = 0.0
                
                # Case B: String contains explicit "$" -> DEFINITELY AMOUNT. Do nothing.
                elif "$" in raw_str:
                    pass

                # Case C: No symbols (e.g. 0.5 float)
                # If it looks like a decimal percentage (0.01 < x <= 1.0), assume Percent.
                # User says: "Amount always has $". So lack of $ implies non-Amount?
                # UPDATED: Exclude 0.01 (and below) to prevents $0.01 (Penny) from becoming 1%.
                elif d_amt != 0.0 and 0.01 < abs(d_amt) <= 1.0:
                    d_pct = round(d_amt * 100, 4)
                    d_amt = 0.0

            # Ignore completely blank distributions to prevent false mismatches
            if d_pct == 0.0 and d_amt == 0.0 and not d_acc and not d_rout:
                continue

            total_dist_pct += d_pct
            total_dist_amt += d_amt

            if d_acc or d_rout:
                d_type = row.get(f"{prefix}Type_Code")
                
                dist_entries.append({
                    "EmpID": emp_id,
                    "Routing": d_rout,
                    "Account": d_acc,
                    "Type": str(d_type) if d_type is not None else "",
                    "Percent": d_pct,
                    "Amount": d_amt,
                    "IsNet": False
                })

        # Filter valid dist entries - don't add empty accounts
        valid_dists = [d for d in dist_entries if d["Amount"] > 0 or d["Percent"] > 0 or d["Account"] or d["Routing"]]
        paycom_accounts.extend(valid_dists)

        # --- Extract NET Pay Account (remainder after distributions) ---
        net_acc = norm_digits(row.get("Net_Acct_Code"))
        net_rout = norm_digits(row.get("Net_Rout_Code"))
        
        if net_acc or net_rout:
             p_type = row.get("Net_Type_Code")
             
             # Calculate Net Percent:
             # Case 1: Partial Percentage Dists -> Net is remainder (100 - total)
             if total_dist_pct > 0:
                 net_pct = round(100.0 - total_dist_pct, 4)
             # Case 2: Flat Dollar Dists (no %) -> Net is just "Remainder" (usually 0% or handled as amount)
             elif total_dist_amt > 0:
                 net_pct = 0.0
             # Case 3: No distributions -> 100% Net Pay
             else:
                 net_pct = 100.0

             # Only add Net Account if there's actually remainder pay going somewhere
             if net_pct > 0 or net_acc or net_rout:
                 paycom_accounts.append({
                     "EmpID": emp_id,
                     "Routing": net_rout,
                     "Account": net_acc,
                     "Type": str(p_type) if p_type is not None else "",
                     "Percent": net_pct,
                     "Amount": 0.0,
                     "IsNet": True
                 })

    # Group Paycom by EmpID
    paycom_map = {}
    for item in paycom_accounts:
        eid = item["EmpID"]
        if eid not in paycom_map:
            paycom_map[eid] = []
        
        # Deduplicate: Only add if unique
        if item not in paycom_map[eid]:
            paycom_map[eid].append(item)

    # 4. Comparison Logic — Long Format (one row per field per account)
    # Fields to compare: Routing Number, Account Number, Account Type, Amount, Percent
    FIELDS = ["Routing Number", "Account Number", "Account Type", "Amount", "Percent"]

    rows = []
    all_emps = set(uzio_emp_names.keys()) | set(paycom_map.keys())

    for emp_id in sorted(all_emps):
        u_accs = uzio_map.get(emp_id, [])
        p_accs = paycom_map.get(emp_id, [])
        
        emp_name = u_accs[0]["Name"] if u_accs else uzio_emp_names.get(emp_id, "")

        # --- Case: Missing in Uzio ---
        if not u_accs and p_accs:
            is_in_uzio = emp_id in uzio_emp_names
            for p in p_accs:
                for field in FIELDS:
                    p_val = _get_field_val(p, field)
                    rows.append({
                        "Employee ID": emp_id,
                        "Employee Name": emp_name,
                        "Field": field,
                        "UZIO_Value": "",
                        "Paycom_Value": p_val,
                        "Paycom_SourceOfTruth_Status": STATUS_VAL_MISSING_UZIO if is_in_uzio else STATUS_MISSING_UZIO
                    })
            continue

        # --- Case: Missing in Paycom ---
        if u_accs and not p_accs:
            for u in u_accs:
                for field in FIELDS:
                    u_val = _get_field_val(u, field)
                    rows.append({
                        "Employee ID": emp_id,
                        "Employee Name": u["Name"],
                        "Paycom_Account_Class": "Not Found",
                        "Field": field,
                        "UZIO_Value": u_val,
                        "Paycom_Value": "",
                        "Paycom_SourceOfTruth_Status": STATUS_MISSING_PAYCOM
                    })
            continue

        # --- Case: Both exist — Match accounts (two-pass strategy) ---
        p_remaining = list(p_accs)
        u_unmatched = []

        # Pass 1: Exact match on Routing + Account
        # UPDATED: Use "Best Fit" strategy to handle multiple identical accounts
        # (e.g. A0BZ has 2 Checking accts with same numbers, formatted as 1% and 99%)
        for u in u_accs:
            match = None
            # Find ALL candidates that match Routing+Account
            candidates = []
            for p in p_remaining:
                if u["Routing"] == p["Routing"] and u["Account"] == p["Account"]:
                    candidates.append(p)
            
            if not candidates:
                u_unmatched.append(u)
                continue
            
            # Select Best Fit among candidates
            if len(candidates) == 1:
                match = candidates[0]
            else:
                # UPDATED: Scoring System to handle ties (e.g. 50% Checking vs 50% Savings)
                best_c = candidates[0]
                max_score = -1
                
                u_pct = u.get("Percent", 0.0)
                u_amt = u.get("Amount", 0.0)
                u_type = strip_type(u["Type"]) # Normalize Uzio type

                for c in candidates:
                    score = 0
                    c_pct = c.get("Percent", 0.0)
                    c_amt = c.get("Amount", 0.0)
                    c_type = strip_type(c["Type"]) # Normalize Paycom type

                    # Score Criteria
                    # 1. Percent Match (+10)
                    if u_pct > 0 and abs(u_pct - c_pct) < 0.01:
                        score += 10
                    
                    # 2. Amount Match (+10)
                    if u_amt > 0 and abs(u_amt - c_amt) < 0.01:
                        score += 10

                    # 3. Type Match (+5) - Key tiebreaker for duplicate accounts
                    if u_type and c_type and u_type == c_type:
                        score += 5
                    
                    if score > max_score:
                        max_score = score
                        best_c = c
                
                match = best_c

            if match:
                p_remaining.remove(match)
                for field in FIELDS:
                    u_val = _get_field_val(u, field)
                    p_val = _get_field_val(match, field)
                    status = _compare_field(field, u_val, p_val, u, match)
                    acc_class = "Net Account" if match.get("IsNet") else "Distribution Account"
                    rows.append({
                        "Employee ID": emp_id,
                        "Employee Name": u["Name"],
                        "Paycom_Account_Class": acc_class,
                        "Field": field,
                        "UZIO_Value": u_val,
                        "Paycom_Value": p_val,
                        "Paycom_SourceOfTruth_Status": status
                    })
            else:
                # Should not happen given logic above, but safety fallback
                u_unmatched.append(u)

        # Pass 2: Fallback match on Routing + Account Type
        # (handles Paycom exports where account numbers lost precision)
        still_unmatched = []
        for u in u_unmatched:
            match = None
            u_type = strip_type(u["Type"])
            for p in p_remaining:
                if u["Routing"] == p["Routing"] and u_type and u_type == strip_type(p["Type"]):
                    match = p
                    break
            if match:
                p_remaining.remove(match)
                for field in FIELDS:
                    u_val = _get_field_val(u, field)
                    p_val = _get_field_val(match, field)
                    status = _compare_field(field, u_val, p_val, u, match)
                    acc_class = "Net Account" if match.get("IsNet") else "Distribution Account"
                    rows.append({
                        "Employee ID": emp_id,
                        "Employee Name": u["Name"],
                        "Paycom_Account_Class": acc_class,
                        "Field": field,
                        "UZIO_Value": u_val,
                        "Paycom_Value": p_val,
                        "Paycom_SourceOfTruth_Status": status
                    })
            else:
                still_unmatched.append(u)

        # Pass 3: Fallback match on Routing Number ONLY (Last Resort)
        # (For cases like A00Z: Routing matches, but Account# precision lost AND Type code mismatch/unknown)
        final_unmatched = []
        for u in still_unmatched:
            match = None
            u_rout = u["Routing"]
            for p in p_remaining:
                # If routing matches, we assume it's the same account bank-wise
                # This prioritizes Routing Number as the primary key if all else fails
                if u_rout and u_rout == p["Routing"]:
                    match = p
                    break
            
            if match:
                p_remaining.remove(match)
                for field in FIELDS:
                    u_val = _get_field_val(u, field)
                    p_val = _get_field_val(match, field)
                    status = _compare_field(field, u_val, p_val, u, match)
                    acc_class = "Net Account" if match.get("IsNet") else "Distribution Account"
                    rows.append({
                        "Employee ID": emp_id,
                        "Employee Name": u["Name"],
                        "Paycom_Account_Class": acc_class,
                        "Field": field,
                        "UZIO_Value": u_val,
                        "Paycom_Value": p_val,
                        "Paycom_SourceOfTruth_Status": status
                    })
            else:
                final_unmatched.append(u)

        # Uzio accounts that couldn't match at all
        for u in final_unmatched:
            for field in FIELDS:
                u_val = _get_field_val(u, field)
                rows.append({
                    "Employee ID": emp_id,
                    "Employee Name": u["Name"],
                    "Paycom_Account_Class": "Not Found",
                    "Field": field,
                    "UZIO_Value": u_val,
                    "Paycom_Value": "Not Found",
                    "Paycom_SourceOfTruth_Status": STATUS_VAL_MISSING_PAYCOM
                })

        # Paycom accounts unmatched
        for p in p_remaining:
            acc_class = "Net Account" if p.get("IsNet") else "Distribution Account"
            for field in FIELDS:
                p_val = _get_field_val(p, field)
                rows.append({
                    "Employee ID": emp_id,
                    "Employee Name": emp_name,
                    "Paycom_Account_Class": acc_class,
                    "Field": field,
                    "UZIO_Value": "Not Found",
                    "Paycom_Value": p_val,
                    "Paycom_SourceOfTruth_Status": STATUS_VAL_MISSING_UZIO
                })

    # ---------- Build Output DataFrames ----------
    comparison_detail = pd.DataFrame(rows)
    # Employee Status sourced from Paycom (blank for Uzio-only employees not in Paycom)
    comparison_detail["Employee Status"] = comparison_detail["Employee ID"].map(
        lambda e: paycom_status_map.get(e, "")
    )
    comparison_detail = comparison_detail[[
        "Employee ID", "Employee Name", "Employee Status", "Paycom_Account_Class", "Field",
        "UZIO_Value", "Paycom_Value", "Paycom_SourceOfTruth_Status"
    ]]

    mismatches_only = comparison_detail[
        comparison_detail["Paycom_SourceOfTruth_Status"] != STATUS_MATCH
    ].copy()

    # ---------- Field Summary By Status ----------
    status_cols = [
        STATUS_MATCH,
        STATUS_MISMATCH,
        STATUS_LEADING_ZERO,
        STATUS_VAL_MISSING_UZIO,
        STATUS_VAL_MISSING_PAYCOM,
        STATUS_MISSING_UZIO,
        STATUS_MISSING_PAYCOM,
        STATUS_COL_MISSING_PAYCOM,
        STATUS_COL_MISSING_UZIO,
    ]

    pivot = comparison_detail.pivot_table(
        index="Field",
        columns="Paycom_SourceOfTruth_Status",
        values="Employee ID",
        aggfunc="count",
        fill_value=0
    )

    for c in status_cols:
        if c not in pivot.columns:
            pivot[c] = 0

    pivot["Total"] = pivot.sum(axis=1)
    pivot[STATUS_MATCH] = pivot[STATUS_MATCH].astype(int)

    field_summary_by_status = pivot.reset_index()[[
        "Field", "Total",
        STATUS_MATCH, STATUS_MISMATCH, STATUS_LEADING_ZERO,
        STATUS_VAL_MISSING_UZIO,
        STATUS_VAL_MISSING_PAYCOM,
        STATUS_MISSING_UZIO, STATUS_MISSING_PAYCOM,
        STATUS_COL_MISSING_PAYCOM, STATUS_COL_MISSING_UZIO,
    ]]

    # ---------- Leading-Zero Issues (dedicated, actionable sheet) ----------
    lz = comparison_detail[
        comparison_detail["Paycom_SourceOfTruth_Status"] == STATUS_LEADING_ZERO
    ]
    leading_zero_issues = pd.DataFrame({
        "Employee ID": lz["Employee ID"],
        "Employee Name": lz["Employee Name"],
        "Status (Paycom)": lz["Employee Status"],
        "Field": lz["Field"],
        "Correct value (from Paycom)": lz["Paycom_Value"],
        "Wrong value now in Uzio": lz["UZIO_Value"],
    }).reset_index(drop=True)

    # ---------- Summary metrics ----------
    uzio_keys = set(uzio_emp_names.keys())
    paycom_keys = set(paycom_map.keys())

    status_counts = comparison_detail["Paycom_SourceOfTruth_Status"].value_counts().to_dict()
    stats = {
        "employees_checked": len(uzio_keys | paycom_keys),
        "matched": int(status_counts.get(STATUS_MATCH, 0)),
        "leading_zero": int(leading_zero_issues.shape[0]),
        "real_mismatch": int(status_counts.get(STATUS_MISMATCH, 0)),
        "missing_uzio": int(status_counts.get(STATUS_VAL_MISSING_UZIO, 0)) + int(status_counts.get(STATUS_MISSING_UZIO, 0)),
        "missing_paycom": int(status_counts.get(STATUS_VAL_MISSING_PAYCOM, 0)) + int(status_counts.get(STATUS_MISSING_PAYCOM, 0)),
    }

    summary = pd.DataFrame({
        "Metric": [
            "Employees in Uzio sheet",
            "Employees in Paycom sheet",
            "Employees present in both",
            "Employees missing in Paycom (Uzio only)",
            "Employees missing in Uzio (Paycom only)",
            "Total comparison rows",
            "Total NOT OK rows",
            "Leading-zero drops (account/routing)"
        ],
        "Value": [
            len(uzio_keys),
            len(paycom_keys),
            len(uzio_keys & paycom_keys),
            len(uzio_keys - paycom_keys),
            len(paycom_keys - uzio_keys),
            comparison_detail.shape[0],
            mismatches_only.shape[0],
            leading_zero_issues.shape[0]
        ]
    })

    # ---------- "Start Here" plain-English guide (first tab) ----------
    n_lz = int(leading_zero_issues.shape[0])
    start_here = pd.DataFrame({"How to read this report": [
        "This report compares every bank account stored in Uzio against Paycom (the source of truth).",
        "",
        (f"ACTION NEEDED: {n_lz} account number(s) lost a starting 0 when imported into Uzio."
         if n_lz else "Good news: no leading-zero problems were found."),
        ("   These are NOT typos - the Paycom value is correct. Just re-import those accounts."
         if n_lz else "   Every account number kept its leading zeros."),
        "",
        "Tabs in this file:",
        "   - 'Accounts To Re-Import'  ->  the exact accounts to fix, with the correct value to use.",
        "   - 'All Comparisons'        ->  every field checked, each with a plain-English Status.",
        "   - 'Results by Field'       ->  how many matched / mismatched for each field.",
        "   - 'Totals'                 ->  overall counts.",
        "",
        "What each Status means:",
        "   - Data Match            = Uzio and Paycom agree. Nothing to do.",
        "   - Leading Zero Dropped  = Uzio lost a starting 0. Re-import from Paycom.",
        "   - Data Mismatch         = the values are genuinely different. Please review.",
        "   - Value missing in ...  = the account exists in only one of the two systems.",
    ]})

    # ---------- Export report (friendly, human-readable tabs) ----------
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        start_here.to_excel(writer, sheet_name="Start Here", index=False)
        if not leading_zero_issues.empty:
            leading_zero_issues.to_excel(writer, sheet_name="Accounts To Re-Import", index=False)
        comparison_detail.to_excel(writer, sheet_name="All Comparisons", index=False)
        field_summary_by_status.to_excel(writer, sheet_name="Results by Field", index=False)
        summary.to_excel(writer, sheet_name="Totals", index=False)
        writer.sheets["Start Here"].column_dimensions["A"].width = 92
        if not leading_zero_issues.empty:
            _ws = writer.sheets["Accounts To Re-Import"]
            for _c, _w in {"A": 12, "B": 24, "C": 16, "D": 16, "E": 26, "F": 24}.items():
                _ws.column_dimensions[_c].width = _w

    return out.getvalue(), leading_zero_issues, stats


# ---------- Helper: Extract field value from account dict ----------
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


# ---------- Helper: Compare a single field ----------
def _compare_field(field, u_val, p_val, u_acc, p_acc):
    u_n = str(u_val).strip() if u_val else ""
    p_n = str(p_val).strip() if p_val else ""

    # Both blank
    if u_n == "" and p_n == "":
        return STATUS_MATCH
    # One blank
    if u_n == "" and p_n != "":
        return STATUS_VAL_MISSING_UZIO
    if u_n != "" and p_n == "":
        return STATUS_VAL_MISSING_PAYCOM

    # Field-specific comparison
    if field == "Account Type":
        if strip_type(u_n) == strip_type(p_n):
            return STATUS_MATCH
        return STATUS_MISMATCH
    
    if field in ("Amount", "Percent"):
        try:
            diff = abs(float(u_n) - float(p_n))
            if diff < 0.01:
                return STATUS_MATCH
            # Special: both are 0, match
            if float(u_n) == 0.0 and float(p_n) == 0.0:
                return STATUS_MATCH
        except ValueError:
            pass
        return STATUS_MISMATCH

    # Account / Routing: a difference that is ONLY leading zeros is the
    # Excel/Uzio numeric-coercion artifact, not a real data mismatch.
    if field in ("Account Number", "Routing Number") and _is_leading_zero_drop(u_n, p_n):
        return STATUS_LEADING_ZERO

    # Default: exact string match (Routing, Account)
    if u_n == p_n:
        return STATUS_MATCH
    return STATUS_MISMATCH
