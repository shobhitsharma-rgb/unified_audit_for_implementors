import streamlit as st
import pandas as pd
import numpy as np
import io
import re
from datetime import datetime, date

# --- Monkeypatch for openpyxl to avoid Invalid datetime value errors ---
import openpyxl.cell.cell
if not hasattr(openpyxl.cell.cell.Cell, '_patched_for_datetime'):
    _orig_bind_value = openpyxl.cell.cell.Cell._bind_value
    def _safe_bind_value(self, value):
        try:
            _orig_bind_value(self, value)
        except ValueError as e:
            if "Invalid datetime value" in str(e):
                self.data_type = 's'
                self._value = str(value)
            else:
                raise
    openpyxl.cell.cell.Cell._bind_value = _safe_bind_value
    openpyxl.cell.cell.Cell._patched_for_datetime = True
# -----------------------------------------------------------------------

# =========================================================
# ADP vs Uzio Payment Audit Tool
# =========================================================

APP_TITLE = "ADP vs Uzio – Payment Audit Tool"

# --- Constants for Status ---
STATUS_MATCH = "Data Match"
STATUS_MISMATCH = "Data Mismatch"
STATUS_VAL_MISSING_UZIO = "Value missing in Uzio (ADP has value)"
STATUS_VAL_MISSING_ADP = "Value missing in ADP (Uzio has value)"
STATUS_MISSING_UZIO = "Employee ID Not Found in Uzio"
STATUS_MISSING_ADP = "Employee ID Not Found in ADP"
STATUS_COL_MISSING_ADP = "Column Missing in ADP Sheet"
STATUS_COL_MISSING_UZIO = "Column Missing in Uzio Sheet"

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
    s = str(x).replace(",", "").replace("$", "").replace("%", "").strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except:
        return 0.0

def norm_id(x):
    """Standardize ID. For ADP, we usually keep as-is or pad depending on Uzio.
    Since Paycom used 4-digit padding, we'll try to be flexible."""
    if x is None: return ""
    s = str(x).strip()
    if s.endswith(".0"): 
        s = s[:-2]
    return s

def normalize_account_type(t):
    """Normalize ADP Deduction/Account Type to standard 'Checking'/'Savings'."""
    if not t: return ""
    s = str(t).strip().lower()
    if "checking" in s or "ck" in s:
        return "Checking"
    if "savings" in s or "sv" in s:
        return "Savings"
    return str(t).strip()

def _get_field_val(record, field):
    """Helper to extract field value from a record dict."""
    if field == "Routing Number": return record.get("Routing", "")
    if field == "Account Number": return record.get("Account", "")
    if field == "Account Type": return record.get("Type", "")
    if field == "Amount": return record.get("Amount", 0.0)
    if field == "Percent": return record.get("Percent", 0.0)
    return ""

def _compare_field(field, u_val, p_val, u_rec, p_rec):
    """Compare single field values."""
    if field in ["Amount", "Percent"]:
        try:
            u_f = float(u_val)
            p_f = float(p_val)
            # Allow small float diff
            if abs(u_f - p_f) < 0.01:
                return STATUS_MATCH
            return STATUS_MISMATCH
        except:
            pass
            
    if str(u_val).strip().lower() == str(p_val).strip().lower():
        return STATUS_MATCH
        
    if not u_val and p_val:
        return STATUS_VAL_MISSING_UZIO
    if u_val and not p_val:
        return STATUS_VAL_MISSING_ADP
        
    return STATUS_MISMATCH

def _read_payment_file(file, header=0):
    """Read an .xlsx/.xls/.csv payment export. Header is the row index of column names."""
    name = (getattr(file, 'name', '') or '').lower()
    file.seek(0)
    if name.endswith('.csv'):
        try:
            return pd.read_csv(file, header=header, dtype=str)
        except UnicodeDecodeError:
            file.seek(0)
            return pd.read_csv(file, header=header, dtype=str, encoding='latin1')
    return pd.read_excel(file, header=header, dtype=str)

def run_audit(file_uzio, file_adp):
    # 1. Load Uzio Data
    # Uzio Export typically starts at Row 2 (Header=1)
    df_uzio = _read_payment_file(file_uzio, header=1)
    
    # Map Uzio Columns
    # Clean column names first (remove newlines/extra spaces)
    df_uzio.columns = [str(c).strip().replace("\n", " ") for c in df_uzio.columns]
    
    def get_col(candidates):
        for cand in candidates:
            # Exact match
            if cand in df_uzio.columns: return cand
            # Partial match
            match = next((c for c in df_uzio.columns if cand in c), None)
            if match: return match
        return candidates[0] # Default
    
    u_cols = {
        "EmpID": get_col(["Employee ID", "Emp Code", "EmpID"]),
        "Routing": get_col(["Routing Number", "Routing"]),
        "Account": get_col(["Account Number", "Account"]),
        "Type": get_col(["Account Type", "Type"]),
        "Percent": get_col(["Paycheck Percentage", "Deposit Percent"]),
        "Amount": get_col(["Paycheck Amount", "Deposit Amount"]),
        "Name": get_col(["Full Name", "Employee Name", "Name"])
    }
    
    uzio_map = {} # EmpID -> List of Accounts
    all_names = {} # EmpID -> Name
    
    for idx, row in df_uzio.iterrows():
        emp_id = norm_id(row.get(u_cols["EmpID"]))
        # Also try "Employee ID" if mapped column failed (fallback safety)
        if not emp_id and "Employee ID" in df_uzio.columns:
             emp_id = norm_id(row.get("Employee ID"))
             
        if not emp_id: continue
        
        # Capture name
        raw_name = norm_str(row.get(u_cols["Name"]))
        if emp_id not in all_names and raw_name:
            all_names[emp_id] = raw_name
            
        # Initialize list if new
        if emp_id not in uzio_map:
            uzio_map[emp_id] = []
        
        acc = {
            "Routing": norm_digits(row.get(u_cols["Routing"])),
            "Account": norm_digits(row.get(u_cols["Account"])),
            "Type": normalize_account_type(row.get(u_cols["Type"])),
            "Percent": norm_money(row.get(u_cols["Percent"])),
            "Amount": norm_money(row.get(u_cols["Amount"])),
            "Name": raw_name
        }
        
        # Only add valid accounts (must have Rout or Acc)
        if acc["Routing"] or acc["Account"]:
            if acc not in uzio_map[emp_id]:
                uzio_map[emp_id].append(acc)

    # 2. Load ADP Data
    df_adp = _read_payment_file(file_adp, header=0)
    df_adp.columns = [str(c).strip() for c in df_adp.columns]
    
    # Map ADP Columns
    a_cols = {
        "EmpID": next((c for c in df_adp.columns if "ASSOCIATE ID" in c.upper()), "ASSOCIATE ID"),
        "Routing": next((c for c in df_adp.columns if "ROUTING NUMBER" in c.upper()), "ROUTING NUMBER"),
        "Account": next((c for c in df_adp.columns if "ACCOUNT NUMBER" in c.upper()), "ACCOUNT NUMBER"),
        "Deduction": next((c for c in df_adp.columns if "DEDUCTION" in c.upper()), "DEDUCTION"), # Account Type
        "DepositType": next((c for c in df_adp.columns if "DEPOSIT TYPE" in c.upper()), "DEPOSIT TYPE"),
        "Percent": next((c for c in df_adp.columns if "DEPOSIT PERCENT" in c.upper()), "DEPOSIT PERCENT"),
        "Amount": next((c for c in df_adp.columns if "DEPOSIT AMOUNT" in c.upper()), "DEPOSIT AMOUNT"),
        "Name": next((c for c in df_adp.columns if "NAME" in c.upper()), "NAME")
    }
    
    adp_map = {}
    
    for idx, row in df_adp.iterrows():
        # Raw ID from ADP
        raw_id = row.get(a_cols["EmpID"])
        emp_id = norm_id(raw_id)
        if not emp_id: continue
        
        # Capture name
        raw_name = norm_str(row.get(a_cols["Name"]))
        if emp_id not in all_names and raw_name:
            all_names[emp_id] = raw_name
            
        if emp_id not in adp_map:
            adp_map[emp_id] = []
        
        # Analyze Deposit Type
        dep_type = str(row.get(a_cols["DepositType"])).strip()
        raw_pct = row.get(a_cols["Percent"])
        raw_amt = row.get(a_cols["Amount"])
        
        pct = 0.0
        amt = 0.0
        is_net = False
        
        if "Full" in dep_type or "Balance" in dep_type:
            # Do NOT default to 100.0 immediately. Respect the source value.
            # We will handle the "Remainder" or "Single Account = 100%" logic in post-processing.
            pct = norm_money(raw_pct)
            is_net = True
        elif "Partial %" in dep_type:
             pct = norm_money(raw_pct)
        elif "Partial" in dep_type:
             amt = norm_money(raw_amt)
             
        acc = {
            "EmpID": emp_id,
            "Routing": norm_digits(row.get(a_cols["Routing"])),
            "Account": norm_digits(row.get(a_cols["Account"])),
            "Type": normalize_account_type(row.get(a_cols["Deduction"])),
            "Percent": pct,
            "Amount": amt,
            "Name": raw_name,
            "IsNet": is_net
        }
        
        if acc["Routing"] or acc["Account"]:
            if acc not in adp_map[emp_id]:
                adp_map[emp_id].append(acc)

    # 2b. Post-Process ADP "Full" / Net Pay
    # Logic:
    # 1. Single Account -> Always 100%
    # 2. Multi-Account + Partial % -> Remainder (100 - sum)
    # 3. Multi-Account + No Partial % -> Keep parsed value (Blank/0 stays Blank/0)
    for emp_id, accs in adp_map.items():
        if not accs: continue
        
        # Rule 1: Single Account
        if len(accs) == 1:
            accs[0]["Percent"] = 100.0
            continue

        # Rule 2/3: Check for Net Pay / Full accounts logic
        net_accs = [a for a in accs if a.get("IsNet")]
        partial_accs = [a for a in accs if not a.get("IsNet") and a.get("Percent") > 0]
        
        if len(net_accs) == 1 and partial_accs:
            # Calculate sum of partials
            total_partial = sum(a["Percent"] for a in partial_accs)
            
            # If we have partial percentages, the Net account IS the remainder.
            # Even if the source was 0 or blank.
            if total_partial < 100.0:
                remainder = 100.0 - total_partial
                if remainder > 0:
                     net_accs[0]["Percent"] = round(remainder, 2)

    # 3. Comparison Logic
    FIELDS = ["Routing Number", "Account Number", "Account Type", "Amount", "Percent"]
    rows = []
    
    all_ids = set(uzio_map.keys()) | set(adp_map.keys())
    
    for emp_id in sorted(all_ids):
        u_accs = uzio_map.get(emp_id, [])
        a_accs = adp_map.get(emp_id, [])
        
        emp_name = all_names.get(emp_id, "")
        
        in_uzio = emp_id in uzio_map
        in_adp = emp_id in adp_map
        
        # Case 1: Missing in Uzio (Present in ADP, Not in Uzio map)
        if in_adp and not in_uzio:
            if not a_accs:
                 # In ADP map but no accounts? (Rare, but possible if blank lines)
                 continue
            for a in a_accs:
                for field in FIELDS:
                    rows.append({
                        "Employee ID": emp_id,
                        "Employee Name": emp_name,
                        "Field": field,
                        "UZIO_Value": "Not Found",
                        "ADP_Value": _get_field_val(a, field),
                        "Status": STATUS_MISSING_UZIO
                    })
            continue

        # Case 2: Missing in ADP (Present in Uzio, Not in ADP map)
        if in_uzio and not in_adp:
            if not u_accs:
                # Employee valid in Uzio (e.g. Paper Check) but missing in ADP
                # Still report as Missing in ADP per request
                for field in FIELDS:
                    rows.append({
                        "Employee ID": emp_id,
                        "Employee Name": emp_name,
                        "Field": field,
                        "UZIO_Value": "No Account Info",
                        "ADP_Value": "Not Found",
                        "Status": STATUS_MISSING_ADP
                    })
            else:
                for u in u_accs:
                    for field in FIELDS:
                        rows.append({
                            "Employee ID": emp_id,
                            "Employee Name": emp_name,
                            "Field": field,
                            "UZIO_Value": _get_field_val(u, field),
                            "ADP_Value": "Not Found",
                            "Status": STATUS_MISSING_ADP
                        })
            continue

        # Case 3: Both Exist (ID is in both maps)
        if not u_accs and not a_accs:
            # Both present but neither has accounts. Ignore? 
            # Or match? User didn't specify. Assuming ignore or "Data Match" on empty state.
            continue
            
        if u_accs and not a_accs:
             # In both maps, but ADP has empty accounts list.
             # Treat as mismatch? Or "Value missing in ADP"?
             # Logic below for "unmatched UZIO" covers this if we don't break early.
             pass

        if not u_accs and a_accs:
             # In both maps, but Uzio has empty accounts.
             pass
             
        # Strategy: Match by Account Number first (Unique ID usually)
        u_remaining = u_accs[:]
        a_remaining = a_accs[:]
        
        # Pass 1: Exact Account Number
        matched_pairs = []
        for u in list(u_remaining):
            match = None
            for a in a_remaining:
                if u["Account"] and u["Account"] == a["Account"]:
                    match = a
                    break
            if match:
                matched_pairs.append((u, match))
                u_remaining.remove(u)
                a_remaining.remove(match)

        # Pass 2: Exact Routing (fallback if account is masked/missing but unlikely)
        for u in list(u_remaining):
            match = None
            for a in a_remaining:
                if u["Routing"] and u["Routing"] == a["Routing"] and u["Type"] == a["Type"]:
                    match = a
                    break
            if match:
                matched_pairs.append((u, match))
                u_remaining.remove(u)
                a_remaining.remove(match)

        # Compare Matched
        for u, a in matched_pairs:
            for field in FIELDS:
                 u_val = _get_field_val(u, field)
                 a_val = _get_field_val(a, field)
                 status = _compare_field(field, u_val, a_val, u, a)
                 
                 rows.append({
                    "Employee ID": emp_id,
                    "Employee Name": emp_name,
                    "Field": field,
                    "UZIO_Value": u_val,
                    "ADP_Value": a_val,
                    "Status": status
                 })
                 
        # Unmatched UZIO (accounts in Uzio not matched to any in ADP)
        for u in u_remaining:
             for field in FIELDS:
                rows.append({
                    "Employee ID": emp_id,
                    "Employee Name": emp_name,
                    "Field": field,
                    "UZIO_Value": _get_field_val(u, field),
                    "ADP_Value": "Not Found",
                    "Status": STATUS_MISMATCH
                })
        
        # Unmatched ADP (accounts in ADP not matched to any in Uzio)
        for a in a_remaining:
            for field in FIELDS:
                rows.append({
                    "Employee ID": emp_id,
                    "Employee Name": emp_name,
                    "Field": field,
                    "UZIO_Value": "Not Found",
                    "ADP_Value": _get_field_val(a, field),
                    "Status": STATUS_MISMATCH
                })

    df_res = pd.DataFrame(rows)
    
    # --- Generate Excel ---
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_res.to_excel(writer, sheet_name='Comparison_Detail', index=False)
        
        if not df_res.empty:
            summary = df_res.groupby(["Status", "Field"]).size().reset_index(name="Count")
            summary.to_excel(writer, sheet_name='Summary', index=False)
            
    return output.getvalue()

def render_ui():
    st.title(APP_TITLE)
    client_name = st.text_input("Client Name", value="Client", key="adp_payment_client")

    st.markdown("""
    **Instructions**:
    1. Upload **Uzio Payment Export** (`HR Report_...xlsx` or `.csv`).
    2. Upload **ADP Payment Export** (`.xlsx`, `.xls`, or `.csv`).
    
    **Notes**:
    - **ADP Account Type** ('CK1 - checking') is normalized to 'Checking'/'Savings'.
    - **ADP Deposit Type** ('Full', 'Partial %', 'Partial') is mapped to Percent/Amount.
    - **Routing/Account Numbers** are stripped of leading zeros for comparison.
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        uzio_file = st.file_uploader("Upload Uzio Payment Export", type=["xlsx", "csv"], key="u_pay")
    with col2:
        adp_file = st.file_uploader("Upload ADP Payment Export", type=["xlsx", "xls", "csv"], key="a_pay")
        
    if st.button("Run Audit"):
        if not uzio_file or not adp_file:
            st.error("Please upload both files.")
            return
            
        try:
            with st.spinner("Processing..."):
                report_bytes = run_audit(uzio_file, adp_file)
            
            st.success("Audit Complete!")
            timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
            filename = f"{client_name}_Uzio_ADP_Payment_Audit_Report_{timestamp}.xlsx"

            st.download_button(
                label="Download Audit Report",
                data=report_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
