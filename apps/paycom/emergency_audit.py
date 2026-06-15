import streamlit as st
import pandas as pd
import io
import re
from datetime import date
from utils.audit_utils import norm_col

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

APP_TITLE = "Paycom vs Uzio – Emergency Contact Audit Tool"

# --- Constants ---
STATUS_MATCH = "Data Match"
STATUS_MISMATCH = "Data Mismatch"
STATUS_MISSING_UZIO = "Missing in Uzio"
STATUS_MISSING_PAYCOM = "Missing in Paycom"
STATUS_EMP_MISSING_UZIO = "Employee ID not in Uzio (present in paycom)"
STATUS_EMP_MISSING_PAYCOM = "Employee ID not in Paycom (Present in uzio)"

def norm_str(x):
    if pd.isna(x) or x is None:
        return ""
    return str(x).strip()

def norm_id(x):
    """Normalize Employee ID (remove .0, strip)."""
    if pd.isna(x): return ""
    s = str(x).strip()
    if s.endswith(".0"): s = s[:-2]
    return s.lstrip("0") # Paycom/Uzio ID match usually requires stripping zeros

def norm_phone(x):
    """Normalize phone to just digits."""
    if pd.isna(x): return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"\D", "", s)
    if not digits: return ""
    if len(digits) == 11 and digits.startswith("1"):
        digits = digits[1:]
    return digits

def norm_relation(x):
    """Normalize relationship (uppercase, strip)."""
    return norm_str(x).upper()

def _get_val(record, field):
    if field == "Name": return record.get("Name", "")
    if field == "Relationship": return record.get("Relation", "")
    if field == "Phone": return record.get("Phone", "")
    if field == "Language": return record.get("Language", "")
    return ""

def _compare_val(field, u_val, p_val):
    u_s = str(u_val).strip().lower()
    p_s = str(p_val).strip().lower()
    
    if u_s == p_s:
        return True
        
    # Phone specific
    if field == "Phone":
        u_p = norm_phone(u_val)
        p_p = norm_phone(p_val)
        if u_p == p_p: return True
        if u_p and p_p and (u_p in p_p or p_p in u_p):
            return True
            
    # Relationship synonyms
    if field == "Relationship":
        if u_s in ["spouse", "husband", "wife"] and p_s in ["spouse", "husband", "wife"]:
            return True
        if u_s in ["mother", "father", "parent"] and p_s in ["mother", "father", "parent"]:
            return True
        if "child" in u_s and "child" in p_s:
            return True

    return False

def run_audit(file_uzio, file_paycom):
    # 1. Load Uzio Data (Same layout as ADP tool)
    df_uzio = pd.read_excel(file_uzio, header=1)
    
    # 2. Load Paycom Data (Census file)
    try:
        if file_paycom.name.lower().endswith('.csv'):
             try:
                 df_paycom = pd.read_csv(file_paycom, dtype=str)
             except UnicodeDecodeError:
                 file_paycom.seek(0)
                 df_paycom = pd.read_csv(file_paycom, dtype=str, encoding='latin1')
        else:
             df_paycom = pd.read_excel(file_paycom, dtype=str)
    except Exception as e:
        st.error(f"Error reading Paycom file: {e}")
        return None

    # Normalize Paycom headers
    df_paycom.columns = [norm_col(c) for c in df_paycom.columns]

    # 3. Map Columns
    # Uzio
    u_map = {
        "EmpID": next((c for c in df_uzio.columns if "Employee ID" in c), "Employee ID"),
        "Name": next((c for c in df_uzio.columns if "Name" in c and "Full" not in c and "Company" not in c), "Name"),
        "Relation": next((c for c in df_uzio.columns if "Relationship" in c), "Relationship"),
        "Phone": next((c for c in df_uzio.columns if "Phone" in c), "Phone")
    }

    # Paycom (Census columns)
    # Emergency_1_Contact, Emergency_1_Language, Emergency_1_Phone, Emergency_1_Relationship
    def get_p_col(keyword):
        # Helper to find column containing keyword (case insensitive)
        for c in df_paycom.columns:
            if keyword.lower() in c.lower():
                return c
        return None

    empid_col = get_p_col("Employee_Code") or get_p_col("Employee ID")
    
    if not empid_col:
        st.error("Could not find Employee ID column in Paycom file.")
        return None

    p_maps = []
    for i in range(1, 4):
        p_maps.append({
            "Name": get_p_col(f"Emergency_{i}_Contact"),
            "Relation": get_p_col(f"Emergency_{i}_Relationship"),
            "Phone": get_p_col(f"Emergency_{i}_Phone"),
            "Language": get_p_col(f"Emergency_{i}_Language")
        })

    # 4. Process Data
    uzio_data = {}
    u_all_eids = set()
    for idx, row in df_uzio.iterrows():
        eid = norm_id(row.get(u_map["EmpID"]))
        if not eid: continue
        u_all_eids.add(eid)
        
        contact = {
            "Name": norm_str(row.get(u_map["Name"])),
            "Relation": norm_relation(row.get(u_map["Relation"])),
            "Phone": norm_phone(row.get(u_map["Phone"])),
            "RawPhone": norm_str(row.get(u_map["Phone"])),
            "Language": "" # Uzio doesn't seem to have language in this report
        }
        if contact["Name"] or contact["Phone"]:
            if eid not in uzio_data: uzio_data[eid] = []
            uzio_data[eid].append(contact)

    paycom_data = {}
    p_all_eids = set()
    for idx, row in df_paycom.iterrows():
        eid = norm_id(row.get(empid_col))
        if not eid: continue
        p_all_eids.add(eid)
        
        for p_map in p_maps:
            # Skip if file doesn't have these columns for this index
            if not p_map["Name"] and not p_map["Phone"]:
                continue

            contact = {
                "Name": norm_str(row.get(p_map["Name"])) if p_map["Name"] else "",
                "Relation": norm_relation(row.get(p_map["Relation"])) if p_map["Relation"] else "",
                "Phone": norm_phone(row.get(p_map["Phone"])) if p_map["Phone"] else "",
                "RawPhone": norm_str(row.get(p_map["Phone"])) if p_map["Phone"] else "",
                "Language": norm_str(row.get(p_map["Language"])) if p_map["Language"] else ""
            }
            
            if contact["Name"] or contact["Phone"]:
                if eid not in paycom_data: paycom_data[eid] = []
                paycom_data[eid].append(contact)

    # 5. Compare
    rows = []
    all_ids = set(uzio_data.keys()) | set(paycom_data.keys())
    FIELDS = ["Name", "Relationship", "Phone"] # We won't compare Language strictly if Uzio doesn't have it

    for eid in sorted(all_ids):
        u_contacts = uzio_data.get(eid, [])
        p_contacts = paycom_data.get(eid, [])
        
        # Missing Employee
        if not u_contacts and p_contacts:
            status = STATUS_EMP_MISSING_UZIO if eid not in u_all_eids else STATUS_MISSING_UZIO
            for p in p_contacts:
                for f in FIELDS:
                    rows.append({
                        "Employee ID": eid,
                        "Status": status,
                        "Field": f,
                        "Uzio Value": "Not Found" if status == STATUS_EMP_MISSING_UZIO else "",
                        "Paycom Value": _get_val(p, f)
                    })
                # Add Language check just for reporting
                rows.append({
                    "Employee ID": eid,
                    "Status": status,
                    "Field": "Language",
                    "Uzio Value": "Not Found" if status == STATUS_EMP_MISSING_UZIO else "",
                    "Paycom Value": p["Language"]
                })
            continue

        if u_contacts and not p_contacts:
            status = STATUS_EMP_MISSING_PAYCOM if eid not in p_all_eids else STATUS_MISSING_PAYCOM
            for u in u_contacts:
                for f in FIELDS:
                    rows.append({
                        "Employee ID": eid,
                        "Status": status,
                        "Field": f,
                        "Uzio Value": _get_val(u, f) if f != "Phone" else u["RawPhone"],
                        "Paycom Value": "Not Found" if status == STATUS_EMP_MISSING_PAYCOM else ""
                    })
            continue

        # Match Contacts
        u_pending = u_contacts[:]
        p_pending = p_contacts[:]
        matched_pairs = []

        # Pass 1: Name Match
        for u in list(u_pending):
            match = None
            for p in p_pending:
                if u["Name"].lower() == p["Name"].lower():
                    match = p
                    break
            if match:
                matched_pairs.append((u, match))
                u_pending.remove(u)
                p_pending.remove(match)

        # Pass 2: Phone Match
        for u in list(u_pending):
            if not u["Phone"]: continue
            match = None
            for p in p_pending:
                if p["Phone"] and u["Phone"] == p["Phone"]:
                    match = p
                    break
            if match:
                matched_pairs.append((u, match))
                u_pending.remove(u)
                p_pending.remove(match)

        # Compare Matched
        for u, p in matched_pairs:
            for f in FIELDS:
                u_val = _get_val(u, f)
                p_val = _get_val(p, f)
                
                u_disp = u["RawPhone"] if f == "Phone" else u_val
                p_disp = p["RawPhone"] if f == "Phone" else p_val
                
                if _compare_val(f, u_val, p_val):
                    status = STATUS_MATCH
                else:
                    status = STATUS_MISMATCH
                
                rows.append({
                    "Employee ID": eid,
                    "Status": status,
                    "Field": f,
                    "Uzio Value": u_disp,
                    "Paycom Value": p_disp
                })
            
            # Show Language (Info Only)
            rows.append({
                "Employee ID": eid,
                "Status": "Info Only",
                "Field": "Language",
                "Uzio Value": "N/A",
                "Paycom Value": p["Language"]
            })

        # Unmatched
        for u in u_pending:
            for f in FIELDS:
                rows.append({
                    "Employee ID": eid,
                    "Status": STATUS_MISSING_PAYCOM,
                    "Field": f,
                    "Uzio Value": _get_val(u, f) if f != "Phone" else u["RawPhone"],
                    "Paycom Value": ""
                })

        for p in p_pending:
            for f in FIELDS:
                rows.append({
                    "Employee ID": eid,
                    "Status": STATUS_MISSING_UZIO,
                    "Field": f,
                    "Uzio Value": "",
                    "Paycom Value": _get_val(p, f)
                })

    df_res = pd.DataFrame(rows)
    
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
        df_res.to_excel(writer, sheet_name='Emergency_Contact_Audit', index=False)
        if not df_res.empty:
            summ = df_res.groupby(["Status", "Field"]).size().reset_index(name="Count")
            summ.to_excel(writer, sheet_name='Summary', index=False)
            
    return out.getvalue()

def render_ui():
    st.title(APP_TITLE)
    st.markdown("""
    **Instructions**:
    1. Upload **Uzio Emergency Contact Export** (.xlsx).
    2. Upload **Paycom Census Export** (.csv or .xlsx).
    
    **Checks**:
    - Emergency_{1,2,3}_Contact (Name)
    - Emergency_{1,2,3}_Relationship
    - Emergency_{1,2,3}_Phone
    - Emergency_{1,2,3}_Language (Info Only)
    """)
    
    client_name = st.text_input("Client Name", value="Client", key="paycom_emergency_client")

    col1, col2 = st.columns(2)
    with col1:
        f_uzio = st.file_uploader("Uzio Emergency Export", type=["xlsx"], key="pec_u")
    with col2:
        f_pay = st.file_uploader("Paycom Census Export", type=["xlsx", "csv"], key="pec_p")

    if st.button("Run Audit", key="run_pec"):
        if not f_uzio or not f_pay:
            st.error("Please upload both files.")
            return
            
        try:
            with st.spinner("Processing..."):
                report = run_audit(f_uzio, f_pay)
                
            if report:
                st.success("Audit Complete!")
                timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
                filename = f"{client_name}_Uzio_Paycom_Emergency_Audit_Report_{timestamp}.xlsx"

                st.download_button(
                    "Download Report",
                    data=report,
                    file_name=filename
                )
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
