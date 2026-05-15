import streamlit as st
import pandas as pd
import io
import re
from datetime import date

APP_TITLE = "ADP vs Uzio – Emergency Contact Audit Tool"

# --- Constants ---
STATUS_MATCH = "Data Match"
STATUS_MISMATCH = "Data Mismatch"
STATUS_MISSING_UZIO = "Missing in Uzio"
STATUS_MISSING_ADP = "Missing in ADP"
STATUS_EMP_MISSING_UZIO = "Employee ID not in Uzio (present in adp)"
STATUS_EMP_MISSING_ADP = "Employee ID not in ADP (Present in uzio)"

def norm_str(x):
    if pd.isna(x) or x is None:
        return ""
    return str(x).strip()

def norm_id(x):
    """Normalize Employee ID (remove .0, strip)."""
    if pd.isna(x): return ""
    s = str(x).strip()
    if s.endswith(".0"): s = s[:-2]
    return s

def norm_phone(x):
    """Normalize phone to just digits."""
    if pd.isna(x): return ""
    s = str(x).strip()
    if s.endswith(".0"):
        s = s[:-2]
    # Remove all non-digit chars
    digits = re.sub(r"\D", "", s)
    # If empty, return empty
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
    return ""

def _compare_val(field, u_val, a_val):
    u_s = str(u_val).strip().lower()
    a_s = str(a_val).strip().lower()
    
    if u_s == a_s:
        return True
        
    # Phone specific: Compare last 10 digits if length differs
    if field == "Phone":
        u_p = norm_phone(u_val)
        a_p = norm_phone(a_val)
        if u_p == a_p: return True
        # If one contains the other (e.g. +1 extension)
        if u_p and a_p and (u_p in a_p or a_p in u_p):
            return True
            
    # Relationship synonyms (Basic)
    if field == "Relationship":
        # Spouse vs Husband/Wife
        if u_s in ["spouse", "husband", "wife"] and a_s in ["spouse", "husband", "wife"]:
            return True
        # Mother/Father vs Parent
        if u_s in ["mother", "father"] and a_s == "parent":
            return True
        if a_s in ["mother", "father"] and u_s == "parent":
            return True

    return False

def run_audit(file_uzio, file_adp):
    # 1. Load Data
    # Uzio: Header=1 based on inspection
    df_uzio = pd.read_excel(file_uzio, header=1)
    df_adp = pd.read_excel(file_adp)
    
    df_adp.columns = [str(c).strip() for c in df_adp.columns]
    
    # 2. Map Columns
    u_map = {
        "EmpID": next((c for c in df_uzio.columns if "Employee ID" in c), "Employee ID"),
        "Name": next((c for c in df_uzio.columns if "Name" in c and "Full" not in c and "Company" not in c), "Name"),
        "Relation": next((c for c in df_uzio.columns if "Relationship" in c), "Relationship"),
        "Phone": next((c for c in df_uzio.columns if "Phone" in c), "Phone")
    }
    
    a_map = {
        "EmpID": next((c for c in df_adp.columns if "Associate ID" in c), "Associate ID"),
        "Name": next((c for c in df_adp.columns if "Contact Name" in c), "Contact Name"),
        "Relation": next((c for c in df_adp.columns if "Relationship Description" in c), "Relationship Description"),
        "Phone": next((c for c in df_adp.columns if "Mobile Phone" in c), "Mobile Phone")
    }

    # 3. Group by Employee
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
            "RawPhone": norm_str(row.get(u_map["Phone"]))
        }
        if contact["Name"] or contact["Phone"]:
            if eid not in uzio_data: uzio_data[eid] = []
            uzio_data[eid].append(contact)

    adp_data = {}
    a_all_eids = set()
    for idx, row in df_adp.iterrows():
        eid = norm_id(row.get(a_map["EmpID"]))
        if not eid: continue
        a_all_eids.add(eid)
        
        contact = {
            "Name": norm_str(row.get(a_map["Name"])),
            "Relation": norm_relation(row.get(a_map["Relation"])),
            "Phone": norm_phone(row.get(a_map["Phone"])),
            "RawPhone": norm_str(row.get(a_map["Phone"]))
        }
        if contact["Name"] or contact["Phone"]:
            if eid not in adp_data: adp_data[eid] = []
            adp_data[eid].append(contact)
            
    # 4. Compare
    rows = []
    all_ids = set(uzio_data.keys()) | set(adp_data.keys())
    FIELDS = ["Name", "Relationship", "Phone"]

    for eid in sorted(all_ids):
        u_contacts = uzio_data.get(eid, [])
        a_contacts = adp_data.get(eid, [])
        
        # Missing Employee
        if not u_contacts and a_contacts:
            status = STATUS_EMP_MISSING_UZIO if eid not in u_all_eids else STATUS_MISSING_UZIO
            for a in a_contacts:
                for f in FIELDS:
                    rows.append({
                        "Employee ID": eid,
                        "Status": status,
                        "Field": f,
                        "Uzio Value": "Not Found" if status == STATUS_EMP_MISSING_UZIO else "",
                        "ADP Value": _get_val(a, f)
                    })
            continue
            
        if u_contacts and not a_contacts:
            status = STATUS_EMP_MISSING_ADP if eid not in a_all_eids else STATUS_MISSING_ADP
            for u in u_contacts:
                for f in FIELDS:
                    rows.append({
                        "Employee ID": eid,
                        "Status": status,
                        "Field": f,
                        "Uzio Value": _get_val(u, f) if f != "Phone" else u["RawPhone"],
                        "ADP Value": "Not Found" if status == STATUS_EMP_MISSING_ADP else ""
                    })
            continue

        # Match Contacts
        u_pending = u_contacts[:]
        a_pending = a_contacts[:]
        matched_pairs = []
        
        # Pass 1: Exact Name Match
        for u in list(u_pending):
            match = None
            for a in a_pending:
                if u["Name"].lower() == a["Name"].lower():
                    match = a
                    break
            if match:
                matched_pairs.append((u, match))
                u_pending.remove(u)
                a_pending.remove(match)
                
        # Pass 2: Phone Match
        for u in list(u_pending):
            if not u["Phone"]: continue
            match = None
            for a in a_pending:
                if a["Phone"] and u["Phone"] == a["Phone"]:
                    match = a
                    break
            if match:
                matched_pairs.append((u, match))
                u_pending.remove(u)
                a_pending.remove(match)

        # Compare Matched
        for u, a in matched_pairs:
            for f in FIELDS:
                u_val = _get_val(u, f)
                a_val = _get_val(a, f)
                
                u_disp = u["RawPhone"] if f == "Phone" else u_val
                a_disp = a["RawPhone"] if f == "Phone" else a_val
                
                if _compare_val(f, u_val, a_val):
                    status = STATUS_MATCH
                else:
                    status = STATUS_MISMATCH
                    
                rows.append({
                    "Employee ID": eid,
                    "Status": status,
                    "Field": f,
                    "Uzio Value": u_disp,
                    "ADP Value": a_disp
                })
        
        # Unmatched
        for u in u_pending:
            for f in FIELDS:
                rows.append({
                    "Employee ID": eid,
                    "Status": STATUS_MISSING_ADP,
                    "Field": f,
                    "Uzio Value": _get_val(u, f) if f != "Phone" else u["RawPhone"],
                    "ADP Value": ""
                })
                
        for a in a_pending:
            for f in FIELDS:
                rows.append({
                    "Employee ID": eid,
                    "Status": STATUS_MISSING_UZIO,
                    "Field": f,
                    "Uzio Value": "",
                    "ADP Value": _get_val(a, f) if f != "Phone" else a["RawPhone"]
                })

    df_res = pd.DataFrame(rows)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_res.to_excel(writer, sheet_name='Emergency_Contact_Audit', index=False)
        if not df_res.empty:
            summ = df_res.groupby(["Status", "Field"]).size().reset_index(name="Count")
            summ.to_excel(writer, sheet_name='Summary', index=False)
            
    return output.getvalue()

def render_ui():
    st.title(APP_TITLE)
    client_name = st.text_input("Client Name", value="Client", key="adp_emergency_client")

    st.markdown("""
    **Instructions**:
    1. Upload **Uzio Emergency Contact Export** (`.xlsx`).
    2. Upload **ADP Emergency Contact Export** (`.xlsx`).
    
    **Logic**:
    - Matches contacts by **Name** or **Phone**.
    - Normalizes Phone numbers (digits only).
    - Compares Name, Relationship, and Phone.
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        f_uzio = st.file_uploader("Uzio Emergency Input", type=["xlsx"], key="ec_u")
    with col2:
        f_adp = st.file_uploader("ADP Emergency Input", type=["xlsx"], key="ec_a")
        
    if st.button("Run Emergency Audit"):
        if not f_uzio or not f_adp:
            st.error("Please upload both files.")
            return
        
        try:
            with st.spinner("Processing..."):
                report = run_audit(f_uzio, f_adp)
            
            st.success("Audit Complete!")
            timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
            filename = f"{client_name}_Uzio_ADP_Emergency_Audit_Report_{timestamp}.xlsx"

            st.download_button(
                "Download Report",
                data=report,
                file_name=filename
            )
        except Exception as e:
            st.error(f"Error: {e}")
            st.exception(e)
