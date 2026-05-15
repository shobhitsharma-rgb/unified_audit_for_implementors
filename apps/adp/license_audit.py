import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime, date
import numpy as np

# --- Monkey-patch openpyxl to handle non-ISO datetime strings gracefully ---
try:
    import openpyxl.utils.datetime as _openpyxl_dt
    _original_from_ISO8601 = _openpyxl_dt.from_ISO8601
    def _patched_from_ISO8601(formatted_string):
        try:
            return _original_from_ISO8601(formatted_string)
        except ValueError:
            return formatted_string
    _openpyxl_dt.from_ISO8601 = _patched_from_ISO8601
except Exception:
    pass

APP_TITLE = "ADP License Details Audit"

def norm_blank(x):
    if pd.isna(x) or x is None: return ""
    if isinstance(x, str):
        s = x.strip()
        return s if s else ""
    return x

def try_parse_date(x):
    x = norm_blank(x)
    if x == "": return ""
    if isinstance(x, (datetime, date, np.datetime64, pd.Timestamp)):
        return pd.to_datetime(x).strftime('%m/%d/%Y')
    if isinstance(x, str):
        s = x.strip().split(' ')[0]
        try: return pd.to_datetime(s, errors="raise").strftime('%m/%d/%Y')
        except Exception: return s
    return str(x)

def read_uzio_license(file) -> pd.DataFrame:
    try:
        df = pd.read_excel(file, header=None, dtype=str)
        header_idx = -1
        for i, row in df.head(20).iterrows():
            if any(str(c).strip() == 'Employee ID' for c in row.values if pd.notna(c)):
                header_idx = i; break
        if header_idx != -1:
            df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.iloc[header_idx])]
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
            return df.replace('None', '').dropna(how='all')
        else:
            file.seek(0); return pd.read_excel(file, dtype=str)
    except Exception as e: st.error(f"Could not read Uzio file: {e}"); return None

def read_adp_license(file) -> pd.DataFrame:
    try:
        df = pd.read_excel(file, header=None, dtype=str)
        header_idx = -1
        for i, row in df.head(20).iterrows():
            if any(str(c).strip() == 'Associate ID' for c in row.values if pd.notna(c)):
                header_idx = i; break
        if header_idx != -1:
            df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.iloc[header_idx])]
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
            return df.replace('None', '').dropna(how='all')
        else:
            file.seek(0); return pd.read_excel(file, dtype=str)
    except Exception as e: st.error(f"Could not read ADP file: {e}"); return None

def run_license_audit(uzio_df, adp_df):
    UZIO_KEY, UZIO_NUM_COL, UZIO_DATE_COL = 'Employee ID', 'License Number', 'License Expiration Date'
    ADP_KEY = 'Associate ID'
    ADP_NUM_COL = 'License/Certification Code' if 'License/Certification Code' in adp_df.columns else 'License/Certification ID'
    ADP_DATE_COL = 'Expiration Date'
    
    uzio_name_col = next((c for c in uzio_df.columns if str(c).strip().lower() == 'full name'), None)
    adp_fname_col = next((c for c in adp_df.columns if str(c).strip().lower() == 'legal first name'), None)
    adp_lname_col = next((c for c in adp_df.columns if str(c).strip().lower() == 'legal last name'), None)
    
    uzio_df[UZIO_KEY] = uzio_df[UZIO_KEY].apply(lambda x: str(x).strip() if norm_blank(x) != "" else "")
    adp_df[ADP_KEY] = adp_df[ADP_KEY].apply(lambda x: str(x).strip() if norm_blank(x) != "" else "")
    
    def clean_val(val):
        v = norm_blank(val)
        if v == "": return ""
        s = str(v).strip()
        return "" if s.lower() in ('nan', 'none', 'nat') else s

    uzio_map = {}; [uzio_map.setdefault(r[UZIO_KEY], []).append(r) for r in uzio_df.to_dict('records') if r[UZIO_KEY]]
    adp_map = {}; [adp_map.setdefault(r[ADP_KEY], []).append(r) for r in adp_df.to_dict('records') if r[ADP_KEY]]
    
    rows = []; processed_adp_pairs = set()
    
    for eid, uzio_recs in uzio_map.items():
        for uz_rec in uzio_recs:
            uz_num = clean_val(uz_rec.get(UZIO_NUM_COL, ""))
            if not uz_num: continue
            uz_date = try_parse_date(uz_rec.get(UZIO_DATE_COL, ""))
            emp_name = clean_val(uz_rec.get(uzio_name_col, "")) if uzio_name_col else ""
            
            adp_licenses = adp_map.get(eid, [])
            match_found = False; matched_adp_rec = None
            for adp_rec in adp_licenses:
                adp_num = clean_val(adp_rec.get(ADP_NUM_COL, ""))
                if adp_num and uz_num.lower() == adp_num.lower():
                    match_found = True; matched_adp_rec = adp_rec
                    processed_adp_pairs.add((eid, adp_num.lower())); break
            
            if not emp_name and matched_adp_rec:
                emp_name = f"{clean_val(matched_adp_rec.get(adp_fname_col, ''))} {clean_val(matched_adp_rec.get(adp_lname_col, ''))}".strip()
            
            if match_found:
                rows.append({"Employee ID": eid, "Employee Name": emp_name, "Field": "License Number", "Status": "Data Match", "Uzio Value": uz_num, "ADP Value": clean_val(matched_adp_rec.get(ADP_NUM_COL, ""))})
                adp_date = try_parse_date(matched_adp_rec.get(ADP_DATE_COL, ""))
                if uz_date == adp_date: status = "Data Match"
                elif uz_date == "" and adp_date != "": status = "Value missing in Uzio (ADP has value)"
                elif uz_date != "" and adp_date == "": status = "Value missing in ADP (Uzio has value)"
                else: status = "Data Mismatch"
                rows.append({"Employee ID": eid, "Employee Name": emp_name, "Field": "Expiration Date", "Status": status, "Uzio Value": uz_date, "ADP Value": adp_date})
            else:
                rows.append({"Employee ID": eid, "Employee Name": emp_name, "Field": "License Number", "Status": "Missing in ADP", "Uzio Value": uz_num, "ADP Value": ""})
    
    for eid, adp_recs in adp_map.items():
        for adp_rec in adp_recs:
            adp_num = clean_val(adp_rec.get(ADP_NUM_COL, ""))
            if not adp_num or (eid, adp_num.lower()) in processed_adp_pairs: continue
            emp_name = f"{clean_val(adp_rec.get(adp_fname_col, ''))} {clean_val(adp_rec.get(adp_lname_col, ''))}".strip()
            status = "Employee Not Found in Uzio" if eid not in uzio_map else "Missing in Uzio"
            rows.append({"Employee ID": eid, "Employee Name": emp_name, "Field": "License Number", "Status": status, "Uzio Value": "", "ADP Value": adp_num})
            adp_date = try_parse_date(adp_rec.get(ADP_DATE_COL, ""))
            if adp_date: rows.append({"Employee ID": eid, "Employee Name": emp_name, "Field": "Expiration Date", "Status": status, "Uzio Value": "", "ADP Value": adp_date})
    
    return pd.DataFrame(rows).fillna("")

def render_ui():
    st.title(APP_TITLE)
    client_name = st.text_input("Client Name", key="adp_license_client_name")
    col1, col2 = st.columns(2)
    with col1: uzio_file = st.file_uploader("Upload UZIO License Report (.xlsx)", type=['xlsx', 'csv'], key='uzio_license_upload')
    with col2: adp_file = st.file_uploader("Upload ADP License Report (.xlsx)", type=['xlsx', 'csv'], key='adp_license_upload')
    if uzio_file and adp_file:
        try:
            with st.spinner("Extracting..."): uzio_df, adp_df = read_uzio_license(uzio_file), read_adp_license(adp_file)
            if uzio_df is not None and adp_df is not None:
                if st.button("Run License Audit", type="primary"):
                    if not client_name: st.warning("Please enter a Client Name."); return
                    with st.spinner("Running Audit..."): result_df = run_license_audit(uzio_df, adp_df)
                    if result_df is not None:
                        st.success("Audit Complete!")
                        st.dataframe(result_df)
                        buffer = io.BytesIO()
                        with pd.ExcelWriter(buffer, engine='openpyxl') as writer: result_df.to_excel(writer, index=False)
                        st.download_button(label="Download Excel Report", data=buffer.getvalue(), file_name=f"{client_name.strip()}_License_Audit_Report_{datetime.now().strftime('%d_%m_%Y')}.xlsx")
        except Exception as e: st.error(f"Error: {e}")
