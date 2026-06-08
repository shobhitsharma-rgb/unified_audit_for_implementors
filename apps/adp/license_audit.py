import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime, date
import numpy as np

# --- Monkey-patch openpyxl to handle non-ISO datetime strings gracefully ---
# openpyxl's from_ISO8601 raises ValueError for date formats like '06/29/2023'
# which can appear in some Excel files. This patch returns the raw string instead of crashing.
try:
    import openpyxl.utils.datetime as _openpyxl_dt
    _original_from_ISO8601 = _openpyxl_dt.from_ISO8601

    def _patched_from_ISO8601(formatted_string):
        try:
            return _original_from_ISO8601(formatted_string)
        except ValueError:
            # Return the raw string instead of crashing
            return formatted_string
    
    _openpyxl_dt.from_ISO8601 = _patched_from_ISO8601
except Exception:
    pass  # If openpyxl is not installed or structure changed, skip silently

APP_TITLE = "ADP License Details Audit"

# --- HELPER FUNCTIONS ---
def norm_blank(x):
    """Normalize NaN, None, or completely whitespace strings to empty string."""
    if pd.isna(x) or x is None:
        return ""
    if isinstance(x, str):
        s = x.strip()
        if not s:
             return ""
        return s
    return x

def try_parse_date(x):
    """Safely parse a date string or object into YYYY-MM-DD string."""
    x = norm_blank(x)
    if x == "":
        return ""
    if isinstance(x, (datetime, date, np.datetime64, pd.Timestamp)):
        return pd.to_datetime(x).strftime('%m/%d/%Y')
    if isinstance(x, str):
        # Handle '1900-01-01 00:00:00' common Excel raw string formats
        s = x.strip().split(' ')[0]
        try:
            return pd.to_datetime(s, errors="raise").strftime('%m/%d/%Y')
        except Exception:
            return s
    return str(x)

def read_uzio_license(file) -> pd.DataFrame:
    """Reads UZIO license report, extracting exact headers while bypassing corrupt metadata."""
    try:
        df = pd.read_excel(file, header=None, dtype=str)
        # Find the header row by looking for 'Employee ID'
        header_idx = -1
        for i, row in df.head(20).iterrows():
            if any(str(c).strip() == 'Employee ID' for c in row.values if pd.notna(c)):
                header_idx = i
                break
        
        if header_idx != -1:
            df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.iloc[header_idx])]
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
            df = df.replace('None', '')
            df.dropna(how='all', inplace=True)
            return df
        else:
            # Fallback if not found in first 20 rows
            file.seek(0)
            df = pd.read_excel(file, dtype=str)
            return df
    except Exception as e:
        st.error(f"Could not read Uzio file: {e}")
        return None

def read_adp_license(file) -> pd.DataFrame:
    """Reads ADP license report, extracting headers while bypassing metadata."""
    is_csv = getattr(file, "name", "").lower().endswith(".csv")

    def _read(header):
        if is_csv:
            try:
                return pd.read_csv(file, header=header, dtype=str)
            except UnicodeDecodeError:
                file.seek(0)
                return pd.read_csv(file, header=header, dtype=str, encoding="latin1")
        return pd.read_excel(file, header=header, dtype=str)

    try:
        df = _read(None)
        # Locate header row
        header_idx = -1
        for i, row in df.head(20).iterrows():
            if any(str(c).strip() == 'Associate ID' for c in row.values if pd.notna(c)):
                header_idx = i
                break

        if header_idx != -1:
            df.columns = [str(c).strip() if pd.notna(c) else f"Unnamed_{i}" for i, c in enumerate(df.iloc[header_idx])]
            df = df.iloc[header_idx + 1:].reset_index(drop=True)
            df = df.replace('None', '')
            df.dropna(how='all', inplace=True)
            return df
        else:
            # Fallback
            file.seek(0)
            df = _read(0)
            return df
    except Exception as e:
        st.error(f"Could not read ADP file: {e}")
        return None

# --- AUDIT LOGIC ---
def run_license_audit(uzio_df, adp_df):
    """
    Compares Uzio and ADP License data bidirectionally:
      - Uzio licenses are checked against ADP
      - ADP licenses are checked against Uzio
      Only employees with at least one license in either system are included.
    """
    UZIO_KEY = 'Employee ID'
    UZIO_NUM_COL = 'License Number'
    UZIO_DATE_COL = 'License Expiration Date'
    
    ADP_KEY = 'Associate ID'
    ADP_NUM_COL = 'License/Certification Code' if 'License/Certification Code' in adp_df.columns else 'License/Certification ID'
    ADP_DATE_COL = 'Expiration Date'
    
    # --- Detect name columns ---
    # Uzio: 'Full Name' column
    uzio_name_col = None
    for c in uzio_df.columns:
        if str(c).strip().lower() == 'full name':
            uzio_name_col = c
            break
    
    # ADP: 'Legal First Name' + 'Legal Last Name'
    adp_fname_col = None
    adp_lname_col = None
    for c in adp_df.columns:
        cl = str(c).strip().lower()
        if cl == 'legal first name':
            adp_fname_col = c
        elif cl == 'legal last name':
            adp_lname_col = c
    
    required_uzio = [UZIO_KEY]
    required_adp = [ADP_KEY]
    
    # Check for key columns; license columns are optional (we handle missing gracefully)
    missing_uzio = [c for c in required_uzio if c not in uzio_df.columns]
    if missing_uzio:
        st.error(f"Missing required Uzio columns: {', '.join(missing_uzio)}")
        return None
            
    missing_adp = [c for c in required_adp if c not in adp_df.columns]
    if missing_adp:
        st.error(f"Missing required ADP columns: {', '.join(missing_adp)}")
        return None
    
    # Check if license columns exist
    has_uzio_num = UZIO_NUM_COL in uzio_df.columns
    has_uzio_date = UZIO_DATE_COL in uzio_df.columns
    has_adp_num = ADP_NUM_COL in adp_df.columns
    has_adp_date = ADP_DATE_COL in adp_df.columns
    
    if not has_uzio_num and not has_adp_num:
        st.error("Neither Uzio nor ADP file has a license number column.")
        return None
            
    # Normalize Keys
    uzio_df[UZIO_KEY] = uzio_df[UZIO_KEY].apply(lambda x: str(x).strip() if norm_blank(x) != "" else "")
    adp_df[ADP_KEY] = adp_df[ADP_KEY].apply(lambda x: str(x).strip() if norm_blank(x) != "" else "")
    
    # --- Helper to get clean string value ---
    def clean_val(val):
        v = norm_blank(val)
        if v == "":
            return ""
        s = str(v).strip()
        if s.lower() in ('nan', 'none', 'nat'):
            return ""
        return s
    
    # --- Helper to get employee name ---
    def get_uzio_name(rec):
        if uzio_name_col and uzio_name_col in rec:
            return clean_val(rec[uzio_name_col])
        return ""
    
    def get_adp_name(rec):
        fname = clean_val(rec.get(adp_fname_col, "")) if adp_fname_col else ""
        lname = clean_val(rec.get(adp_lname_col, "")) if adp_lname_col else ""
        return f"{fname} {lname}".strip()
    
    # Pre-process into maps
    uzio_map = {}
    for r in uzio_df.to_dict('records'):
        k = r[UZIO_KEY]
        if k:
            if k not in uzio_map:
                uzio_map[k] = []
            uzio_map[k].append(r)
            
    adp_map = {}
    for r in adp_df.to_dict('records'):
        k = r[ADP_KEY]
        if k:
            if k not in adp_map:
                adp_map[k] = []
            adp_map[k].append(r)
    
    rows = []
    processed_adp_pairs = set()  # Track (eid, license_num) pairs processed from ADP side
    
    # =============================================
    # PASS 1: Iterate Uzio licenses -> check ADP
    # =============================================
    for eid, uzio_recs in uzio_map.items():
        for uz_rec in uzio_recs:
            uz_num = clean_val(uz_rec.get(UZIO_NUM_COL, "")) if has_uzio_num else ""
            uz_date_raw = uz_rec.get(UZIO_DATE_COL, "") if has_uzio_date else ""
            uz_date = try_parse_date(uz_date_raw)
            emp_name = get_uzio_name(uz_rec)
            
            # Skip employees with no license in Uzio
            if not uz_num:
                continue
            
            adp_licenses = adp_map.get(eid, [])
            
            # Try to find matching license in ADP
            match_found = False
            matched_adp_rec = None
            
            for adp_rec in adp_licenses:
                adp_num = clean_val(adp_rec.get(ADP_NUM_COL, "")) if has_adp_num else ""
                if adp_num and uz_num.lower() == adp_num.lower():
                    match_found = True
                    matched_adp_rec = adp_rec
                    processed_adp_pairs.add((eid, adp_num.lower()))
                    break
            
            if not emp_name and matched_adp_rec:
                emp_name = get_adp_name(matched_adp_rec)
            
            if match_found:
                # License Number matched
                adp_num_val = clean_val(matched_adp_rec.get(ADP_NUM_COL, ""))
                rows.append({
                    "Employee ID": eid,
                    "Employee Name": emp_name,
                    "Field": "License Number",
                    "Status": "Data Match",
                    "Uzio Value": uz_num,
                    "ADP Value": adp_num_val
                })
                
                # Check Expiration Date
                adp_date_raw = matched_adp_rec.get(ADP_DATE_COL, "") if has_adp_date else ""
                adp_date = try_parse_date(adp_date_raw)
                
                if uz_date == adp_date:
                    date_status = "Data Match"
                elif uz_date == "" and adp_date != "":
                    date_status = "Value missing in Uzio (ADP has value)"
                elif uz_date != "" and adp_date == "":
                    date_status = "Value missing in ADP (Uzio has value)"
                else:
                    date_status = "Data Mismatch"
                    
                rows.append({
                    "Employee ID": eid,
                    "Employee Name": emp_name,
                    "Field": "Expiration Date",
                    "Status": date_status,
                    "Uzio Value": uz_date,
                    "ADP Value": adp_date
                })
            else:
                # Not found in ADP
                if not emp_name and adp_licenses:
                    emp_name = get_adp_name(adp_licenses[0])
                rows.append({
                    "Employee ID": eid,
                    "Employee Name": emp_name,
                    "Field": "License Number",
                    "Status": "Missing in ADP",
                    "Uzio Value": uz_num,
                    "ADP Value": ""
                })
    
    # =============================================
    # PASS 2: Iterate ADP licenses -> find ones missing from Uzio
    # =============================================
    for eid, adp_recs in adp_map.items():
        for adp_rec in adp_recs:
            adp_num = clean_val(adp_rec.get(ADP_NUM_COL, "")) if has_adp_num else ""
            
            # Skip blank licenses
            if not adp_num:
                continue
            
            # Skip if already matched in Pass 1
            if (eid, adp_num.lower()) in processed_adp_pairs:
                continue
            
            emp_name = get_adp_name(adp_rec)
            
            # Check if Uzio has this employee at all
            uzio_recs = uzio_map.get(eid, [])
            if not emp_name and uzio_recs:
                emp_name = get_uzio_name(uzio_recs[0])
            
            adp_date_raw = adp_rec.get(ADP_DATE_COL, "") if has_adp_date else ""
            adp_date = try_parse_date(adp_date_raw)
            
            if eid not in uzio_map:
                status = "Employee Not Found in Uzio"
            else:
                status = "Missing in Uzio"
            
            rows.append({
                "Employee ID": eid,
                "Employee Name": emp_name,
                "Field": "License Number",
                "Status": status,
                "Uzio Value": "",
                "ADP Value": adp_num
            })
            
            # Also flag the date
            if adp_date:
                rows.append({
                    "Employee ID": eid,
                    "Employee Name": emp_name,
                    "Field": "Expiration Date",
                    "Status": status,
                    "Uzio Value": "",
                    "ADP Value": adp_date
                })
    
    result_df = pd.DataFrame(rows)
    
    # Final NaN cleanup
    if not result_df.empty:
        result_df = result_df.fillna("")
    
    return result_df

# --- UI RENDER FLOW ---
def render_ui():
    st.title("ADP License Details Audit Tool")
    st.write("Compare license numbers and expiration dates between ADP and Uzio. Output will be generated as a single sheet.")

    client_name = st.text_input("Client Name", key="adp_license_client_name")

    st.markdown("### Step 1: Upload Files")
    col1, col2 = st.columns(2)
    with col1:
        uzio_file = st.file_uploader("Upload UZIO License Report (.xlsx)", type=['xlsx', 'csv'], key='uzio_license_upload')
    with col2:
        adp_file = st.file_uploader("Upload ADP License Report (.xlsx)", type=['xlsx', 'csv'], key='adp_license_upload')

    if uzio_file and adp_file:
        try:
            with st.spinner("Extracting Licenses..."):
                uzio_df = read_uzio_license(uzio_file)
                adp_df = read_adp_license(adp_file)

            if uzio_df is None or adp_df is None:
                st.error("Failed to parse one or both files.")
                return

            submit_audit = st.button("Run License Audit", type="primary")

            if submit_audit:
                 if not client_name:
                     st.warning("Please enter a Client Name before running the audit.")
                     return

                 with st.spinner("Running Audit & Generating Match Report..."):
                     result_df = run_license_audit(uzio_df, adp_df)
                     
                 if result_df is not None:
                      st.success("Audit Complete!")
                      
                      st.markdown("### Audit Summary")
                      
                      col1, col2, col3, col4, col5 = st.columns(5)
                      counts = result_df['Status'].value_counts().to_dict()
                      
                      col1.metric("Total Match", counts.get("Data Match", 0))
                      col2.metric("Total Mismatch", counts.get("Data Mismatch", 0))
                      col3.metric("Missing in UZIO", counts.get("Missing in Uzio", 0) + counts.get("Employee Not Found in Uzio", 0))
                      col4.metric("Missing in ADP", counts.get("Missing in ADP", 0))
                      col5.metric("Value Missing", counts.get("Value missing in Uzio (ADP has value)", 0) + counts.get("Value missing in ADP (Uzio has value)", 0))

                      st.dataframe(result_df)

                      # Download Button
                      output_filename = f"{client_name.strip()}_License_Audit_Report_{datetime.now().strftime('%d_%m_%Y')}.xlsx"
                      
                      buffer = io.BytesIO()
                      with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                          result_df.to_excel(writer, sheet_name='License Audit Results', index=False)

                      st.download_button(
                          label="Download Excel Report",
                          data=buffer.getvalue(),
                          file_name=output_filename,
                          mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                      )
                      
        except Exception as e:
            st.error(f"Error processing files: {e}")
