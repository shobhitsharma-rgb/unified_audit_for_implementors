import streamlit as st
import pandas as pd
import io
import re
from datetime import datetime
from utils.audit_utils import get_identity_match_map, norm_ssn_canonical

# =========================================================
# ADP to Uzio Deduction Audit Tool
# INPUT: One Excel File with 3 Tabs:
#   1. Uzio Data
#   2. ADP Data
#   3. Mapping Sheet
# =========================================================

def norm_col(c):
    """Normalize column names to be case-insensitive and stripped."""
    if c is None: return ""
    return str(c).strip().replace("\n", " ").strip()

def clean_money_val(x):
    """Parse money/percentage strings to float. Returns original string if not a number."""
    if pd.isna(x) or x == "":
        return 0.0
    s = str(x).strip()
    s_clean = s.replace("$", "").replace("%", "").replace(",", "")
    s_clean = s_clean.replace("(", "-").replace(")", "") # Handle accounting negative
    try:
        return float(s_clean)
    except:
        # If it's not a number (like an SSN), return the string itself for comparison
        return s

def read_uzio_deduction(file):
    """
    Read Uzio Deduction Export.
    Search all sheets for header row containing 'Employee Id' and 'Deduction Name'.
    """
    xls = pd.ExcelFile(io.BytesIO(file.getvalue()), engine='openpyxl')
    
    for sheet in xls.sheet_names:
        # Read first 20 rows
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=20)
        
        header_row_idx = None
        for idx, row in df_raw.iterrows():
            row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
            # Strict check: Must have Employee Id AND Deduction Name
            if any("employee id" in v for v in row_vals) and any("deduction name" in v for v in row_vals):
                header_row_idx = idx
                break
        
        if header_row_idx is not None:
             # Found it!
             df = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx, dtype=str)
             # Normalize columns
             df.columns = [norm_col(c) for c in df.columns]
             return df

    # Fallback if strict check fails: Try just Employee Id
    for sheet in xls.sheet_names:
        df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=20)
        for idx, row in df_raw.iterrows():
             row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
             if any("employee id" in v for v in row_vals):
                  df = pd.read_excel(xls, sheet_name=sheet, header=idx, dtype=str)
                  df.columns = [norm_col(c) for c in df.columns]
                  return df
                  
    raise ValueError("Could not find 'Employee Id' column in any sheet.")

def run_audit(file_uzio, file_adp, UI_MAPPING):
    # 1. Load Data
    
    # Uzio Data File
    try:
        df_uzio = read_uzio_deduction(file_uzio)
    except Exception as e:
        return None, f"Error reading Uzio Data File: {e}", []

    # ADP Data File
    try:
        xls_adp = pd.ExcelFile(io.BytesIO(file_adp.getvalue()), engine='openpyxl')
        
        # Determine ADP header row
        adp_sheet = xls_adp.sheet_names[0]
        # Peek at first few rows to find "EMPLOYEE NAME"
        peek_df = pd.read_excel(xls_adp, sheet_name=adp_sheet, nrows=20, header=None)
        
        header_row_idx = 0
        for idx, row in peek_df.iterrows():
            row_str = " ".join([str(val).upper() for val in row.values])
            if "EMPLOYEE NAME" in row_str or "ASSOCIATE ID" in row_str:
                header_row_idx = idx
                break
                
        df_adp = pd.read_excel(xls_adp, sheet_name=adp_sheet, header=header_row_idx, dtype=str)
    except Exception as e:
        return None, f"Error reading ADP Data File: {e}", []

    return _run_deduction_audit(df_uzio, df_adp, UI_MAPPING)


def _run_deduction_audit(df_uzio, df_adp, UI_MAPPING):
    # Normalize Columns
    df_uzio.columns = [norm_col(c) for c in df_uzio.columns]
    df_adp.columns = [norm_col(c) for c in df_adp.columns]

    # Process Mapping
    mapping = {k.lower(): v for k, v in UI_MAPPING.items()}
    mapping.update(UI_MAPPING)

    # Required Cols
    adp_id_col = next((c for c in df_adp.columns if "associate" in c.lower() and "id" in c.lower()), None)
    adp_code_col = next((c for c in df_adp.columns if "deduction" in c.lower() and "code" in c.lower()), None)
    adp_amt_col = next((c for c in df_adp.columns if "amount" in c.lower() or "rate" in c.lower()), None)
    adp_desc_col = next((c for c in df_adp.columns if "deduction" in c.lower() and "description" in c.lower()), None)
    adp_pct_col = next((c for c in df_adp.columns if "deduction" in c.lower() and "%" in c.lower()), None)
    adp_ssn_col = next((c for c in df_adp.columns if "ssn" in c.lower() or "tax id" in c.lower()), None)

    # Process Uzio Columns
    uz_id_col = next((c for c in df_uzio.columns if "employee" in c.lower() and "id" in c.lower()), None)
    uz_ded_col = next((c for c in df_uzio.columns if "deduction" in c.lower() and "name" in c.lower()), None)
    uz_amt_col = next((c for c in df_uzio.columns if "amount" in c.lower() or "percent" in c.lower()), None)
    uz_ssn_col = next((c for c in df_uzio.columns if "ssn" in c.lower()), None)

    if not all([adp_id_col, adp_code_col, adp_amt_col]):
        return None, f"ADP Sheet missing required columns (Associate ID, Deduction Code, Deduction Amount). Found: {list(df_adp.columns)}", []

    if not all([uz_id_col, uz_ded_col, uz_amt_col]):
        return None, f"Uzio Sheet missing required columns (Employee ID, Deduction Name, Amount/Percentage). Found: {list(df_uzio.columns)}", []

    # 1. Resolve Identity Match Map (UZIO_ID -> ADP_ID)
    uz_to_adp_id_map = {}
    if uz_ssn_col and adp_ssn_col:
        uz_to_adp_id_map = get_identity_match_map(
            df_uzio, df_adp, 
            uzio_id_col=uz_id_col, 
            vendor_id_col=adp_id_col,
            uzio_ssn_col=uz_ssn_col,
            vendor_ssn_col=adp_ssn_col
        )
    # Reverse map for ADP -> Uzio lookup
    adp_to_uz_id_map = {v: k for k, v in uz_to_adp_id_map.items()}

    adp_records = []
    for _, row in df_adp.iterrows():
        emp_id = str(row[adp_id_col]).strip()
        raw_code = str(row[adp_code_col]).strip()
        raw_desc = str(row[adp_desc_col]).strip() if adp_desc_col else ""
        
        deduction_name = None
        if raw_desc:
            deduction_name = mapping.get(raw_desc, mapping.get(raw_desc.lower()))
        if not deduction_name and raw_code:
            deduction_name = mapping.get(raw_code, mapping.get(raw_code.lower()))
            
        if not deduction_name:
            continue
        
        amt = clean_money_val(row[adp_amt_col])
        if amt == 0.0 and adp_pct_col:
            pct_val = clean_money_val(row[adp_pct_col])
            if pct_val != 0.0:
                amt = pct_val
        
        # Normalize for matching
        match_id = adp_to_uz_id_map.get(emp_id, emp_id)
        
        adp_records.append({
            "Employee_ID": emp_id,
            "Deduction_Name": deduction_name,
            "ADP_Raw_Code": raw_code,
            "ADP_Description": raw_desc,
            "ADP_Amount": amt,
            "Key": f"{match_id}|{deduction_name}".lower()
        })
    
    df_adp_clean = pd.DataFrame(adp_records)
    if not df_adp_clean.empty:
        df_adp_clean = df_adp_clean.groupby(["Employee_ID", "Deduction_Name", "ADP_Raw_Code", "ADP_Description", "Key"], as_index=False)["ADP_Amount"].sum()
    else:
        df_adp_clean = pd.DataFrame(columns=["Employee_ID", "Deduction_Name", "ADP_Raw_Code", "ADP_Description", "Key", "ADP_Amount"])

    uzio_records = []
    for _, row in df_uzio.iterrows():
        emp_id = str(row[uz_id_col]).strip()
        ded_name = str(row[uz_ded_col]).strip()
        amt = clean_money_val(row[uz_amt_col])
        
        uzio_records.append({
            "Uzio_Employee_ID": emp_id,
            "Uzio_Deduction_Name": ded_name,
            "Uzio_Amount": amt,
            "Key": f"{emp_id}|{ded_name}".lower()
        })
    
    df_uz_clean = pd.DataFrame(uzio_records)
    if not df_uz_clean.empty:
        df_uz_clean = df_uz_clean.groupby(["Uzio_Employee_ID", "Uzio_Deduction_Name", "Key"], as_index=False)["Uzio_Amount"].sum()
    else:
        df_uz_clean = pd.DataFrame(columns=["Uzio_Employee_ID", "Uzio_Deduction_Name", "Key", "Uzio_Amount"])

    # Merge
    merged = pd.merge(df_adp_clean, df_uz_clean, on="Key", how="outer", suffixes=('_ADP', '_UZIO'))
    
    # IDs lists
    adp_emps = set(df_adp_clean["Employee_ID"].unique()) if not df_adp_clean.empty else set()
    uzio_emps = set(df_uz_clean["Uzio_Employee_ID"].unique()) if not df_uz_clean.empty else set()
    
    results = []
    for _, row in merged.iterrows():
        adp_id = row["Employee_ID"] if pd.notna(row["Employee_ID"]) else ""
        uz_id = row["Uzio_Employee_ID"] if pd.notna(row["Uzio_Employee_ID"]) else ""
        
        # Display ID: Use Uzio ID if possible
        display_id = uz_id if uz_id else adp_id
        
        adp_final_name = row["ADP_Description"] if pd.notna(row["ADP_Amount"]) and pd.notna(row["ADP_Description"]) else (row["ADP_Raw_Code"] if pd.notna(row["ADP_Amount"]) else "Not Available")
        uzio_final_name = row["Uzio_Deduction_Name"] if pd.notna(row["Uzio_Amount"]) else "Not Available"
        
        raw_code = row["ADP_Raw_Code"] if pd.notna(row["ADP_Raw_Code"]) else ""
        adp_val = row["ADP_Amount"] if pd.notna(row["ADP_Amount"]) else 0.0
        uz_val = row["Uzio_Amount"] if pd.notna(row["Uzio_Amount"]) else 0.0
        
        has_adp = pd.notna(row["ADP_Amount"])
        has_uzio = pd.notna(row["Uzio_Amount"])
        
        status = ""
        if has_adp and has_uzio:
            if abs(adp_val - uz_val) < 0.01:
                status = "Data Match"
            else:
                status = "Data Mismatch"
        elif has_adp and not has_uzio:
            if adp_id in adp_to_uz_id_map and adp_to_uz_id_map[adp_id] in uzio_emps:
                 status = "Value missing in Uzio (ADP has value)"
            elif adp_id in uzio_emps:
                 status = "Value missing in Uzio (ADP has value)"
            else:
                status = "Employee ID Not Found in Uzio"
        elif has_uzio and not has_adp:
            if uz_id in uz_to_adp_id_map and uz_to_adp_id_map[uz_id] in adp_emps:
                 status = "Value missing in ADP (Uzio has value)"
            elif uz_id in adp_emps:
                 status = "Value missing in ADP (Uzio has value)"
            else:
                status = "Employee ID Not Found in ADP"
        
        # Flag ID mismatch specifically if relevant
        if has_adp and has_uzio and adp_id != uz_id:
            status += " (Identity matched via SSN)"

        results.append({
            "Employee ID": display_id,
            "ADP ID": adp_id,
            "Uzio ID": uz_id,
            "ADP Deduction Description": adp_final_name,
            "Uzio Deduction Name": uzio_final_name,
            "ADP Code": raw_code,
            "ADP Amount": adp_val,
            "Uzio Amount": uz_val,
            "Status": status
        })
        
    return _generate_output(results)

def _generate_output(results):
    df_res = pd.DataFrame(results)
    
    # Consolidate Field Logic for Deduction Audit
    def get_field_name(row):
        uz_name = row.get("Uzio Deduction Name", "Not Available")
        adp_name = row.get("ADP Deduction Description", "Not Available")
        
        if uz_name != "Not Available":
            return uz_name
        return adp_name

    df_res["Field"] = df_res.apply(get_field_name, axis=1)

    # Pivot Summary
    expected_statuses = [
        "Data Match", "Data Mismatch", 
        "Value missing in Uzio (ADP has value)", "Value missing in ADP (Uzio has value)", 
        "Employee ID Not Found in Uzio", "Employee ID Not Found in ADP",
        "Column Missing in ADP Sheet", "Column Missing in Uzio Sheet"
    ]
    
    if not df_res.empty:
        field_summary = df_res.groupby(["Field", "Status"]).size().unstack(fill_value=0)
    else:
        field_summary = pd.DataFrame()

    for col in expected_statuses:
        if col not in field_summary.columns:
            field_summary[col] = 0
            
    field_summary["Total"] = field_summary.sum(axis=1) if not field_summary.empty else 0
    
    # Reorder
    cols_order = ["Total"] + [c for c in expected_statuses if c in field_summary.columns] + [c for c in field_summary.columns if c not in expected_statuses and c != "Total"]
    field_summary = field_summary[cols_order]
    
    out_buffer = io.BytesIO()
    with pd.ExcelWriter(out_buffer, engine='openpyxl') as writer:
        summary_data = {
            "Total Records": [len(df_res)],
            "Matches": [len(df_res[df_res["Status"] == "Data Match"])] if not df_res.empty else [0],
            "Mismatches": [len(df_res[df_res["Status"] == "Data Mismatch"])] if not df_res.empty else [0],
            "Value Missing in Uzio": [len(df_res[df_res["Status"] == "Value missing in Uzio (ADP has value)"])] if not df_res.empty else [0],
            "Emp Missing in Uzio": [len(df_res[df_res["Status"] == "Employee ID Not Found in Uzio"])] if not df_res.empty else [0],
             "Value Missing in ADP": [len(df_res[df_res["Status"] == "Value missing in ADP (Uzio has value)"])] if not df_res.empty else [0],
            "Emp Missing in ADP": [len(df_res[df_res["Status"] == "Employee ID Not Found in ADP"])] if not df_res.empty else [0]
        }
        pd.DataFrame(summary_data).transpose().reset_index().rename(columns={"index": "Metric", 0: "Count"}).to_excel(writer, sheet_name="Summary", index=False)
        field_summary.to_excel(writer, sheet_name="Field_Summary_By_Status")
        df_res.drop(columns=["Field"], inplace=True)
        df_res.to_excel(writer, sheet_name="Audit Details", index=False)
    
    return out_buffer.getvalue(), None, []


def get_unique_uzio_deductions_from_excel(file):
    try:
        file.seek(0)
        df_uzio = read_uzio_deduction(file)
        
        u_ded_col = next((c for c in df_uzio.columns if "deduction name" in c.lower()), None)
        if not u_ded_col: return []

        unique_deductions = df_uzio[u_ded_col].dropna().unique().tolist()
        return [str(d).strip() for d in unique_deductions if str(d).strip() != ""]
    except Exception as e:
        return []

def get_unique_adp_deductions_from_excel(file):
    try:
        file.seek(0)
        xls_adp = pd.ExcelFile(io.BytesIO(file.getvalue()), engine='openpyxl')
        
        adp_sheet = xls_adp.sheet_names[0]
        peek_df = pd.read_excel(xls_adp, sheet_name=adp_sheet, nrows=20, header=None)
        
        header_row_idx = 0
        for idx, row in peek_df.iterrows():
            row_str = " ".join([str(val).upper() for val in row.values])
            if "EMPLOYEE NAME" in row_str or "DEDUCTION CODE" in row_str:
                header_row_idx = idx
                break
                
        df_adp = pd.read_excel(xls_adp, sheet_name=adp_sheet, header=header_row_idx, dtype=str)
        df_adp.columns = [norm_col(c) for c in df_adp.columns]
        
        adp_ded_desc_col = next((c for c in df_adp.columns if "deduction description" in c.lower()), None)
        if not adp_ded_desc_col:
            adp_ded_desc_col = next((c for c in df_adp.columns if "deduction code" in c.lower()), None)
            
        if not adp_ded_desc_col: return []

        unique_deductions = df_adp[adp_ded_desc_col].dropna().unique().tolist()
        
        filtered_deductions = []
        for d in unique_deductions:
             s = str(d).strip()
             if not s:
                 continue
             s_lower = s.lower()
             if "checking" in s_lower or "savings" in s_lower:
                 continue
             filtered_deductions.append(s)
             
        return filtered_deductions
    except Exception as e:
        return []

def render_ui():
    st.title("ADP to Uzio Deduction Audit Tool")
    st.markdown("""
    **Instructions**:
    1. Upload **Uzio Deduction Export** (Excel).
    2. Upload **ADP Voluntary Deduction Export** (Excel).
    3. Map the extracted ADP deductions to Uzio deductions, then click **Run Comparison**.
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        u_file = st.file_uploader("Upload Uzio Deduction File", type=["xlsx", "xls"], key="adp_ded_uzio")
    with col2:
        a_file = st.file_uploader("Upload ADP Deduction File", type=["xlsx", "xls"], key="adp_ded_adp")

    client_name = st.text_input("Enter Client Name (for Report Filename)", value="Client_Name")

    if u_file and a_file:
         st.markdown("---")
         st.subheader("Map Deductions")
         
         uzio_deductions = get_unique_uzio_deductions_from_excel(u_file)
         adp_deductions = get_unique_adp_deductions_from_excel(a_file)
         
         if not uzio_deductions:
              st.error("Could not find any 'Deduction Name' values in the Uzio file.")
         elif not adp_deductions:
              st.error("Could not find any Deduction Descriptions or Codes in the ADP file.")
         else:
              st.markdown("Please map the ADP Deductions to the corresponding Uzio Deductions below:")
              
              ui_mapping = {}
              
              # Initialize session state for mappings
              for a_ded in adp_deductions:
                  key = f"map_adp_{a_ded}"
                  if key not in st.session_state:
                      default_val = "— Ignore / Skip —"
                      for opt in uzio_deductions:
                          if opt.lower() == a_ded.lower():
                              default_val = opt
                              break
                      st.session_state[key] = default_val

              
              for a_ded in sorted(adp_deductions):
                  col_a, col_b = st.columns([1, 1])
                  with col_a:
                       st.write(a_ded)
                  with col_b:
                       key = f"map_adp_{a_ded}"
                       current_val = st.session_state.get(key, "— Ignore / Skip —")
                       
                       available_options = ["— Ignore / Skip —"] + sorted(uzio_deductions)
                               
                       selected = st.selectbox(
                           f"Map for {a_ded}", 
                           available_options,
                           key=key,
                           label_visibility="collapsed"
                       )
                       if selected != "— Ignore / Skip —":
                            ui_mapping[a_ded] = selected
              
              st.markdown("---")
              if st.button("Run Audit", type="primary"):
                  with st.spinner("Processing..."):
                      try:
                          u_file.seek(0)
                          a_file.seek(0)
                          report_data, error_msg, _ = run_audit(u_file, a_file, ui_mapping)
                          
                          if error_msg:
                              st.error(error_msg)
                          else:
                              st.success("Audit Completed Successfully!")
                              
                              timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
                              filename = f"{client_name}_Uzio_ADP_Deduction_Audit_Report_{timestamp}.xlsx"
                              
                              st.download_button(
                                  label="Download Audit Report",
                                  data=report_data,
                                  file_name=filename,
                                  mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                              )
                      except Exception as e:
                          st.error(f"An unexpected error occurred: {e}")
                          st.exception(e)

if __name__ == "__main__":
    st.set_page_config(page_title="ADP Deduction Audit", layout="wide")
    render_ui()
