import streamlit as st
import pandas as pd
import io
import re
from datetime import date
from utils.audit_utils import norm_col, clean_money_val

# =========================================================
# Paycom to Uzio Deduction Audit Tool
# =========================================================

APP_TITLE = "Paycom to Uzio Deduction Audit Tool"

def norm_str(x):
    """Normalize string, handle None/NaN."""
    if pd.isna(x) or x is None:
        return ""
    return str(x).strip()

def norm_id(x):
    """Normalize Employee ID (remove leading zeros, strip)."""
    s = norm_str(x)
    return s.lstrip("0")

def read_df_flexible(file, required_columns, fallback_columns=None):
    """
    Read a file (CSV or Excel) and find the header row containing required_columns.
    Returns a DataFrame or raises ValueError.
    """
    file.seek(0)
    file_bytes = file.getvalue()
    
    # Try Excel
    try:
        # We use pd.ExcelFile to avoid re-reading the file multiple times for multiple sheets
        xls = pd.ExcelFile(io.BytesIO(file_bytes))
        for sheet in xls.sheet_names:
            # Read first 50 rows to find header
            df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=50)
            
            header_row_idx = None
            for idx, row in df_raw.iterrows():
                row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                if all(any(col.lower() in v for v in row_vals) for col in required_columns):
                    header_row_idx = idx
                    break
            
            if header_row_idx is not None:
                df = pd.read_excel(xls, sheet_name=sheet, header=header_row_idx, dtype=str)
                df.columns = [norm_col(c) for c in df.columns]
                return df
                
        # Fallback columns if provided
        if fallback_columns:
            for sheet in xls.sheet_names:
                df_raw = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=50)
                for idx, row in df_raw.iterrows():
                    row_vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                    if all(any(col.lower() in v for v in row_vals) for col in fallback_columns):
                        df = pd.read_excel(xls, sheet_name=sheet, header=idx, dtype=str)
                        df.columns = [norm_col(c) for c in df.columns]
                        return df
    except Exception:
        pass
        
    # Try CSV
    try:
        header_row_idx = None
        file.seek(0)
        wrapper = io.TextIOWrapper(io.BytesIO(file_bytes), encoding='utf-8', errors='replace')
        for i, line in enumerate(wrapper):
            line_lower = line.lower()
            if all(col.lower() in line_lower for col in required_columns):
                header_row_idx = i
                break
            if i > 100: break
            
        if header_row_idx is not None:
            file.seek(0)
            df = pd.read_csv(io.BytesIO(file_bytes), header=header_row_idx, dtype=str)
            df.columns = [norm_col(c) for c in df.columns]
            return df
            
        # Fallback for CSV
        if fallback_columns:
            file.seek(0)
            wrapper = io.TextIOWrapper(io.BytesIO(file_bytes), encoding='utf-8', errors='replace')
            for i, line in enumerate(wrapper):
                line_lower = line.lower()
                if all(col.lower() in line_lower for col in fallback_columns):
                    header_row_idx = i
                    break
                if i > 100: break
            
            if header_row_idx is not None:
                file.seek(0)
                df = pd.read_csv(io.BytesIO(file_bytes), header=header_row_idx, dtype=str)
                df.columns = [norm_col(c) for c in df.columns]
                return df
    except Exception:
        pass

    raise ValueError(f"Could not find header containing {required_columns} in Excel or CSV.")

def read_uzio_deduction(file):
    """Read Uzio Deduction Export (Excel or CSV)."""
    return read_df_flexible(file, ["employee id", "deduction name"], fallback_columns=["employee id"])

def read_paycom_deduction(file):
    """Read Paycom Deduction Export (Excel or CSV)."""
    return read_df_flexible(file, ["code", "amount"])


def run_audit(file_uzio, file_paycom, UI_MAPPING):
    # 1. Load Data
    
    # Uzio
    try:
        df_uzio = read_uzio_deduction(file_uzio)
    except Exception as e:
        return None, f"Error reading Uzio file: {e}", []

    # Paycom
    try:
        df_paycom = read_paycom_deduction(file_paycom)
    except Exception as e:
        return None, f"Error reading Paycom file: {e}", []


    # 2. Process Mapping
    mapping = {k.lower(): v for k, v in UI_MAPPING.items()}
    mapping.update(UI_MAPPING) # include original case too

    # 3. Process Paycom
    # Columns: EE Code, EE Name, Deduction Code, Deduction Desc, Amount, Percent
    p_id_col = next((c for c in df_paycom.columns if any(x in c.lower() for x in ["ee code", "employee code", "employee id"])), "EE Code")
    p_code_col = next((c for c in df_paycom.columns if "deduction code" in c.lower()), next((c for c in df_paycom.columns if "code" in c.lower() and "employee" not in c.lower() and "ee" not in c.lower()), "Code"))
    p_desc_col = next((c for c in df_paycom.columns if "deduction desc" in c.lower()), next((c for c in df_paycom.columns if "description" in c.lower()), "Description"))
    
    p_amt_col = next((c for c in df_paycom.columns if "amount" in c.lower() and "exempt" not in c.lower()), "Amount")
    p_rate_col = next((c for c in df_paycom.columns if any(x in c.lower() for x in ["percent", "rate"])), "Rate")
    
    paycom_data = []
    
    for _, row in df_paycom.iterrows():
        emp_id = norm_id(row.get(p_id_col))
        if not emp_id: continue
        
        raw_code = norm_str(row.get(p_code_col))
        raw_desc = norm_str(row.get(p_desc_col))
        
        # Map to Uzio Name
        # Try Description first, then Code
        ded_name = mapping.get(raw_desc.lower())
        if not ded_name:
             ded_name = mapping.get(raw_desc)
             
        if not ded_name:
             ded_name = mapping.get(raw_code.lower())
             
        if not ded_name:
             continue
             
        amt = clean_money_val(row.get(p_amt_col))
        rate = clean_money_val(row.get(p_rate_col))
        
        # Use Rate if Amount is 0?
        # Often deductions like 401k use Rate (Percentage).
        # We'll store both, but comparison usually checks amount if non-zero, else rate?
        # Actually Uzio usually has "Employee Amount" and sometimes "Percentage".
        # Let's sum Amount.
        
        paycom_data.append({
            "ID": emp_id,
            "Deduction": ded_name,
            "Amount": amt,
            "Rate": rate,
            "Code": raw_code,
            "Key": f"{emp_id}|{ded_name}".lower()
        })
        
    df_p_clean = pd.DataFrame(paycom_data)
    if not df_p_clean.empty:
        # Sum duplicates?
        df_p_clean = df_p_clean.groupby(["ID", "Deduction", "Key"], as_index=False).agg({
            "Amount": "sum",
            "Rate": "max", # Rate usually constant
            "Code": "first"
        })
    else:
         df_p_clean = pd.DataFrame(columns=["ID", "Deduction", "Key", "Amount", "Rate", "Code"])

    # 4. Process Uzio
    # Columns expected: Employee Id, Deduction Name, Employee Amount
    u_id_col = next((c for c in df_uzio.columns if "employee id" in c.lower()), None)
    u_ded_col = next((c for c in df_uzio.columns if "deduction name" in c.lower()), None)
    u_amt_col = next((c for c in df_uzio.columns if "employee amount" in c.lower()), None)
    if not u_amt_col:
        u_amt_col = next((c for c in df_uzio.columns if "amount" in c.lower()), None)
        
    if not u_id_col or not u_ded_col:
         return None, f"Uzio file missing required columns (Employee Id, Deduction Name). Found: {list(df_uzio.columns)}", []

    uzio_data = []
    
    for _, row in df_uzio.iterrows():
        emp_id = norm_id(row.get(u_id_col))
        if not emp_id: continue
        
        ded_name = norm_str(row.get(u_ded_col))
        amt = clean_money_val(row.get(u_amt_col))
        
        uzio_data.append({
            "ID": emp_id,
            "Deduction": ded_name,
            "Amount": amt,
            "Key": f"{emp_id}|{ded_name}".lower()
        })
        
    df_u_clean = pd.DataFrame(uzio_data)
    if not df_u_clean.empty:
         df_u_clean = df_u_clean.groupby(["ID", "Deduction", "Key"], as_index=False)["Amount"].sum()
    else:
         df_u_clean = pd.DataFrame(columns=["ID", "Deduction", "Key", "Amount"])

    # 5. Merge and Compare
    merged = pd.merge(df_p_clean, df_u_clean, on="Key", how="outer", suffixes=("_P", "_U"))
    
    results = []
    
    for _, row in merged.iterrows():
        emp_id = row["ID_P"] if pd.notna(row["ID_P"]) else row["ID_U"]
        ded_name = row["Deduction_P"] if pd.notna(row["Deduction_P"]) else row["Deduction_U"]
        
        p_amt = row["Amount_P"] if pd.notna(row["Amount_P"]) else 0.0
        p_rate = row["Rate"] if pd.notna(row["Rate"]) else 0.0
        u_amt = row["Amount_U"] if pd.notna(row["Amount_U"]) else 0.0
        
        in_p = pd.notna(row["ID_P"])
        in_u = pd.notna(row["ID_U"])
        
        status = ""
        
        if in_p and in_u:
            # Compare
            # Logic: If Paycom Amount is 0, maybe check Rate vs Amount?
            # Or just compare Amounts.
            diff = abs(p_amt - u_amt)
            if diff < 0.01:
                status = "Data Match"
            elif p_amt == 0.0 and abs(p_rate - u_amt) < 0.01:
                 # Check if paycom rate matches uzio amount (common for percentages)
                 status = "Data Match"
            elif p_amt == 0.0 and abs(p_rate * 100 - u_amt) < 0.01:
                 # Handle decimal percentage formats (e.g. 0.05 rate vs 5.0 amount)
                 status = "Data Match"
            else:
                # Check rate mismatch? E.g. 0.04 vs 4.0?
                # User had issue with % mismatch.
                status = "Data Mismatch"
        elif in_p and not in_u:
             status = "Value missing in Uzio (Paycom has value)"
        elif in_u and not in_p:
             status = "Value missing in Paycom (Uzio has value)"
             
        results.append({
            "Employee ID": emp_id,
            "Deduction Name": ded_name,
            "Paycom Code": row["Code"] if pd.notna(row["Code"]) else "",
            "Paycom Amount": p_amt,
            "Paycom Rate": p_rate,
            "Uzio Amount": u_amt,
            "Status": status
        })
        
    # Generate Output
    df_res = pd.DataFrame(results)
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_res.to_excel(writer, sheet_name='Audit Details', index=False)
        
        # Summary
        if not df_res.empty:
             summary = df_res.groupby(["Status"]).size().reset_index(name="Count")
             summary.to_excel(writer, sheet_name='Summary', index=False)
             
    return output.getvalue(), None, results

def get_unique_uzio_deductions(file):
    try:
        # Seek to start
        file.seek(0)
        df_uzio = read_uzio_deduction(file)
        
        u_ded_col = next((c for c in df_uzio.columns if "deduction name" in c.lower()), None)
        if not u_ded_col: return []

        unique_deductions = df_uzio[u_ded_col].dropna().unique().tolist()
        return [str(d).strip() for d in unique_deductions if str(d).strip() != ""]
    except Exception as e:
        return []

def get_unique_paycom_deductions(file):
    try:
        df_paycom = read_paycom_deduction(file)
        
        p_desc_col = next((c for c in df_paycom.columns if "deduction desc" in c.lower()), next((c for c in df_paycom.columns if "description" in c.lower()), None))

        
        if not p_desc_col:
             p_desc_col = next((c for c in df_paycom.columns if "deduction code" in c.lower()), next((c for c in df_paycom.columns if "code" in c.lower() and "employee" not in c.lower() and "ee" not in c.lower()), None))

        if not p_desc_col: return []

        unique_deductions = df_paycom[p_desc_col].dropna().unique().tolist()
        
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
    st.title(APP_TITLE)
    st.markdown("""
    **Instructions**:
    1. Upload **Uzio Deduction Export** (Excel or CSV).
    2. Upload **Paycom Deduction Export** (Excel or CSV).
    3. Map the extracted Paycom deductions to Uzio deductions, then click **Run Comparison**.
    """)

    
    col1, col2 = st.columns(2)
    with col1:
        u_file = st.file_uploader("Uzio Deduction File", type=["xlsx", "xls", "csv"], key="pd_u")
    with col2:
        p_file = st.file_uploader("Paycom Deduction File", type=["xlsx", "xls", "csv"], key="pd_p")

        
    client_name = st.text_input("Client Name", value="Client", key="paycom_deduction_client")
    
    if u_file and p_file:
         st.markdown("---")
         st.subheader("Map Deductions")
         
         uzio_deductions = get_unique_uzio_deductions(u_file)
         paycom_deductions = get_unique_paycom_deductions(p_file)
         
         if not uzio_deductions:
              st.error("Could not find any 'Deduction Name' values in the Uzio file.")
         elif not paycom_deductions:
              st.error("Could not find any Deduction Descriptions or Codes in the Paycom file.")
         else:
              st.markdown("Please map the Paycom Deductions to the corresponding Uzio Deductions below:")
              
              ui_mapping = {}
              
              # Initialize session state for mappings
              for p_ded in paycom_deductions:
                  key = f"map_p_{p_ded}"
                  if key not in st.session_state:
                      default_val = "— Ignore / Skip —"
                      for opt in uzio_deductions:
                          if opt.lower() == p_ded.lower():
                              default_val = opt
                              break
                      st.session_state[key] = default_val

              # Collect all currently selected values
              selected_values = set()
              for p_ded in paycom_deductions:
                  val = st.session_state.get(f"map_p_{p_ded}", "— Ignore / Skip —")
                  if val != "— Ignore / Skip —":
                      selected_values.add(val)
              
              # Render mapping UI
              for p_ded in sorted(paycom_deductions):
                  col_a, col_b = st.columns([1, 1])
                  with col_a:
                       st.write(p_ded)
                  with col_b:
                       key = f"map_p_{p_ded}"
                       current_val = st.session_state.get(key, "— Ignore / Skip —")
                       
                       available_options = ["— Ignore / Skip —"]
                       for opt in sorted(uzio_deductions):
                           if opt == current_val or opt not in selected_values:
                               available_options.append(opt)
                               
                       selected = st.selectbox(
                           f"Map for {p_ded}", 
                           available_options,
                           key=key,
                           label_visibility="collapsed"
                       )
                       if selected != "— Ignore / Skip —":
                            ui_mapping[p_ded] = selected
              
              st.markdown("---")
              if st.button("Run Comparison", type="primary"):
                  with st.spinner("Processing..."):
                      u_file.seek(0)
                      p_file.seek(0)
                      report, err, _ = run_audit(u_file, p_file, ui_mapping)
                      
                      if err:
                          st.error(err)
                      else:
                          st.success("Audit Complete!")
                          timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
                          filename = f"{client_name}_Uzio_Paycom_Deduction_Audit_Report_{timestamp}.xlsx"
          
                          st.download_button(
                              "Download Report",
                              data=report,
                              file_name=filename,
                              mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                          )

