import streamlit as st
import pandas as pd
import csv
import io

from utils.deduction_mapping import (
    load_deduction_mapping, unmapped_source_deductions, REQUIRED_COLUMNS,
)
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

def read_tabular(file, matcher, preview_rows=20):
    """Read an uploaded Excel OR CSV and return the sheet whose header row matches.

    `matcher` is given the lower-cased, non-empty cell values of a candidate row and
    says whether that row is the header. Excel is tried first (every sheet), then CSV.
    Returns None when nothing matches, so the caller can fall back or raise.
    """
    data = file.getvalue()
    try:
        xls = pd.ExcelFile(io.BytesIO(data), engine='openpyxl')
        for sheet in xls.sheet_names:
            peek = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=preview_rows)
            for idx, row in peek.iterrows():
                vals = [str(v).strip().lower() for v in row.values if pd.notna(v)]
                if matcher(vals):
                    df = pd.read_excel(xls, sheet_name=sheet, header=idx, dtype=str)
                    df.columns = [norm_col(c) for c in df.columns]
                    return df
    except Exception:
        pass   # not an Excel file (or no sheet matched) -- fall through to CSV

    try:
        text = data.decode("utf-8-sig", errors="replace")
        row_index = 0     # counts the rows pandas will see, so `header=` lines up
        for raw_row in csv.reader(io.StringIO(text)):
            if not raw_row:
                continue  # a wholly blank line: csv.reader yields it, pandas skips it
            vals = [str(v).strip().lower() for v in raw_row if str(v).strip()]
            if matcher(vals):
                df = pd.read_csv(io.BytesIO(data), header=row_index, dtype=str)
                df.columns = [norm_col(c) for c in df.columns]
                return df
            row_index += 1
            if row_index >= preview_rows:
                break
    except Exception:
        pass
    return None


def read_adp_deduction(file):
    """ADP Voluntary Deduction export, Excel or CSV. Falls back to the first row
    as the header when none of the expected labels is found."""
    tokens = ("employee name", "associate id", "deduction code", "deduction description")
    df = read_tabular(file, lambda vals: any(t in v for v in vals for t in tokens))
    if df is None:
        try:
            df = pd.read_csv(io.BytesIO(file.getvalue()), header=0, dtype=str)
        except Exception:
            df = pd.read_excel(io.BytesIO(file.getvalue()), header=0, dtype=str)
        df.columns = [norm_col(c) for c in df.columns]
    return df


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
    """Uzio Deduction Export, Excel or CSV.

    Prefers a header row carrying both 'Employee Id' and 'Deduction Name'; falls
    back to one carrying 'Employee Id' alone.
    """
    df = read_tabular(file, lambda vals: any("employee id" in v for v in vals)
                      and any("deduction name" in v for v in vals))
    if df is None:
        df = read_tabular(file, lambda vals: any("employee id" in v for v in vals))
    if df is None:
        raise ValueError("Could not find an 'Employee Id' column in any sheet of the Excel file, "
                         "or in the first rows of the CSV.")
    return df


def run_audit(file_uzio, file_adp, UI_MAPPING):
    # 1. Load Data
    
    # Uzio Data File
    try:
        df_uzio = read_uzio_deduction(file_uzio)
    except Exception as e:
        return None, f"Error reading Uzio Data File: {e}", []

    # ADP Data File
    try:
        df_adp = read_adp_deduction(file_adp)
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
        # Keep only Employee ID visible; hide the ADP ID / Uzio ID columns in the download.
        from openpyxl.utils import get_column_letter
        _ws_details = writer.sheets["Audit Details"]
        for _hide_col in ["ADP ID", "Uzio ID"]:
            if _hide_col in df_res.columns:
                _ci = df_res.columns.get_loc(_hide_col)
                _ws_details.column_dimensions[get_column_letter(_ci + 1)].hidden = True
    
    return out_buffer.getvalue(), None, []


# Kept for reference only: this used to populate the Uzio side of the manual
# mapping dropdowns, which the Employee Deduction Mapping upload replaced.
# Nothing calls it now.
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
        df_adp = read_adp_deduction(file)

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
    1. Upload **Uzio Deduction Export** (Excel or CSV).
    2. Upload **ADP Voluntary Deduction Export** (Excel or CSV).
    3. Upload the **Employee Deduction Mapping** CSV from the ADP Prior Payroll
       Setup Helper (`<Client>_EE_Deductions_mapping.csv`), then click **Run Audit**.
    """)
    
    col1, col2 = st.columns(2)
    with col1:
        u_file = st.file_uploader("Upload Uzio Deduction File", type=["xlsx", "xls", "csv"], key="adp_ded_uzio")
    with col2:
        a_file = st.file_uploader("Upload ADP Deduction File", type=["xlsx", "xls", "csv"], key="adp_ded_adp")

    m_file = st.file_uploader(
        "Upload Employee Deduction Mapping (from the Prior Payroll Setup Helper)",
        type=["csv", "xlsx", "xls"], key="adp_ded_mapping",
        help="The <Client>_EE_Deductions_mapping.csv the setup helper produces. "
             "Columns: " + ", ".join(REQUIRED_COLUMNS),
    )

    client_name = st.text_input("Enter Client Name (for Report Filename)", value="Client_Name")

    if u_file and a_file and m_file:
         st.markdown("---")
         st.subheader("Deduction Mapping")

         adp_deductions = get_unique_adp_deductions_from_excel(a_file)

         try:
             ui_mapping, mrep = load_deduction_mapping(m_file)
         except Exception as e:
             st.error(str(e))
             ui_mapping, mrep = {}, None

         if mrep is not None:
             if not ui_mapping:
                  st.error("The mapping file has no usable rows — every row is missing "
                           "its Uzio deduction name.")
             else:
                  st.success(f"Loaded {len(mrep['pairs'])} mapped deduction(s) from "
                             f"{m_file.name}.")

                  # The audit skips an unmapped deduction with a bare `continue`,
                  # so anything the mapping cannot answer has to be shown here or
                  # it disappears from the comparison with no trace.
                  unmapped = unmapped_source_deductions(adp_deductions, ui_mapping)
                  if unmapped:
                       st.warning(
                           f"{len(unmapped)} deduction(s) in the ADP file are not in the "
                           "mapping and will be EXCLUDED from the audit:"
                       )
                       st.write(", ".join(unmapped))

                  if mrep["blank_target"]:
                       st.info(
                           f"{len(mrep['blank_target'])} row(s) in the mapping have no Uzio "
                           "deduction (the setup helper leaves garnishments, child support "
                           "and tax liens unassigned) — also excluded: "
                           + ", ".join(mrep["blank_target"])
                       )

                  if mrep["conflicts"]:
                       st.warning(
                           "The mapping points the same source deduction at two different "
                           "Uzio deductions; the first was kept: "
                           + "; ".join(f"{k}: kept '{a}', ignored '{b}'"
                                       for k, a, b in mrep["conflicts"])
                       )

                  with st.expander(f"View the {len(mrep['pairs'])} mapped deduction(s)"):
                       st.dataframe(
                           pd.DataFrame(mrep["pairs"],
                                        columns=["ADP Deduction", "ADP Code", "Uzio Deduction"]),
                           use_container_width=True, hide_index=True,
                       )

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
