import io
import pandas as pd
import streamlit as st
from utils.audit_utils import generate_uzio_template, check_duplicate_columns, format_datetime_strings
from utils.ui_components import inject_premium_styles, render_premium_header, render_finding_card

APP_TITLE = "Paycom to Uzio Census Template Generator"

PAYCOM_FIELD_MAP = {
    'Employee ID': ['Employee_Code'],
    'First Name': ['Legal_Firstname'],
    'Last Name': ['Legal_Lastname'],
    'Middle Initial': ['Legal_Middle_Name'],
    'Employment Status': ['Employee_Status'],
    'Employment Type': ['DOL_Status'],
    'Hire Date': ['Most_Recent_Hire_Date'],
    'Original Hire Date': ['Hire_Date'],
    'Termination Date': ['Termination_Date'],
    'Termination Reason': ['Termination_Reason'],
    'Pay Type': ['Pay_Type'],
    'Annual Salary': ['Annual_Salary'],
    'Hourly Pay Rate': ['Rate_1'],
    'Working Hours': ['Scheduled_Pay_Period_Hours'],
    'Job Title': ['Position'],
    'Department': ['Department_Desc'],
    'Work Email': ['Work_Email'],
    'Personal Email': ['Personal_Email'],
    'Phone Number': ['Primary_Phone'],
    'SSN': ['SS_Number'],
    'DOB': ['Birth_Date_(MM/DD/YYYY)'],
    'Gender': ['Gender'],
    'Tobacco User': ['Tobacco_User'],
    'FLSA Classification': ['Exempt_Status'],
    'Address Line 1': ['Primary_Address_Line_1'],
    'Address Line 2': ['Primary_Address_Line_2'],
    'City': ['Primary_City/Municipality'],
    'Zip': ['Primary_Zip/Postal_Code'],
    'State': ['Primary_State/Province'],
    'Mailing Address Line 1': ['Mailing_Address_Line_1'],
    'Mailing Address Line 2': ['Mailing_Address_Line_2'],
    'Mailing City': ['Mailing_City/Municipality'],
    'Mailing Zip': ['Mailing_Zip/Postal_Code'],
    'Mailing State': ['Mailing_State/Province'],
    'License Number': ['DriversLicense'],
    'License Expiration Date': ['DLExpirationDate'],
    'Work Location': ['Work_Location']
}

ALLOWED_JOB_TITLES = [
    'DSP Owner', 'Operations Manager', 'Operations Lead', 'Fleet Manager', 
    'Safety Manager', 'Performance Manager', 'Trainer', 'Human Resources', 
    'Recruiter', 'Office Personnel', 'Payroll Assistant', 'Finance', 
    'Dispatch', 'Management', 'Admin', 'Survey', 'Warehouse', 'Walker', 
    'Driver', 'Helper', 'Driver-Lite', 'Driver-Step Van', 
    'Driver-Unscheduled', 'Lead Driver', 'DDU Dedicated', 'DDU Shared', 
    'Non-DSP Related', 'Driver-Major Appliance', 'E-Biker', 'TSO-PV Driver'
]

def norm_colname(c: str) -> str:
    import re
    if c is None: return ""
    c = str(c).replace("\n", " ").replace("\r", " ")
    c = c.replace("\u00A0", " ")
    c = c.replace("\u2019", "'").replace("\u201C", '"').replace("\u201D", '"')
    # Remove bracketed suffixes like (Personal Profile) or (Employment Profile - Pay Rates)
    c = re.sub(r'\(.*?\)', '', c)
    c = re.sub(r"\s+", " ", c).strip()
    c = c.replace("*", "")
    c = c.strip('"').strip("'")
    return c.lower()

def preprocess_paycom_file(paycom_file):
    """Common logic for reading and normalizing Paycom file."""
    # --- CRITICAL ERROR: Duplicate Column Check ---
    dupes = check_duplicate_columns(paycom_file)
    if dupes:
        st.error(f"⛔ **Critical Error: Duplicate Column Headers Found!**")
        st.markdown(f"The following column headers appear multiple times in your file: **{', '.join(dupes)}**")
        st.warning("Pandas cannot process files with duplicate headers accurately. Please delete the duplicate columns and re-upload the file.")
        return None, None, None, None

    try:
        if paycom_file.name.lower().endswith('.csv'):
            try:
                df_paycom = pd.read_csv(paycom_file, dtype=str)
            except UnicodeDecodeError:
                paycom_file.seek(0)
                df_paycom = pd.read_csv(paycom_file, dtype=str, encoding='latin1')
        else:
            df_paycom = pd.read_excel(paycom_file, dtype=str)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None, None, None, None

    # Save original column headers before normalization
    original_columns = list(df_paycom.columns)
    
    # Normalize source columns
    df_paycom.columns = [norm_colname(c) for c in df_paycom.columns]
    
    # Build mapping: normalized -> original (for restoring headers on download)
    norm_to_orig = dict(zip(df_paycom.columns, original_columns))
    
    # Resolve field map
    resolved_field_map = {}
    for std_name, vendor_cols in PAYCOM_FIELD_MAP.items():
        found = False
        for vc in vendor_cols:
            norm_vc = norm_colname(vc)
            if norm_vc in df_paycom.columns:
                resolved_field_map[std_name] = norm_vc
                found = True
                break
        if not found:
            resolved_field_map[std_name] = norm_colname(vendor_cols[0])
            
    return df_paycom, original_columns, norm_to_orig, resolved_field_map

def render_auto_fix_options(key_prefix):
    """All corrections are applied automatically — no user toggles needed."""
    return {
        'fix_flsa': True,
        'fix_emails': True,
        'fix_status': True,
        'fix_inactive': True,
        'fix_type': True,
        'fix_position': True,
        'fix_job_title': True,
        'fix_dol_status': True,
        'fix_std_hours': True,
        'fix_zip': True,
        'fix_driver_smart': True
    }

def get_manager_info(df_paycom, resolved_field_map):
    """Detection logic for top manager."""
    col_sup_code = None
    for cand in ['supervisor_primary_code', 'supervisor primary code', 'supervisorcode']:
        if cand in df_paycom.columns:
            col_sup_code = cand
            break

    top_manager_id = None
    top_manager_name = ""
    has_managers = False

    if col_sup_code:
        valid_sups = df_paycom[df_paycom[col_sup_code].notna() & (df_paycom[col_sup_code].astype(str).str.strip() != "")]
        if not valid_sups.empty:
            has_managers = True
            sup_counts = valid_sups[col_sup_code].value_counts()
            if not sup_counts.empty:
                top_manager_id = str(sup_counts.index[0]).strip()
                emp_code_col = resolved_field_map.get('Employee ID')
                if emp_code_col and emp_code_col in df_paycom.columns:
                    match = df_paycom[df_paycom[emp_code_col].astype(str).str.strip() == top_manager_id]
                    if not match.empty:
                        fn = match.iloc[0].get(resolved_field_map.get('First Name'), '')
                        ln = match.iloc[0].get(resolved_field_map.get('Last Name'), '')
                        if pd.notna(fn) and pd.notna(ln):
                            top_manager_name = f"{str(fn).strip()} {str(ln).strip()}".strip()
    return has_managers, top_manager_id, top_manager_name, col_sup_code

def render_census_sanity_check():
    inject_premium_styles()
    st.markdown("<div class='premium-card'>", unsafe_allow_html=True)
    st.title("📑 Paycom Census Sanity Check")
    st.markdown("Ensure your Paycom Census Export is audit-ready for Uzio benefits. This tool identifies critical data gaps, corrects formatting issues, and performs automated logic checks.")
    st.markdown("</div>", unsafe_allow_html=True)

    file_paycom = st.file_uploader("Upload Paycom Census Export (.xlsx, .csv)", type=["xlsx", "csv"], key="paycom_sanity_upload")
    if not file_paycom: return

    df_paycom, original_columns, norm_to_orig, resolved_field_map = preprocess_paycom_file(file_paycom)
    if df_paycom is None: return

    # --- PROVIDER MISMATCH DETECTION ---
    critical_paycom_cols = ['Employee ID', 'First Name', 'Last Name']
    missing_critical = [c for c in critical_paycom_cols if not resolved_field_map.get(c)]
    
    if len(missing_critical) > 1:
        st.error("### ⚠️ Mismatched File Detected")
        st.markdown(f"""
            <div class='action-hub-error'>
                <p>It looks like you've uploaded a file that doesn't match the <b>Paycom Census</b> format. 
                This often happens if an ADP export is uploaded here by mistake.</p>
                <p><b>Recommendation:</b> Please switch to the <b>ADP Hub</b> in the sidebar or upload a valid Paycom Census Export.</p>
            </div>
        """, unsafe_allow_html=True)
        return

    # --- MANAGER DETECTION ---
    has_managers, top_manager_id, top_manager_name, col_sup_code = get_manager_info(df_paycom, resolved_field_map)
    sort_by_manager = has_managers and top_manager_id is not None

    # --- What This Tool Does (Informational) ---
    with st.expander("ℹ️ What does this tool do automatically?", expanded=False):
        st.markdown("""
This tool automatically applies the following corrections to your Paycom Census data when you download the **Corrected Source**:

| Category | Auto-Fix Applied |
|---|---|
| **FLSA Alignment** | If FLSA is blank, sets Non-Exempt for Hourly and Exempt for Salaried employees |
| **Smart Driver Correction** | Fills blank Job Title, FLSA, and Pay Type for Driver roles using Department data |
| **Employment Status Mapping** | Maps non-standard statuses (e.g. Inactive → Terminated) |
| **Worker Category Mapping** | Maps categories like Intern → Part Time |
| **Email Fallback** | Uses Personal Email when Work Email is missing |
| **Job Title Fallback** | Fills blank Position from Department Description |
| **DOL Status Default** | Defaults blank DOL Status to Full-Time |
| **Working Hours** | Forces zero hours for Hourly employees |
| **Zip Code Cleanup** | Pads 4-digit zips, trims to 5-digits, removes special characters |
| **Manager Sorting** | Clusters managers at the top of the file by reportee count |
| **Date Formatting** | Standardizes all dates to MM/DD/YYYY |
        """)

    fix_options = render_auto_fix_options("pc_sanity")
    
    # --- MAPPING UI ---
    render_premium_header("🗺️ Mapping Configuration", "Provide mappings here to include them in the **Corrected Source** download.")
    
    src_loc_col = resolved_field_map.get('Work Location')
    unique_locs = sorted([str(l).strip() for l in df_paycom[src_loc_col].dropna().unique()]) if src_loc_col and src_loc_col in df_paycom.columns else []

    loc_dict = {}
    if unique_locs:
        st.write("**Work Location Mapping**")
        edited_locs = st.data_editor(
            pd.DataFrame({"Source Work Location": unique_locs, "Mapped Uzio Work Location": [""]*len(unique_locs)}),
            column_config={"Mapped Uzio Work Location": st.column_config.TextColumn("Enter Uzio Location", required=False)},
            hide_index=True, use_container_width=True, key="pc_sanity_loc_editor"
        )
        loc_dict = dict(zip(edited_locs['Source Work Location'], edited_locs['Mapped Uzio Work Location']))
    else:
        st.info("💡 No unique work locations detected in this file for mapping.")

    st.markdown("---")
    
    # --- PRE-GENERATION SANITY CHECKS ---
    from utils.audit_utils import validate_source_data
    validation = validate_source_data(df_paycom, resolved_field_map)

    hard_errors = validation['hard_errors']
    flsa_corrections = validation['flsa_corrections']
    flsa_blanks = validation['flsa_blanks']
    intern_corrections = validation['intern_corrections']
    email_fallbacks = validation['email_fallbacks']
    salaried_drivers = validation.get('salaried_drivers', pd.DataFrame())
    anomalies = validation.get('anomalies', pd.DataFrame())
    inactive_statuses = validation.get('inactive_statuses', pd.DataFrame())
    position_blanks = validation.get('position_blanks', pd.DataFrame())
    dol_status_blanks = validation.get('dol_status_blanks', pd.DataFrame())
    smart_driver_fixes = validation.get('smart_driver_fixes', pd.DataFrame())

    # --- UNIFIED VALIDATION & MAPPING CENTER ---
    has_issues = not hard_errors.empty or not flsa_corrections.empty or not flsa_blanks.empty or not intern_corrections.empty or not email_fallbacks.empty or not anomalies.empty
    
    if has_issues:
        with st.expander("🛠️ Census Integrity & Mapping Action Center", expanded=not hard_errors.empty):
            # 1. Critical Hard Errors (If Any)
            if not hard_errors.empty:
                # Categorized breakdown for the card (Actionable IDs with Consolidation & Scroll)
                import re
                from collections import defaultdict
                legend = {
                    'Missing/Duplicate Info': defaultdict(list), 
                    'Date & Status Logic': defaultdict(list), 
                    'Contact Formatting': defaultdict(list)
                }
                zip_count = 0
                
                for _, err in hard_errors.iterrows():
                    eid = str(err['Employee ID'])
                    issue_text = str(err['Issue'])
                    parts = [p.strip() for p in issue_text.split(",") if p.strip()]
                    for p in parts:
                        if "Zip Code" in p:
                            zip_count += 1
                            continue
                            
                        # --- CONSOLIDATION LOGIC (Sanitize specific values for grouping) ---
                        clean_issue = p
                        # 1. Group Date Mismatches
                        if "predates date of hire" in p:
                            clean_issue = "Termination date predates Hire date"
                        # 2. Group Special Characters
                        elif "Special characters in" in p:
                            match = re.search(r"Special characters in (.*?)(?:\s|$)", p)
                            if match:
                                clean_issue = f"Special characters in {match.group(1).strip()}"
                        
                        # Categorization Matcher
                        if any(m in clean_issue for m in ['Termination', 'Hire', 'Terminated', 'Non-standard', 'Annual Salary']): 
                            cat = 'Date & Status Logic'
                        elif 'Special characters' in clean_issue: 
                            cat = 'Contact Formatting'
                        else: 
                            cat = 'Missing/Duplicate Info'
                        
                        if eid not in legend[cat][clean_issue]:
                            legend[cat][clean_issue].append(eid)
                
                render_finding_card(
                    "Integrity Audit Results", 
                    {
                        "Critical Issues": len(hard_errors),
                        "Employees Affected": len(hard_errors['Employee ID'].unique()),
                        "Category": "Data Integrity"
                    },
                    type='error'
                )

                # --- SCROLLABLE ACTION CENTER ---
                with st.container(height=400, border=True):
                    lcol1, lcol2, lcol3 = st.columns(3)
                    
                    def render_actionable_list(title, items_dict):
                        st.markdown(f"**{title}**")
                        if items_dict:
                            for issue, ids in sorted(items_dict.items()):
                                id_str = ", ".join(ids[:3])
                                if len(ids) > 3: id_str += f" (+{len(ids)-3} more)"
                                # Highlight the issue name
                                st.markdown(f"- **{issue}** \n  `IDs: {id_str}`")
                        else:
                            st.markdown("_None_")

                    with lcol1:
                        render_actionable_list("Field Integrity", legend['Missing/Duplicate Info'])
                        if zip_count > 0:
                            st.info(f"ℹ️ {zip_count} Zip Code issues are hidden from this view but included in the download.")
                            
                    with lcol2:
                        render_actionable_list("Date & Status", legend['Date & Status Logic'])
                    with lcol3:
                        render_actionable_list("Formatting", legend['Contact Formatting'])
                    
                st.markdown("<br>", unsafe_allow_html=True)
                
                with st.expander("🔍 Show Affected Employee IDs", expanded=False):
                    st.dataframe(hard_errors, hide_index=True, use_container_width=True)
                
                st.markdown("<br>", unsafe_allow_html=True)

            # 2. Automated Mapping Suggestions
            st.markdown("##### 💡 Automated Mapping Suggestions")
            st.info("The following can be automatically applied using the checkboxes at the top.")
            wcol1, wcol2 = st.columns(2)
            
            def get_ids_str(df_in):
                if df_in.empty: return ""
                ids = df_in['Employee ID'].unique().tolist()
                id_str = ", ".join(str(x) for x in ids[:3])
                if len(ids) > 3: id_str += f" (+{len(ids)-3} more)"
                return f" `IDs: {id_str}`"

            with wcol1:
                st.markdown("**FLSA & Pay Rules**")
                if not flsa_corrections.empty: 
                    st.markdown(f"- ℹ️ **FLSA Mismatches:** {len(flsa_corrections)} employee(s) *(Pay Type vs FLSA conflict)*.{get_ids_str(flsa_corrections)}")
                if not flsa_blanks.empty: 
                    st.markdown(f"- ⚠️ **Blank FLSA:** {len(flsa_blanks)} employee(s) *(Missing Exempt/Non-Exempt)*.{get_ids_str(flsa_blanks)}")
                if not anomalies.empty: 
                    st.markdown(f"- ⚠️ **FLSA Anomalies:** {len(anomalies)} employee(s) *(Contradictory Logic)*.{get_ids_str(anomalies)}")
                if not smart_driver_fixes.empty: 
                    st.markdown(f"- 🚛 **Smart Driver Fixes:** {len(smart_driver_fixes)} employee(s) *(Auto-fill Driver data)*.{get_ids_str(smart_driver_fixes)}")
            with wcol2:
                st.markdown("**Employment & Contact**")
                if not intern_corrections.empty: 
                    st.markdown(f"- ⚠️ **Intern Codes:** {len(intern_corrections)} employee(s) *(Mapping Intern to Part Time)*.{get_ids_str(intern_corrections)}")
                if not email_fallbacks.empty: 
                    st.markdown(f"- 📧 **Email Fallbacks:** {len(email_fallbacks)} employee(s) *(Using personal where work email is missing)*.{get_ids_str(email_fallbacks)}")
                if not position_blanks.empty: 
                    st.markdown(f"- ℹ️ **Position Auto-Fill:** {len(position_blanks)} employee(s) *(Fallback blank Job Title to Dept)*.{get_ids_str(position_blanks)}")
    else:
        st.success("✅ Source data passed all integrity checks!")

    if st.button("Download Corrected Source", type="primary"):
        df_download = df_paycom.copy()

        # --- Collect Audit Info ---
        audit_trail = []
        emp_id_col = resolved_field_map.get('Employee ID')
        emp_name_col = next((c for c in df_download.columns if 'name' in str(c).lower()), None)
        
        def get_row_name(row_idx):
            if emp_name_col: return str(df_download.at[row_idx, emp_name_col]).strip()
            return "N/A"
        
        def log_change(row_idx, field, old_val, new_val, comment):
            eid = str(df_download.at[row_idx, emp_id_col]) if emp_id_col else "N/A"
            audit_trail.append({
                'Employee ID': eid,
                'Employee Name': get_row_name(row_idx),
                'Field Changed': field,
                'Old Value': str(old_val) if pd.notna(old_val) else "(blank)",
                'Assumed Value': str(new_val),
                'Comments': comment
            })
        
        # Resolve Position and Department Desc columns (normalized)
        col_dol = next((c for c in df_download.columns if str(c).lower().strip().replace('_',' ') == 'dol status'), None)
        col_emp_status = next((c for c in df_download.columns if str(c).lower().strip() in ['employee_status', 'employee status', 'employment status', 'status', 'ee status']), None)
        # Apply Fixes
        if fix_options.get('fix_position'):
            c_job = resolved_field_map.get('Job Title')
            c_dep = resolved_field_map.get('Department')
            if c_job and c_dep:
                mask = df_download[c_job].isna() | (df_download[c_job].astype(str).str.strip().str.lower() == "nan") | (df_download[c_job].astype(str).str.strip() == "")
                for idx in df_download[mask].index:
                    old_v = df_download.at[idx, c_job]
                    new_v = df_download.at[idx, c_dep]
                    if pd.notna(new_v) and str(new_v).strip():
                        df_download.at[idx, c_job] = new_v
                        log_change(idx, "Position", old_v, new_v, "Position was blank; filled using Department Description.")

        if fix_options.get('fix_driver_smart'):
            c_jt = resolved_field_map.get('Job Title')
            c_dept = resolved_field_map.get('Department')
            c_flsa = resolved_field_map.get('FLSA Classification')
            if c_jt and c_dept and c_flsa and c_jt in df_download.columns and c_dept in df_download.columns and c_flsa in df_download.columns:
                # 1. If Job is blank, check Dept for 'driver'
                mask_jt_blank = df_download[c_jt].isna() | (df_download[c_jt].astype(str).str.strip().str.lower() == "nan") | (df_download[c_jt].astype(str).str.strip() == "")
                mask_dept_driver = df_download[c_dept].astype(str).str.lower().str.contains("driver", na=False)
                
                for idx in df_download[mask_jt_blank & mask_dept_driver].index:
                    old_j = df_download.at[idx, c_jt]
                    new_j = df_download.at[idx, c_dept]
                    df_download.at[idx, c_jt] = new_j
                    log_change(idx, "Position (Smart Driver)", old_j, new_j, "Automatically assigned 'Driver' title from Department.")
                
                # 2. Now check if Job is Driver and FLSA is blank
                mask_job_driver = df_download[c_jt].astype(str).str.lower().str.contains("driver", na=False)
                mask_flsa_blank = df_download[c_flsa].isna() | (df_download[c_flsa].astype(str).str.strip().str.lower() == "nan") | (df_download[c_flsa].astype(str).str.strip() == "")
                
                for idx in df_download[mask_job_driver & mask_flsa_blank].index:
                    old_f = df_download.at[idx, c_flsa]
                    df_download.at[idx, c_flsa] = "Non-Exempt"
                    log_change(idx, "FLSA Classification (Smart Driver)", old_f, "Non-Exempt", "Automatic Non-Exempt status for Driver roles.")
                
                # 3. New: Check if Job is Driver and Pay Type is blank
                c_pt = resolved_field_map.get('Pay Type')
                if c_pt and c_pt in df_download.columns:
                    mask_pt_blank = df_download[c_pt].isna() | (df_download[c_pt].astype(str).str.strip().str.lower() == "nan") | (df_download[c_pt].astype(str).str.strip() == "")
                    for idx in df_download[mask_job_driver & mask_pt_blank].index:
                        old_p = df_download.at[idx, c_pt]
                        df_download.at[idx, c_pt] = "Hourly"
                        log_change(idx, "Pay Type (Smart Driver)", old_p, "Hourly", "Automatic Hourly pay type for Driver roles.")


        if fix_options.get('fix_emails'):
            c_work = next((col for col in df_download.columns if 'work_email' in str(col).lower()), None)
            c_pers = next((col for col in df_download.columns if 'personal_email' in str(col).lower()), None)
            if c_work and c_pers:
                mask = df_download[c_work].isna() | (df_download[c_work].astype(str).str.strip() == "")
                for idx in df_download[mask].index:
                    old_e = df_download.at[idx, c_work]
                    new_e = df_download.at[idx, c_pers]
                    if pd.notna(new_e) and str(new_e).strip():
                        df_download.at[idx, c_work] = new_e
                        log_change(idx, "Work Email", old_e, new_e, "Personal email used as fallback for missing work email.")

        if fix_options.get('fix_dol_status') and col_dol:
            mask_blank_dol = df_download[col_dol].isna() | (df_download[col_dol].astype(str).str.strip() == "")
            for idx in df_download[mask_blank_dol].index:
                old_d = df_download.at[idx, col_dol]
                df_download.at[idx, col_dol] = "Full-Time"
                log_change(idx, "DOL Status", old_d, "Full-Time", "Defaulted blank value to 'Full-Time' for active employee.")

        if fix_options.get('fix_std_hours'):
            c_sh = resolved_field_map.get('Working Hours')
            c_pt = resolved_field_map.get('Pay Type')
            if c_sh and c_sh in df_download.columns:
                if c_pt and c_pt in df_download.columns:
                    pt_lower = df_download[c_pt].astype(str).str.lower().str.strip()
                    mask_hourly = pt_lower.str.contains('hour', na=False)
                    for idx in df_download[mask_hourly].index:
                        old_v = str(df_download.at[idx, c_sh]).strip()
                        if old_v not in ["0", "0.0", ""]:
                            df_download.at[idx, c_sh] = "0"
                            log_change(idx, "Working Hours", old_v, "0", "Forced zero hours for Hourly employee.")
                        else:
                            df_download.at[idx, c_sh] = "0"

        if fix_options.get('fix_zip'):
            c_zip = resolved_field_map.get('Zip')
            c_mzip = resolved_field_map.get('Mailing Zip')
            for cz in [c_zip, c_mzip]:
                if cz and cz in df_download.columns:
                    def _fix_zip_local(z):
                        if pd.isna(z) or str(z).strip() == "": return ""
                        import re
                        s = str(z).split('.')[0].split('-')[0]
                        s = re.sub(r'[^0-9]', '', s)
                        if not s: return ""
                        if len(s) == 4: s = '0' + s
                        return s[:5]
                    
                    orig_zips = df_download[cz].copy()
                    df_download[cz] = df_download[cz].apply(_fix_zip_local).astype(str)
                    
                    # Log ZIP changes
                    for idx in df_download.index:
                        if str(df_download.at[idx, cz]) != str(orig_zips.at[idx]):
                            log_change(idx, cz, orig_zips.at[idx], df_download.at[idx, cz], "Standardized zip code format.")

        # Apply Mappings to download file
        if src_loc_col and src_loc_col in df_download.columns:
            def _map_loc(x):
                if pd.isna(x) or str(x).lower().strip() == 'nan' or str(x).strip() == "":
                    return ""
                curr = str(x).strip()
                return loc_dict.get(curr, curr)
            df_download[src_loc_col] = df_download[src_loc_col].apply(_map_loc)

        # Standardize ALL Date Columns to MM/DD/YYYY
        date_cols = [
            resolved_field_map.get('Hire Date'),              # Most_Recent_Hire_Date
            resolved_field_map.get('Original Hire Date'),     # Hire_Date
            resolved_field_map.get('Termination Date'),       # Termination_Date
            resolved_field_map.get('DOB'),                    # Birth_Date_(MM/DD/YYYY)
            resolved_field_map.get('License Expiration Date'),# DLExpirationDate
        ]
        date_cols = [c for c in date_cols if c is not None]
        df_download = format_datetime_strings(df_download, date_cols)

        if sort_by_manager and col_sup_code and col_sup_code in df_download.columns:
            emp_id_col = resolved_field_map.get('Employee ID')
            if emp_id_col and emp_id_col in df_download.columns:
                # Count reportees for each manager ID
                sup_counts = df_download[df_download[col_sup_code].notna()][col_sup_code].value_counts().to_dict()
                
                # 1. Primary Sort Key: Reportee Count (Managers with most reportees first)
                df_download['__mgr_count'] = df_download[emp_id_col].astype(str).str.strip().map(lambda x: sup_counts.get(x, 0))
                
                # 2. Secondary Sort Key: Manager's ID (to keep reportees under their specific manager)
                # We want the manager to be at the top of their group, followed by reportees.
                # So we map each employee to THEIR supervisor's count.
                df_download['__group_count'] = df_download[col_sup_code].astype(str).str.strip().map(lambda x: sup_counts.get(x, 0))
                
                # Sort: 
                # Highest Manager Count -> Descending
                # Group Count -> Descending (to keep reportees near their high-count managers)
                df_download = df_download.sort_values(by=['__mgr_count', '__group_count'], ascending=[False, False])
                df_download = df_download.drop(columns=['__mgr_count', '__group_count'])

        # --- New: Apply Strict Column Sequencing (Original Headers preserved) ---
        priority_keys = [
            'Employee ID', 'First Name', 'Last Name', 'Reports To ID',
            'Employment Type', 'Pay Type', 'Work Location', 'Workers Comp Code',
            'FLSA Classification', 'Employment Status', 'Job Title', 'Department'
        ]
        
        final_col_order = []
        renaming_dict = {}
        used_orig_cols = set()
        

            
        # 1. Add Priority Columns in exact order (Using original labels)
        for norm_key in priority_keys:
            orig_col_norm = resolved_field_map.get(norm_key)
            if orig_col_norm and orig_col_norm in df_download.columns:
                final_col_order.append(orig_col_norm)
                # Restore EXACT original header from source
                original_label = norm_to_orig.get(orig_col_norm, orig_col_norm)
                renaming_dict[orig_col_norm] = original_label
                used_orig_cols.add(orig_col_norm)
        
        # 2. Append all other original columns that weren't in the priority list
        for col in df_download.columns:
            if col not in used_orig_cols:
                final_col_order.append(col)
                # Ensure even non-priority columns map back to their original source names
                original_label = norm_to_orig.get(col, col)
                if original_label != col:
                    renaming_dict[col] = original_label

        # Reorder and Rename back to source headers
        df_download = df_download[final_col_order]
        df_download = df_download.rename(columns=renaming_dict)

        from utils.audit_utils import generate_excel_with_audit
        excel_data = generate_excel_with_audit(df_download, pd.DataFrame(audit_trail))
        stamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            st.download_button(
                label="📥 Download Corrected Source (XLSX)",
                data=excel_data,
                file_name=f"Paycom_Cleaned_{stamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pc_sanity_dl_xlsx",
            )
        with col_dl2:
            st.download_button(
                label="📥 Download Corrected Source (CSV)",
                data=df_download.to_csv(index=False).encode("utf-8"),
                file_name=f"Paycom_Cleaned_{stamp}.csv",
                mime="text/csv",
                key="pc_sanity_dl_csv",
                help="Single-sheet CSV of the cleaned census data.",
            )
        with col_dl3:
            df_audit = pd.DataFrame(audit_trail)
            st.download_button(
                label="📜 Download Change Log (CSV)",
                data=df_audit.to_csv(index=False).encode("utf-8"),
                file_name=f"Paycom_Change_Log_{stamp}.csv",
                mime="text/csv",
                key="pc_sanity_dl_audit",
                help="Download the audit trail showing all automated corrections made to the file.",
            )

    # --- Job Title Mapping Section ---
    from utils.job_title_mapper import render_streamlit_section as render_job_title_mapping
    render_job_title_mapping(df_paycom, "paycom", resolved_field_map, key_prefix="pc_sanity")

def render_census_generator():
    st.title("Paycom - Census Generator")
    st.markdown("""
    **Instructions**:
    1. Upload your **Paycom Census Export**.
    2. Map **Job Titles** and **Work Locations**.
    3. Download the fresh **Uzio Census Template**.
    """)
    
    paycom_file = st.file_uploader("Upload Paycom Census Export", type=["xlsx", "csv"], key="pc_gen_upload")
    if not paycom_file: return

    df_paycom, _, _, resolved_field_map = preprocess_paycom_file(paycom_file)
    if df_paycom is None: return

    fix_options = render_auto_fix_options("pc_gen")
    
    # Mapping Logic
    src_job_col = resolved_field_map.get('Job Title')
    src_loc_col = resolved_field_map.get('Work Location')
    unique_jobs = sorted([str(j).strip() for j in df_paycom[src_job_col].dropna().unique()]) if src_job_col else []
    unique_locs = sorted([str(l).strip() for l in df_paycom[src_loc_col].dropna().unique()]) if src_loc_col else []

    col1, col2 = st.columns(2)
    with col1:
        st.write("**Job Title Mapping**")
        edited_jobs = st.data_editor(
            pd.DataFrame({"Source Job Title": unique_jobs, "Mapped Uzio Job Title": [None]*len(unique_jobs)}),
            column_config={"Mapped Uzio Job Title": st.column_config.SelectboxColumn("Select Uzio Role", options=ALLOWED_JOB_TITLES, required=True)},
            hide_index=True, use_container_width=True, key="pc_job_editor"
        )
    with col2:
        st.write("**Work Location Mapping**")
        edited_locs = st.data_editor(
            pd.DataFrame({"Source Work Location": unique_locs, "Mapped Uzio Work Location": [""]*len(unique_locs)}),
            column_config={"Mapped Uzio Work Location": st.column_config.TextColumn("Enter Uzio Location", required=True)},
            hide_index=True, use_container_width=True, key="pc_loc_editor"
        )

    if st.button("Generate Uzio Template", type="primary", key="pc_gen_btn"):
        with st.spinner("Processing..."):
            try:
                job_dict = dict(zip(edited_jobs['Source Job Title'], edited_jobs['Mapped Uzio Job Title']))
                loc_dict = dict(zip(edited_locs['Source Work Location'], edited_locs['Mapped Uzio Work Location']))
                
                df_uzio = generate_uzio_template(df_paycom, resolved_field_map, fix_options=fix_options)
                
                # Apply Mappings
                if src_job_col: df_uzio['Job Title'] = df_paycom[src_job_col].astype(str).str.strip().map(job_dict).fillna(df_paycom[src_job_col])
                if src_loc_col: df_uzio['Work Location'] = df_paycom[src_loc_col].astype(str).str.strip().map(loc_dict).fillna(df_paycom[src_loc_col])

                # Inject into template
                from utils.audit_utils import inject_into_uzio_template
                wb = inject_into_uzio_template(df_uzio, template_path="templates/Uzio_Census_Template.xlsm")
                out = io.BytesIO()
                wb.save(out)
                out.seek(0)

                st.success("Template Generated!")
                st.download_button("Download Uzio Template", out.getvalue(), f"Uzio_Paycom_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsm", "application/vnd.ms-excel.sheet.macroEnabled.12")
            except Exception as e:
                st.error(f"Error: {e}")

def render_selective_census_generator():
    st.title("Paycom - Selective Census Generator")
    st.markdown("""
    **Instructions**:
    1. Upload your **Paycom Census Export**.
    2. Select columns to update and upload an existing **Uzio Template**.
    3. Download the updated template.
    """)
    
    paycom_file = st.file_uploader("Upload Paycom Census Export", type=["xlsx", "csv"], key="pc_sel_upload")
    if not paycom_file: return

    df_paycom, _, _, resolved_field_map = preprocess_paycom_file(paycom_file)
    if df_paycom is None: return

    fix_options = render_auto_fix_options("pc_sel")
    
    from utils.audit_utils import UZIO_RAW_MAPPING, read_uzio_raw_file, extract_mappings_from_uzio
    selected_uzio_cols = st.multiselect("🎯 Select Uzio Columns to Sync/Update", options=list(UZIO_RAW_MAPPING.keys()), default=["Employee SSN"], key="pc_sel_cols")
    
    uzio_template_file = st.file_uploader("📤 Upload Pre-filled Uzio Template (.xlsm)", type=["xlsm"], key="pc_uzio_template_v2")
    
    job_seeds, loc_seeds = {}, {}
    if uzio_template_file:
        df_seeds = read_uzio_raw_file(uzio_template_file)
        if df_seeds is not None:
            job_seeds, loc_seeds = extract_mappings_from_uzio(df_paycom, df_seeds, resolved_field_map)
        uzio_template_file.seek(0)

    src_job_col = resolved_field_map.get('Job Title')
    src_loc_col = resolved_field_map.get('Work Location')
    unique_jobs = sorted([str(j).strip() for j in df_paycom[src_job_col].dropna().unique()]) if src_job_col else []
    unique_locs = sorted([str(l).strip() for l in df_paycom[src_loc_col].dropna().unique()]) if src_loc_col else []

    col1, col2 = st.columns(2)
    with col1:
        edited_jobs = st.data_editor(
            pd.DataFrame({"Source Job Title": unique_jobs, "Mapped Uzio Job Title": [job_seeds.get(j) for j in unique_jobs]}),
            column_config={"Mapped Uzio Job Title": st.column_config.SelectboxColumn("Select Uzio Role", options=ALLOWED_JOB_TITLES, required=True)},
            hide_index=True, use_container_width=True, key="pc_job_editor_sel"
        )
    with col2:
        edited_locs = st.data_editor(
            pd.DataFrame({"Source Work Location": unique_locs, "Mapped Uzio Work Location": [loc_seeds.get(l, "") for l in unique_locs]}),
            column_config={"Mapped Uzio Work Location": st.column_config.TextColumn("Enter Uzio Location", required=True)},
            hide_index=True, use_container_width=True, key="pc_loc_editor_sel"
        )

    if st.button("Update Uzio Template", type="primary", key="pc_gen_btn_sel"):
        if not uzio_template_file: return st.error("Upload Uzio Template first.")
        with st.spinner("Processing..."):
            try:
                from utils.audit_utils import read_uzio_template_df, selective_update_uzio
                df_template = read_uzio_template_df(uzio_template_file)
                df_uzio, summary, _ = selective_update_uzio(df_paycom, df_template, selected_uzio_cols, resolved_field_map, fix_options=fix_options)
                
                # Apply Mappings
                job_dict = dict(zip(edited_jobs['Source Job Title'], edited_jobs['Mapped Uzio Job Title']))
                loc_dict = dict(zip(edited_locs['Source Work Location'], edited_locs['Mapped Uzio Work Location']))
                if src_job_col: df_uzio['Job Title'] = df_paycom[src_job_col].astype(str).str.strip().map(job_dict).fillna(df_paycom[src_job_col])
                if src_loc_col: df_uzio['Work Location'] = df_paycom[src_loc_col].astype(str).str.strip().map(loc_dict).fillna(df_paycom[src_loc_col])

                from utils.audit_utils import inject_into_uzio_template
                uzio_template_file.seek(0)
                wb = inject_into_uzio_template(df_uzio, uzio_template_file)
                out = io.BytesIO()
                wb.save(out)
                out.seek(0)

                st.success(summary)
                st.download_button("Download Updated Template", out.getvalue(), f"Uzio_Updated_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsm")
            except Exception as e:
                st.error(f"Error: {e}")

def render_ui():
    """Dispatcher for app.py (Optional)"""
    st.sidebar.title("Census Tools")
    tool = st.sidebar.selectbox("Select Tool", ["Sanity Check", "Full Generation", "Selective Sync"], key="pc_tool_select")
    if tool == "Sanity Check": render_census_sanity_check()
    elif tool == "Full Generation": render_census_generator()
    else: render_selective_census_generator()
