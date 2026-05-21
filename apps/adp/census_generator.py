import io
import pandas as pd
import streamlit as st
from utils.audit_utils import generate_uzio_template, check_duplicate_columns, format_datetime_strings, is_hourly_only_job_title
from utils.ui_components import inject_premium_styles, render_premium_header, render_validation_results, render_duplicate_column_error, render_missing_column_error, render_standardization_notice, render_sanity_disclaimer, REQUIRED_CENSUS_FIELDS

APP_TITLE = "ADP to Uzio Census Template Generator"

ADP_FIELD_MAP = {
    'Employee ID': ['Associate ID'],
    'First Name': ['Legal First Name'],
    'Last Name': ['Legal Last Name'],
    'Middle Initial': ['Legal Middle Name'],
    'Suffix': ['Generation Suffix Code'],
    'Employment Status': ['Position Status'],
    'Employment Type': ['Worker Category Description'],
    'Hire Date': ['Hire/Rehire Date'],
    'Original Hire Date': ['Hire Date'],
    'Termination Date': ['Termination Date'],
    'Termination Reason': ['Termination Reason Description'],
    'Pay Type': ['Regular Pay Rate Description'],
    'Annual Salary': ['Annual Salary'],
    'Hourly Pay Rate': ['Regular Pay Rate Amount'],
    'Working Hours': ['Standard Hours'],
    'Job Title': ['Job Title Description'],
    'Department': ['Department Description'],
    'Work Email': ['Work Contact: Work Email'],
    'Personal Email': ['Personal Contact: Personal Email'],
    'SSN': ['Tax ID (SSN)'],
    'DOB': ['Birth Date'],
    'Gender': ['Gender / Sex (Self-ID)'],
    'Tobacco User': ['Tobacco User'],
    'FLSA Classification': ['FLSA Description'],
    'Address Line 1': ['Primary Address: Address Line 1'],
    'Address Line 2': ['Primary Address: Address Line 2'],
    'City': ['Primary Address: City'],
    'Zip': ['Primary Address: Zip / Postal Code'],
    'State': ['Primary Address: State / Territory Code'],
    'Mailing Address Line 1': ['Legal / Preferred Address: Address Line 1'],
    'Mailing Address Line 2': ['Legal / Preferred Address: Address Line 2'],
    'Mailing City': ['Legal / Preferred Address: City'],
    'Mailing Zip': ['Legal / Preferred Address: Zip / Postal Code'],
    'Mailing State': ['Legal / Preferred Address: State / Territory Code'],
    'Reports To ID': ['Reports To Associate ID'],
    'Work Location': ['Location Description']
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
    c = re.sub(r'\(.*?\)', '', c)
    c = re.sub(r"\s+", " ", c).strip()
    c = c.replace("*", "")
    c = c.strip('"').strip("'")
    return c.lower()

def preprocess_adp_file(adp_file):
    """Common logic for reading and normalizing ADP file."""
    # --- CRITICAL ERROR: Duplicate Column Check ---
    dupes = check_duplicate_columns(adp_file)
    if dupes:
        render_duplicate_column_error(dupes)
        return None, None, None, None

    try:
        if adp_file.name.lower().endswith('.csv'):
            try:
                df_adp = pd.read_csv(adp_file, dtype=str)
            except UnicodeDecodeError:
                adp_file.seek(0)
                df_adp = pd.read_csv(adp_file, dtype=str, encoding='latin1')
        else:
            df_adp = pd.read_excel(adp_file, dtype=str)
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None, None, None, None

    # Save original column headers before normalization
    original_columns = list(df_adp.columns)
    
    # Normalize source columns
    df_adp.columns = [norm_colname(c) for c in df_adp.columns]
    
    # Build mapping: normalized -> original (for restoring headers on download)
    norm_to_orig = dict(zip(df_adp.columns, original_columns))
    
    # Resolve field map
    resolved_field_map = {}
    missing_required = []
    for std_name, vendor_cols in ADP_FIELD_MAP.items():
        found = False
        for vc in vendor_cols:
            norm_vc = norm_colname(vc)
            if norm_vc in df_adp.columns:
                resolved_field_map[std_name] = norm_vc
                found = True
                break
        if not found:
            resolved_field_map[std_name] = norm_colname(vendor_cols[0])
            if std_name in REQUIRED_CENSUS_FIELDS:
                missing_required.append((vendor_cols[0], std_name))

    # --- CRITICAL ERROR: Missing required columns ---
    if missing_required:
        render_missing_column_error(missing_required)
        return None, None, None, None

    return df_adp, original_columns, norm_to_orig, resolved_field_map

def render_auto_fix_options(key_prefix):
    """All corrections are applied automatically — no user toggles needed."""
    return {
        'fix_flsa': True,
        'fix_emails': True,
        'fix_job_title': True,
        'fix_status': True,
        'fix_inactive': True,
        'fix_type': True,
        'fix_dol_status': True,
        'fix_std_hours': True,
        'fix_zip': True,
        'fix_driver_smart': True,
        'fix_leave_to_active': True,
        'rename_zip_col': True,
        'replace_gender_col': True,
        'fix_blank_jt_to_driver': True
    }

def get_manager_info(df_adp, resolved_field_map):
    """Detection logic for top manager (ADP uses 'Reports To Associate ID')."""
    col_sup_code = resolved_field_map.get('Reports To ID')
    if not col_sup_code or col_sup_code not in df_adp.columns:
        if 'reports to associate id' in df_adp.columns:
            col_sup_code = 'reports to associate id'

    top_manager_id = None
    top_manager_name = ""
    has_managers = False

    if col_sup_code and col_sup_code in df_adp.columns:
        valid_sups = df_adp[df_adp[col_sup_code].notna() & (df_adp[col_sup_code].astype(str).str.strip() != "")]
        if not valid_sups.empty:
            has_managers = True
            sup_counts = valid_sups[col_sup_code].value_counts()
            if not sup_counts.empty:
                top_manager_id = str(sup_counts.index[0]).strip()
                emp_code_col = resolved_field_map.get('Employee ID')
                if emp_code_col and emp_code_col in df_adp.columns:
                    match = df_adp[df_adp[emp_code_col].astype(str).str.strip() == top_manager_id]
                    if not match.empty:
                        fn = match.iloc[0].get(resolved_field_map.get('First Name'), '')
                        ln = match.iloc[0].get(resolved_field_map.get('Last Name'), '')
                        if pd.notna(fn) and pd.notna(ln):
                            top_manager_name = f"{str(fn).strip()} {str(ln).strip()}".strip()
    return has_managers, top_manager_id, top_manager_name, col_sup_code

def render_census_sanity_check():
    inject_premium_styles()
    st.markdown("<div class='premium-card'>", unsafe_allow_html=True)
    st.title("📑 ADP Census Sanity Check")
    st.markdown("Ensure your ADP Census Export is audit-ready for Uzio benefits. This tool identifies critical data gaps, corrects formatting issues, and performs automated logic checks.")
    st.markdown("</div>", unsafe_allow_html=True)
    render_sanity_disclaimer()

    adp_file = st.file_uploader("Upload ADP Census Export (.xlsx, .csv)", type=["xlsx", "csv"], key="adp_sanity_upload")
    if not adp_file: return

    df_adp, original_columns, norm_to_orig, resolved_field_map = preprocess_adp_file(adp_file)
    if df_adp is None: return

    # --- PROVIDER MISMATCH DETECTION ---
    critical_adp_cols = ['Employee ID', 'First Name', 'Last Name']
    missing_critical = [c for c in critical_adp_cols if not resolved_field_map.get(c)]
    
    if len(missing_critical) > 1:
        st.error("⚠️ **Wrong file format** — this doesn't look like an ADP Census export. If you meant to upload a Paycom file, switch to the Paycom section in the sidebar.")
        return

    has_managers, top_manager_id, top_manager_name, col_sup_code = get_manager_info(df_adp, resolved_field_map)
    sort_by_manager = has_managers and top_manager_id is not None

    # --- What This Tool Does (Informational) ---
    with st.expander("ℹ️ What does this tool fix automatically?", expanded=False):
        st.markdown("""
When you click **Download Corrected Source**, the following corrections are applied automatically:

| What we check | What happens automatically |
|---|---|
| **Driver/Walker/Helper roles** | Forced to hourly pay and non-exempt status, even if the source file says salary |
| **Missing pay classification** | Filled in based on pay type — hourly employees get Non-Exempt, salaried get Exempt |
| **Unresolvable pay classification** | Left blank and flagged in the Change Log for manual review |
| **Driver roles with missing pay info** | Job title, pay type, and classification auto-filled using department data |
| **Blank job title (hourly non-exempt)** | Defaulted to "Driver" |
| **Non-standard employment status** | Updated to Active or Terminated (e.g., Inactive → Terminated) |
| **Intern employment type** | Changed to Part-Time |
| **Employees on leave with no termination date** | Set to Active — please mark them as excluded from payroll in Uzio |
| **Missing work email** | Filled using personal email as a backup |
| **Blank job title** | Filled using the department name |
| **Blank employment type** | Set to Full-Time |
| **Working hours (all employees)** | Set to 0 for every employee — hourly and salaried, blank or filled |
| **Zip codes** | Padded to 5 digits, trimmed if too long, special characters removed |
| **Zip column header** | The "Zip / Postal Code" column header is renamed to "Zip Code" |
| **Gender column** | The Gender column is populated from the "Sex" column's values |
| **Manager ordering** | Managers moved to the top of the file |
| **Dates** | All dates formatted as MM/DD/YYYY |
        """)

    fix_options = render_auto_fix_options("adp_sanity")
    
    # --- MAPPING UI ---
    render_premium_header("🗺️ Mapping Configuration", "Provide mappings here to include them in the **Corrected Source** download.")
    
    src_loc_col = resolved_field_map.get('Work Location')
    unique_locs = sorted([str(l).strip() for l in df_adp[src_loc_col].dropna().unique()]) if src_loc_col and src_loc_col in df_adp.columns else []

    loc_dict = {}
    if unique_locs:
        st.write("**Work Location Mapping**")
        edited_locs = st.data_editor(
            pd.DataFrame({"Source Work Location": unique_locs, "Mapped Uzio Work Location": [""]*len(unique_locs)}),
            column_config={"Mapped Uzio Work Location": st.column_config.TextColumn("Enter Uzio Location", required=False)},
            hide_index=True, use_container_width=True, key="adp_sanity_loc_editor"
        )
        loc_dict = dict(zip(edited_locs['Source Work Location'], edited_locs['Mapped Uzio Work Location']))
    else:
        st.info("💡 No unique work locations detected in this file for mapping.")

    # New: Default Work Location for Blanks
    default_work_loc = st.text_input("📍 Default Work Location for Blanks", placeholder="Type to fill all blank work locations...", key="adp_sanity_default_loc_input")
    st.markdown("---")
    
    from utils.audit_utils import validate_source_data
    validation = validate_source_data(df_adp, resolved_field_map)

    hard_errors = validation['hard_errors']
    flsa_corrections = validation['flsa_corrections']
    flsa_blanks = validation['flsa_blanks']
    intern_corrections = validation['intern_corrections']
    email_fallbacks = validation['email_fallbacks']
    anomalies = validation.get('anomalies', pd.DataFrame())
    smart_driver_fixes = validation.get('smart_driver_fixes', pd.DataFrame())
    position_blanks = validation.get('position_blanks', pd.DataFrame())

    # --- VALIDATION RESULTS (plain-English, two-section layout) ---
    render_validation_results(
        hard_errors=hard_errors,
        flsa_corrections=flsa_corrections,
        flsa_blanks=flsa_blanks,
        anomalies=anomalies,
        intern_corrections=intern_corrections,
        email_fallbacks=email_fallbacks,
        smart_driver_fixes=smart_driver_fixes,
        position_blanks=position_blanks,
    )

    # --- Persistent Download Section ---
    st.markdown("---")
    st.subheader("📥 Download Cleaned Results")
    render_standardization_notice(include_column_renames=True)
    st.caption("Click the button below to prepare and download the corrected census files.")
    show_key = f"adp_sanity_show_dl_v3"
    if st.button("Download Corrected Source", type="primary", key="adp_sanity_main_btn_v2"):
        st.session_state[show_key] = True

    if st.session_state.get(show_key):
        # We use session state to cache the data so it doesn't re-run the logic on every download click
        data_key = "adp_sanity_cached_files"
        if data_key not in st.session_state:
            with st.spinner("Preparing downloads..."):
                df_download = df_adp.copy()
                
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

                def log_summary(field, new_val, comment):
                    # File-wide standardization — one summary row, not per-employee.
                    audit_trail.append({
                        'Employee ID': '(All employees)',
                        'Employee Name': '—',
                        'Field Changed': field,
                        'Old Value': '—',
                        'Assumed Value': new_val,
                        'Comments': comment
                    })
                
                # Apply Fixes
                if fix_options.get('fix_emails'):
                    c_work = resolved_field_map.get('Work Email')
                    c_pers = resolved_field_map.get('Personal Email')
                    if c_work and c_pers and c_work in df_download.columns and c_pers in df_download.columns:
                        mask = df_download[c_work].isna() | (df_download[c_work].astype(str).str.strip() == "")
                        for idx in df_download[mask].index:
                            old_e = df_download.at[idx, c_work]
                            new_e = df_download.at[idx, c_pers]
                            if pd.notna(new_e) and str(new_e).strip():
                                df_download.at[idx, c_work] = new_e
                                log_change(idx, "Work Email", old_e, new_e, "Personal email used as fallback for missing work email.")

                if fix_options.get('fix_leave_to_active'):
                    c_pos = resolved_field_map.get('Employment Status')
                    c_term = resolved_field_map.get('Termination Date')
                    if c_pos and c_term and c_pos in df_download.columns and c_term in df_download.columns:
                        pos_series = df_download[c_pos].astype(str).str.strip().str.lower()
                        term_series = df_download[c_term].astype(str).str.strip().str.lower()
                        
                        mask_special = pos_series.str.contains('leave|inactive', na=False)
                        mask_term_blank = df_download[c_term].isna() | (term_series == "") | (term_series == "nan")
                        
                        # Case A: Special Status & No Term Date -> Active (Exclude from Payroll)
                        for idx in df_download[mask_special & mask_term_blank].index:
                            old_p = df_download.at[idx, c_pos]
                            df_download.at[idx, c_pos] = "Active"
                            log_change(idx, "Employment Status", old_p, "Active", "Please make it exclude from payroll in Uzio")
                        
                        # Case B: Special Status & HAS Term Date -> Terminated
                        for idx in df_download[mask_special & ~mask_term_blank].index:
                            old_p = df_download.at[idx, c_pos]
                            df_download.at[idx, c_pos] = "Terminated"
                            log_change(idx, "Employment Status", old_p, "Terminated", "Converted to 'Terminated' due to presence of Termination Date.")

                if fix_options.get('fix_dol_status'):
                    c_dol = resolved_field_map.get('Employment Type')
                    if c_dol and c_dol in df_download.columns:
                        mask_blank = df_download[c_dol].isna() | (df_download[c_dol].astype(str).str.strip().str.lower() == "nan") | (df_download[c_dol].astype(str).str.strip() == "")
                        for idx in df_download[mask_blank].index:
                            old_d = df_download.at[idx, c_dol]
                            df_download.at[idx, c_dol] = "Full Time"
                            log_change(idx, "Employment Type", old_d, "Full Time", "Defaulted blank value to 'Full Time' for active employee.")

                if fix_options.get('fix_job_title'):
                    c_job = resolved_field_map.get('Job Title')
                    c_dep = resolved_field_map.get('Department')
                    if c_job and c_dep and c_job in df_download.columns and c_dep in df_download.columns:
                        mask = df_download[c_job].isna() | (df_download[c_job].astype(str).str.strip().str.lower() == "nan") | (df_download[c_job].astype(str).str.strip() == "")
                        for idx in df_download[mask].index:
                            old_val = df_download.at[idx, c_job]
                            new_val = df_download.at[idx, c_dep]
                            if pd.notna(new_val) and str(new_val).strip():
                                df_download.at[idx, c_job] = new_val
                                log_change(idx, "Job Title", old_val, new_val, "Position was blank; filled using Department Description.")

                if fix_options.get('fix_flsa'):
                    c_flsa = resolved_field_map.get('FLSA Classification')
                    c_pt = resolved_field_map.get('Pay Type')
                    c_jt = resolved_field_map.get('Job Title')
                    if c_flsa and c_flsa in df_download.columns:
                        # Rule 1 (always wins): Driver / hourly-only Job Title forces
                        # Pay Type = Hourly + FLSA = Non-Exempt, overwriting source.
                        mask_jt_driver = pd.Series(False, index=df_download.index)
                        if c_jt and c_jt in df_download.columns:
                            mask_jt_driver = df_download[c_jt].apply(is_hourly_only_job_title)
                            for idx in df_download[mask_jt_driver].index:
                                old_f = df_download.at[idx, c_flsa]
                                cur_lower = str(old_f).strip().lower() if pd.notna(old_f) else ""
                                if cur_lower != 'non-exempt':
                                    df_download.at[idx, c_flsa] = "Non-Exempt"
                                    log_change(idx, "FLSA Classification", old_f, "Non-Exempt", "Forced Non-Exempt for Driver/Hourly-only Position.")
                            if c_pt and c_pt in df_download.columns:
                                for idx in df_download[mask_jt_driver].index:
                                    old_p = df_download.at[idx, c_pt]
                                    cur_lower = str(old_p).strip().lower() if pd.notna(old_p) else ""
                                    if cur_lower != 'hourly':
                                        df_download.at[idx, c_pt] = "Hourly"
                                        log_change(idx, "Pay Type", old_p, "Hourly", "Forced Hourly for Driver/Hourly-only Position.")

                        # Rules 2–4 (non-Driver rows only, blanks only — never overwrite
                        # a populated source FLSA value by Pay Type alone).
                        if c_pt and c_pt in df_download.columns:
                            mask_flsa_blank = df_download[c_flsa].isna() | (df_download[c_flsa].astype(str).str.strip().str.lower().isin(["nan", ""]))
                            for idx in df_download[mask_flsa_blank & ~mask_jt_driver].index:
                                pt_val = str(df_download.at[idx, c_pt]).lower().strip()
                                pt_raw = str(df_download.at[idx, c_pt]).strip() if pd.notna(df_download.at[idx, c_pt]) else ""
                                old_f = df_download.at[idx, c_flsa]
                                if 'hour' in pt_val:
                                    df_download.at[idx, c_flsa] = "Non-Exempt"
                                    log_change(idx, "FLSA Classification", old_f, "Non-Exempt", "Filled blank FLSA based on Hourly Pay Type.")
                                elif 'salar' in pt_val:
                                    df_download.at[idx, c_flsa] = "Exempt"
                                    log_change(idx, "FLSA Classification", old_f, "Exempt", "Filled blank FLSA based on Salaried Pay Type.")
                                else:
                                    # Rule 4: cannot determine — leave blank, flag for review.
                                    log_change(idx, "FLSA Classification", old_f, "(Blank — Not Filled)",
                                               f"Cannot derive FLSA — source FLSA is blank, Job Title is not in Driver/Hourly-only list, and Pay Type is '{pt_raw or '(Blank)'}'. Manual review required.")

                if fix_options.get('fix_driver_smart'):
                    c_jt = resolved_field_map.get('Job Title')
                    c_dept = resolved_field_map.get('Department')
                    c_flsa = resolved_field_map.get('FLSA Classification')
                    if c_jt and c_dept and c_flsa and c_jt in df_download.columns and c_dept in df_download.columns and c_flsa in df_download.columns:
                        mask_jt_blank = df_download[c_jt].isna() | (df_download[c_jt].astype(str).str.strip().str.lower() == "nan") | (df_download[c_jt].astype(str).str.strip() == "")
                        mask_dept_driver = df_download[c_dept].astype(str).str.lower().str.contains("driver", na=False)
                        for idx in df_download[mask_jt_blank & mask_dept_driver].index:
                            old_j = df_download.at[idx, c_jt]
                            new_j = df_download.at[idx, c_dept]
                            df_download.at[idx, c_jt] = new_j
                            log_change(idx, "Job Title (Smart Driver)", old_j, new_j, "Automatically assigned 'Driver' title from Department.")
                        
                        mask_job_driver = df_download[c_jt].astype(str).str.lower().str.contains("driver", na=False)
                        mask_flsa_blank = df_download[c_flsa].isna() | (df_download[c_flsa].astype(str).str.strip().str.lower() == "nan") | (df_download[c_flsa].astype(str).str.strip() == "")
                        for idx in df_download[mask_job_driver & mask_flsa_blank].index:
                            old_f = df_download.at[idx, c_flsa]
                            df_download.at[idx, c_flsa] = "Non-Exempt"
                            log_change(idx, "FLSA Classification (Smart Driver)", old_f, "Non-Exempt", "Automatic Non-Exempt status for Driver roles.")
                        
                        c_pt = resolved_field_map.get('Pay Type')
                        if c_pt and c_pt in df_download.columns:
                            mask_pt_blank = df_download[c_pt].isna() | (df_download[c_pt].astype(str).str.strip().str.lower() == "nan") | (df_download[c_pt].astype(str).str.strip() == "")
                            for idx in df_download[mask_job_driver & mask_pt_blank].index:
                                old_p = df_download.at[idx, c_pt]
                                df_download.at[idx, c_pt] = "Hourly"
                                log_change(idx, "Pay Type (Smart Driver)", old_p, "Hourly", "Automatic Hourly pay type for Driver roles.")

                if fix_options.get('fix_blank_jt_to_driver'):
                    c_jt = resolved_field_map.get('Job Title')
                    c_flsa = resolved_field_map.get('FLSA Classification')
                    c_pt = resolved_field_map.get('Pay Type')
                    if c_jt and c_flsa and c_pt and c_jt in df_download.columns and c_flsa in df_download.columns and c_pt in df_download.columns:
                        jt_series = df_download[c_jt].astype(str).str.strip().str.lower()
                        flsa_series = df_download[c_flsa].astype(str).str.strip().str.lower()
                        pt_series = df_download[c_pt].astype(str).str.strip().str.lower()
                        mask_jt_blank = df_download[c_jt].isna() | (jt_series == "") | (jt_series == "nan")
                        mask_non_exempt = flsa_series.str.contains("non-exempt", na=False) | flsa_series.str.contains("non exempt", na=False)
                        mask_hourly = pt_series.str.contains("hourly", na=False)
                        for idx in df_download[mask_jt_blank & mask_non_exempt & mask_hourly].index:
                            old_j = df_download.at[idx, c_jt]
                            df_download.at[idx, c_jt] = "Driver"
                            log_change(idx, "Job Title", old_j, "Driver", "Job Title was blank; defaulted to 'Driver' for Non-Exempt Hourly employee.")

                if fix_options.get('fix_std_hours'):
                    c_sh = resolved_field_map.get('Working Hours')
                    if c_sh and c_sh in df_download.columns:
                        # Working Hours are zeroed for EVERY employee — hourly and
                        # salaried — regardless of the source value (blank or filled).
                        for idx in df_download.index:
                            old_v = str(df_download.at[idx, c_sh]).strip()
                            df_download.at[idx, c_sh] = "0"
                            if old_v.lower() not in ["0", "0.0", "", "nan"]:
                                log_change(idx, "Working Hours", old_v, "0", "Working hours set to 0 for all employees.")

                if fix_options.get('rename_zip_col'):
                    c_zip = resolved_field_map.get('Zip')
                    if c_zip and c_zip in norm_to_orig:
                        old_label = norm_to_orig[c_zip]
                        norm_to_orig[c_zip] = "Primary Address: Zip Code"
                        log_summary("Column header — Zip", "Primary Address: Zip Code",
                                    f"Home-zip column header standardized from '{old_label}' to 'Primary Address: Zip Code'.")

                if fix_options.get('replace_gender_col'):
                    sex_col = norm_colname("Sex")
                    c_gender = resolved_field_map.get('Gender')
                    if sex_col in df_download.columns:
                        if c_gender and c_gender in df_download.columns and c_gender != sex_col:
                            df_download = df_download.drop(columns=[c_gender])
                        norm_to_orig[sex_col] = "Gender / Sex (Self-ID)"
                        log_summary("Gender column", "Populated from Sex column",
                                    "Gender column populated from the 'Sex' column's values.")

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
                            for idx in df_download.index:
                                if str(df_download.at[idx, cz]) != str(orig_zips.at[idx]):
                                    log_change(idx, cz, orig_zips.at[idx], df_download.at[idx, cz], "Standardized zip code format.")

                # Emergency Contact Cleanup
                emergency_cols = [c for c in df_download.columns if 'emergency' in str(c).lower()]
                for ec in emergency_cols:
                    mask_fian = df_download[ec].astype(str).str.lower().str.startswith("fian", na=False)
                    for idx in df_download[mask_fian].index:
                        old_v = df_download.at[idx, ec]
                        if old_v != "Fiancee":
                            df_download.at[idx, ec] = "Fiancee"
                            log_change(idx, ec, old_v, "Fiancee", "Standardized 'Fiancée' or similar to 'Fiancee' for system compatibility.")

                # Fix Work Locations
                if src_loc_col and src_loc_col in df_download.columns:
                    # 1. Fill blanks with user default if provided
                    dwl = st.session_state.get("adp_sanity_default_loc_input", "").strip()
                    if dwl:
                        mask_blank = df_download[src_loc_col].isna() | (df_download[src_loc_col].astype(str).str.strip().str.lower().isin(["nan", ""]))
                        for idx in df_download[mask_blank].index:
                            old_val = df_download.at[idx, src_loc_col]
                            df_download.at[idx, src_loc_col] = dwl
                            log_change(idx, "Work Location", old_val, dwl, "Filled blank location with user-provided default.")
                    
                    # 2. Apply mapping from data editor
                    df_download[src_loc_col] = df_download[src_loc_col].astype(str).str.strip().map(lambda x: loc_dict.get(x, x))

                date_cols = [resolved_field_map.get('Hire Date'), resolved_field_map.get('Original Hire Date'), resolved_field_map.get('Termination Date'), resolved_field_map.get('DOB')]
                date_cols = [c for c in date_cols if c is not None]
                if date_cols:
                    df_download = format_datetime_strings(df_download, date_cols)
                    log_summary("Date format", "MM/DD/YYYY",
                                "All date columns (hire, termination, birth) standardized to MM/DD/YYYY format.")

                if sort_by_manager and col_sup_code and col_sup_code in df_download.columns:
                    emp_id_col = resolved_field_map.get('Employee ID')
                    if emp_id_col and emp_id_col in df_download.columns:
                        sup_counts = df_download[df_download[col_sup_code].notna()][col_sup_code].value_counts().to_dict()
                        df_download['__mgr_count'] = df_download[emp_id_col].astype(str).str.strip().map(lambda x: sup_counts.get(x, 0))
                        df_download['__group_count'] = df_download[col_sup_code].astype(str).str.strip().map(lambda x: sup_counts.get(x, 0))
                        df_download = df_download.sort_values(by=['__mgr_count', '__group_count'], ascending=[False, False])
                        df_download = df_download.drop(columns=['__mgr_count', '__group_count'])
                        log_summary("Row order", "Grouped by manager",
                                    "Employee rows reordered so each manager is clustered with their reportees.")

                priority_keys = ['Employee ID', 'First Name', 'Last Name', 'Reports To ID', 'Employment Type', 'Pay Type', 'Work Location', 'Workers Comp Code', 'FLSA Classification', 'Employment Status', 'Job Title', 'Department']
                final_col_order = []
                renaming_dict = {}
                used_orig_cols = set()
                for norm_key in priority_keys:
                    orig_col_norm = resolved_field_map.get(norm_key)
                    if orig_col_norm and orig_col_norm in df_download.columns:
                        final_col_order.append(orig_col_norm)
                        original_label = norm_to_orig.get(orig_col_norm, orig_col_norm)
                        renaming_dict[orig_col_norm] = original_label
                        used_orig_cols.add(orig_col_norm)
                for col in df_download.columns:
                    if col not in used_orig_cols:
                        final_col_order.append(col)
                        original_label = norm_to_orig.get(col, col)
                        if original_label != col:
                            renaming_dict[col] = original_label

                df_download = df_download[final_col_order].rename(columns=renaming_dict)
                log_summary("Column order", "Key fields first",
                            "Columns reordered so key fields (Employee ID, Name, Pay Type, FLSA, etc.) appear first.")
                from utils.audit_utils import generate_excel_with_audit
                st.session_state[data_key] = {
                    "xlsx": generate_excel_with_audit(df_download, pd.DataFrame(audit_trail)),
                    "csv": df_download.to_csv(index=False).encode("utf-8-sig"),
                    "audit": pd.DataFrame(audit_trail).to_csv(index=False).encode("utf-8-sig")
                }

        stamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
        cached = st.session_state.get(data_key, {})
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            st.download_button("📥 Download Corrected Source (XLSX)", cached.get("xlsx", b""), f"ADP_Cleaned_{stamp}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="adp_sanity_dl_xlsx")
        with col_dl2:
            st.download_button("📥 Download Corrected Source (CSV)", cached.get("csv", b""), f"ADP_Cleaned_{stamp}.csv", "text/csv", key="adp_sanity_dl_csv")
        with col_dl3:
            st.download_button("📜 Download Change Log (CSV)", cached.get("audit", b""), f"ADP_Change_Log_{stamp}.csv", "text/csv", key="adp_sanity_dl_audit")
        
        st.info("The Change Log is a separate audit trail showing all automated corrections made to the file.")

    # --- Job Title Mapping Section ---
    from utils.job_title_mapper import render_streamlit_section as render_job_title_mapping
    render_job_title_mapping(df_adp, "adp", resolved_field_map, key_prefix="adp_sanity")

def render_census_generator():
    st.title("ADP - Full Census Generation")
    
    adp_file = st.file_uploader("Upload ADP Census Export", type=["xlsx", "csv"], key="adp_gen_upload")
    if not adp_file: return

    df_adp, _, _, resolved_field_map = preprocess_adp_file(adp_file)
    if df_adp is None: return

    fix_options = render_auto_fix_options("adp_gen")
    
    src_job_col = resolved_field_map.get('Job Title')
    src_loc_col = resolved_field_map.get('Work Location')
    unique_jobs = sorted([str(j).strip() for j in df_adp[src_job_col].dropna().unique()]) if src_job_col else []
    unique_locs = sorted([str(l).strip() for l in df_adp[src_loc_col].dropna().unique()]) if src_loc_col else []

    col1, col2 = st.columns(2)
    with col1:
        st.write("**Job Title Mapping**")
        edited_jobs = st.data_editor(
            pd.DataFrame({"Source Job Title": unique_jobs, "Mapped Uzio Job Title": [None]*len(unique_jobs)}),
            column_config={"Mapped Uzio Job Title": st.column_config.SelectboxColumn("Select Uzio Role", options=ALLOWED_JOB_TITLES, required=True)},
            hide_index=True, use_container_width=True, key="adp_job_editor"
        )
    with col2:
        st.write("**Work Location Mapping**")
        edited_locs = st.data_editor(
            pd.DataFrame({"Source Work Location": unique_locs, "Mapped Uzio Work Location": [""]*len(unique_locs)}),
            column_config={"Mapped Uzio Work Location": st.column_config.TextColumn("Enter Uzio Location", required=True)},
            hide_index=True, use_container_width=True, key="adp_loc_editor"
        )

    if st.button("Generate Uzio Template", type="primary"):
        with st.spinner("Processing..."):
            try:
                job_dict = dict(zip(edited_jobs['Source Job Title'], edited_jobs['Mapped Uzio Job Title']))
                loc_dict = dict(zip(edited_locs['Source Work Location'], edited_locs['Mapped Uzio Work Location']))
                
                df_uzio = generate_uzio_template(df_adp, resolved_field_map, fix_options=fix_options)
                
                if src_job_col: df_uzio['Job Title'] = df_adp[src_job_col].astype(str).str.strip().map(job_dict).fillna(df_adp[src_job_col])
                if src_loc_col: df_uzio['Work Location'] = df_adp[src_loc_col].astype(str).str.strip().map(loc_dict).fillna(df_adp[src_loc_col])

                from utils.audit_utils import inject_into_uzio_template
                wb = inject_into_uzio_template(df_uzio, template_path="templates/Uzio_Census_Template.xlsm")
                out = io.BytesIO()
                wb.save(out)
                out.seek(0)

                st.success("Template Generated!")
                st.download_button("Download Uzio Template", out.getvalue(), f"Uzio_ADP_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsm")
            except Exception as e:
                st.error(f"Error: {e}")

def render_selective_census_generator():
    st.title("ADP - Selective Census Sync")
    
    adp_file = st.file_uploader("Upload ADP Census Export", type=["xlsx", "csv"], key="adp_sel_upload")
    if not adp_file: return

    df_adp, _, _, resolved_field_map = preprocess_adp_file(adp_file)
    if df_adp is None: return

    fix_options = render_auto_fix_options("adp_sel")
    
    from utils.audit_utils import UZIO_RAW_MAPPING, read_uzio_raw_file, extract_mappings_from_uzio
    selected_uzio_cols = st.multiselect("🎯 Select Uzio Columns to Sync", options=list(UZIO_RAW_MAPPING.keys()), default=["Employee SSN"])
    
    uzio_template_file = st.file_uploader("📤 Upload Pre-filled Uzio Template (.xlsm)", type=["xlsm"], key="adp_uzio_template_sel")
    
    job_seeds, loc_seeds = {}, {}
    if uzio_template_file:
        df_seeds = read_uzio_raw_file(uzio_template_file)
        if df_seeds is not None: job_seeds, loc_seeds = extract_mappings_from_uzio(df_adp, df_seeds, resolved_field_map)
        uzio_template_file.seek(0)

    src_job_col = resolved_field_map.get('Job Title')
    src_loc_col = resolved_field_map.get('Work Location')
    unique_jobs = sorted([str(j).strip() for j in df_adp[src_job_col].dropna().unique()]) if src_job_col else []
    unique_locs = sorted([str(l).strip() for l in df_adp[src_loc_col].dropna().unique()]) if src_loc_col else []

    col1, col2 = st.columns(2)
    with col1:
        edited_jobs = st.data_editor(
            pd.DataFrame({"Source Job Title": unique_jobs, "Mapped Uzio Job Title": [job_seeds.get(j) for j in unique_jobs]}),
            column_config={"Mapped Uzio Job Title": st.column_config.SelectboxColumn("Select Uzio Role", options=ALLOWED_JOB_TITLES, required=True)},
            hide_index=True, use_container_width=True, key="adp_job_editor_sel"
        )
    with col2:
        edited_locs = st.data_editor(
            pd.DataFrame({"Source Work Location": unique_locs, "Mapped Uzio Work Location": [loc_seeds.get(l, "") for l in unique_locs]}),
            column_config={"Mapped Uzio Work Location": st.column_config.TextColumn("Enter Uzio Location", required=True)},
            hide_index=True, use_container_width=True, key="adp_loc_editor_sel"
        )

    if st.button("Update Uzio Template", type="primary"):
        if not uzio_template_file: return st.error("Upload Uzio Template first.")
        with st.spinner("Processing..."):
            try:
                from utils.audit_utils import read_uzio_template_df, selective_update_uzio
                df_template = read_uzio_template_df(uzio_template_file)
                df_uzio, summary, _ = selective_update_uzio(df_adp, df_template, selected_uzio_cols, resolved_field_map, fix_options=fix_options)
                
                # Apply Mappings
                job_dict = dict(zip(edited_jobs['Source Job Title'], edited_jobs['Mapped Uzio Job Title']))
                loc_dict = dict(zip(edited_locs['Source Work Location'], edited_locs['Mapped Uzio Work Location']))
                if src_job_col: df_uzio['Job Title'] = df_adp[src_job_col].astype(str).str.strip().map(job_dict).fillna(df_adp[src_job_col])
                if src_loc_col: df_uzio['Work Location'] = df_adp[src_loc_col].astype(str).str.strip().map(loc_dict).fillna(df_adp[src_loc_col])

                from utils.audit_utils import inject_into_uzio_template
                uzio_template_file.seek(0)
                wb = inject_into_uzio_template(df_uzio, uzio_template_file)
                out = io.BytesIO()
                wb.save(out)
                out.seek(0)

                st.success(summary)
                st.download_button("Download Updated Template", out.getvalue(), f"Uzio_Updated_ADP_{pd.Timestamp.now().strftime('%Y%m%d_%H%M')}.xlsm")
            except Exception as e:
                st.error(f"Error: {e}")

def render_ui():
    st.sidebar.title("Census Tools")
    tool = st.sidebar.selectbox("Select Tool", ["Sanity Check", "Full Generation", "Selective Sync"], key="adp_tool_select")
    if tool == "Sanity Check": render_census_sanity_check()
    elif tool == "Full Generation": render_census_generator()
    else: render_selective_census_generator()
