import streamlit as st
import pandas as pd
import io
from utils.audit_utils import check_duplicate_columns, format_datetime_strings, convert_state_to_abbreviation

def render_employee_extractor():
    st.title("Selective Employee Extractor (Selective Sync & Sequence)")
    st.markdown("""
    **Purpose**: Extract a specific subset of employees from any census file while maintaining **100% data integrity** and **controlling the sequence**.
    
    **Features**:
    - **Sequencing**: Re-order rows to match an Uzio Census or a custom list.
    - **Column Selection**: Pick only the fields you need (e.g., ID + License Expiry).
    - **Zero Tampering**: Preserves original formats/leading zeros.
    """)
    
    # 1. FILE UPLOADERS
    col_u1, col_u2 = st.columns(2)
    with col_u1:
        source_file = st.file_uploader(
            "1. Upload SOURCE File",
            type=["xlsx", "csv", "xlsm"],
            key="ee_source",
            help="Supported: ADP Census, Paycom Census, ADP Direct Deposit, **Uzio Multi-Client Census** (Employee Details sheet), **ADP Emergency & License Report**"
        )
    with col_u2:
        ref_file = st.file_uploader("2. Upload REFERENCE Order (Uzio Census) - OPTIONAL", type=["xlsx", "xlsm"], key="ee_ref")

    if not source_file:
        st.info("Please upload a source file to begin.")
        return

    # 2. READ SOURCE (Strict no-mutation)
    # Track detected file type for informational banners
    _detected_source_type = None

    try:
        source_file.seek(0)
        if source_file.name.lower().endswith('.csv'):
            df_source = pd.read_csv(source_file, dtype=str)
        else:
            # --- Peek at available sheets ---
            import warnings
            with warnings.catch_warnings():
                warnings.simplefilter('ignore')
                source_file.seek(0)
                _xl_peek = pd.ExcelFile(source_file)
                _available_sheets = _xl_peek.sheet_names

            # CASE 1: Uzio Multi-Client Census — has 'Employee Details' sheet
            if 'Employee Details' in _available_sheets:
                source_file.seek(0)
                df_source = pd.read_excel(source_file, sheet_name='Employee Details', header=3, dtype=str)
                _detected_source_type = 'uzio'

            else:
                # CASE 2: Generic Excel — read first sheet with auto header detection
                _target_sheet = _available_sheets[0] if _available_sheets else 0
                source_file.seek(0)
                df_header = pd.read_excel(source_file, sheet_name=_target_sheet, nrows=10, header=None)
                header_idx = 0
                _HEADER_KEYWORDS = [
                    'associate id', 'employee_code', 'employee id*', 'employee id',
                    'legal first name', 'name', 'company code',
                    'license/certification description',  # Emergency & License Report
                    'contact name',                       # Emergency & License Report
                    'lived in state tax code',            # SIT/FIT Report
                    'worked in state description',        # SIT/FIT Report
                ]
                for idx, row in df_header.iterrows():
                    row_vals = [str(x).lower().strip() for x in row.tolist() if pd.notna(x)]
                    if any(k in row_vals for k in _HEADER_KEYWORDS):
                        header_idx = idx
                        break
                source_file.seek(0)
                df_source = pd.read_excel(source_file, sheet_name=_target_sheet, header=header_idx, dtype=str)

                # Detect Emergency & License Report
                _emer_lic_cols = {'License/Certification Description', 'License/Certification ID', 'Issued By', 'Expiration Date'}
                if _emer_lic_cols.issubset(set(df_source.columns.tolist())):
                    _detected_source_type = 'emergency_license'
                
                # Detect SIT/FIT Withholding Report
                _sit_fit_cols = {'Lived In State Tax Code', 'Worked in State Description', 'Worked in State Code'}
                if _sit_fit_cols.issubset(set(df_source.columns.tolist())):
                    _detected_source_type = 'sit_fit'

    except Exception as e:
        st.error(f"Error reading source file: {e}")
        return

    # Banner based on detected file type
    if _detected_source_type == 'uzio':
        st.success(f"✅ **Uzio Multi-Client Census detected** — 'Employee Details' sheet loaded: {len(df_source)} rows, {len(df_source.columns)} columns.")
    elif _detected_source_type == 'emergency_license':
        st.success(f"✅ **ADP Emergency & License Report detected** — {len(df_source)} rows, {len(df_source.columns)} columns. Note: one row per license record (employees may repeat).")
    elif _detected_source_type == 'sit_fit':
        st.success(f"✅ **ADP SIT/FIT Withholding Report detected** — {len(df_source)} rows, {len(df_source.columns)} columns. Note: one row per tax config (employees may repeat).")
    else:
        st.success(f"✅ Source file loaded: {len(df_source)} rows, {len(df_source.columns)} columns.")


    # 3. IDENTIFY ID COLUMN (Source)
    id_col_source = None
    all_cols = df_source.columns.tolist()
    # Candidates cover: ADP Census, Paycom, Uzio, ADP Direct Deposit (ASSOCIATE ID - uppercase)
    candidates = ['ASSOCIATE ID', 'Associate ID', 'Employee_Code', ' Employee ID*', 'Employee ID', 'Employee Code', 'EE ID']
    for cand in candidates:
        if cand in all_cols:
            id_col_source = cand
            break
    if not id_col_source:
        # Fuzzy match
        for col in all_cols:
            c_norm = str(col).lower().strip().replace('*', '').replace(' ', '_')
            if c_norm in ['associate_id', 'employee_id', 'employee_code', 'ee_id', 'eid', 'associateid', 'eeid']:
                id_col_source = col
                break
    
    if not id_col_source:
        st.error("Could not identify 'Employee ID' column in source. Headers found: " + ", ".join(all_cols[:10]))
        return

    # Detect multi-row-per-employee scenarios
    is_direct_deposit = 'ROUTING NUMBER' in all_cols or 'ACCOUNT NUMBER' in all_cols
    if is_direct_deposit:
        st.info("📋 **ADP Direct Deposit report detected.** Multiple rows per employee (split accounts) will all be included in the output.")
    if _detected_source_type == 'emergency_license':
        st.info("📋 **Emergency & License Report mode.** An employee may have multiple rows (one per license). All matching rows will be preserved in the output.")
    if _detected_source_type == 'sit_fit':
        st.info("📋 **ADP SIT/FIT Report mode.** An employee may have multiple rows (one per state tax config). All matching rows will be preserved in the output.")

    # 4. COLUMN SELECTOR
    st.markdown("---")
    st.markdown("### 2. Choose Columns to Include")
    include_all = st.checkbox("Include ALL columns", value=False)
    selected_cols = []
    if include_all:
        selected_cols = all_cols
    else:
        selected_cols = st.multiselect("Select columns from source", all_cols, default=[id_col_source] if id_col_source in all_cols else [])

    if not selected_cols:
        st.warning("Please select at least one column.")
        return

    # 4b. INCLUSION OPTIONS
    include_remaining = st.checkbox("Include remaining Source employees in a separate tab?", value=False, help="If checked, employees in the source but NOT in the reference will be added to Sheet 2.")

    # 5. DEFINE SEQUENCE
    st.markdown("---")
    st.markdown("### 3. Define Employee Sequence")
    
    ordered_ids = []
    ref_ids_set = set()
    if ref_file:
        try:
            # Attempt to read Uzio Multi-Client Template
            # Usually 'Employee Details' sheet, Header row 4 (index 3)
            df_ref = pd.read_excel(ref_file, sheet_name='Employee Details', header=3, dtype=str)
            ref_id_col = ' Employee ID*'
            if ref_id_col in df_ref.columns:
                ordered_ids = df_ref[ref_id_col].dropna().unique().tolist()
                ref_ids_set = set(ordered_ids)
                st.info(f"Loaded **{len(ordered_ids)}** IDs from Uzio Reference (Order strictly matched).")
            else:
                st.error("Reference file uploaded but ' Employee ID*' column not found in 'Employee Details' sheet.")
        except Exception as e:
            st.error(f"Error reading Reference file: {e}. Ensure it is a valid Uzio Census Template.")
    
    # Manual Input (Fallback or Hybrid)
    manual_ids_input = st.text_area("Paste Employee IDs (Comma-separated) - Use this if no reference file or to override", 
                                   height=100, 
                                   help="IDs provided here will be used in exactly this order.")
    if manual_ids_input.strip():
        manual_ids = [i.strip() for i in manual_ids_input.split(',') if i.strip()]
        if ordered_ids:
            st.warning("Both Reference File and Manual IDs provided. **Using Manual List** for final sequence.")
        ordered_ids = manual_ids

    # --- PROGRESSIVE ID MATCHING HELPERS ---
    def _strip_separators(val):
        """Level 2: strip hyphens and spaces only."""
        return str(val).strip().replace('-', '').replace(' ', '')
    
    def _strip_all(val):
        """Level 3: strip hyphens, spaces, AND leading zeros."""
        s = _strip_separators(val)
        return s.lstrip('0') or '0'

    # Try to find a name column for collision alerts
    name_col = None
    for c in all_cols:
        cl = str(c).lower().strip()
        if cl in ['name', 'legal first name', 'legal_firstname', 'first name']:
            name_col = c
            break

    # --- ID MISMATCH FLAGGING (Source vs Reference) ---
    if ref_ids_set and not manual_ids_input.strip():
        source_ids_raw = set(df_source[id_col_source].astype(str).str.strip().dropna().unique())
        source_ids_l2 = set(_strip_separators(x) for x in source_ids_raw)
        ref_ids_l2 = set(_strip_separators(x) for x in ref_ids_set)
        
        in_ref_not_source = sorted([x for x in ref_ids_set if _strip_separators(x) not in source_ids_l2])
        in_source_not_ref = sorted([x for x in source_ids_raw if _strip_separators(x) not in ref_ids_l2])
        
        if in_ref_not_source or in_source_not_ref:
            st.warning(f"⚠️ **ID Mismatch Detected** — Sequencing may be affected!")
            if in_ref_not_source:
                with st.expander(f"🟡 {len(in_ref_not_source)} ID(s) in Uzio Reference but MISSING from Source", expanded=False):
                    st.markdown("These employees exist in your Uzio template but were **not found** in the uploaded source file. They will be **skipped** in the output.")
                    st.dataframe(pd.DataFrame({"Missing Employee ID": in_ref_not_source}), hide_index=True, use_container_width=True)
            if in_source_not_ref:
                with st.expander(f"🔵 {len(in_source_not_ref)} ID(s) in Source but MISSING from Uzio Reference", expanded=False):
                    st.markdown("These employees exist in your source file but are **not listed** in the Uzio reference. They will be **excluded** from the output since sequencing follows the reference order.")
                    st.dataframe(pd.DataFrame({"Extra Employee ID": in_source_not_ref}), hide_index=True, use_container_width=True)
        else:
            st.success("✅ All IDs match perfectly between Source and Uzio Reference!")

    if not ordered_ids:
        st.info("Waiting for Reference File or Manual ID list to define the sequence...")
        return

    # 6. EXTRACTION LOGIC — 3-Level Progressive Matching with Placeholders
    df_source[id_col_source] = df_source[id_col_source].astype(str).str.strip()
    source_id_set = set(df_source[id_col_source].tolist())
    
    # Build Level 2 & Level 3 lookup indexes: normalized -> [list of original source IDs]
    from collections import defaultdict
    l2_index = defaultdict(set)  # separator_stripped -> original IDs
    l3_index = defaultdict(set)  # fully-stripped -> original IDs
    for sid in source_id_set:
        l2_index[_strip_separators(sid)].add(sid)
        l3_index[_strip_all(sid)].add(sid)
    
    # Process each user-entered ID through 3 levels
    tab1_frames = []         # DataFrames to concat for Tab 1
    ids_matched_in_source = set() # To identify leftovers
    collision_alerts = []    # IDs that had collisions
    unmatched_ids = []       # IDs that could not be matched
    
    for pos, user_id in enumerate(ordered_ids):
        user_id_clean = user_id.strip()
        matched_orig_id = None
        
        # LEVEL 1: Exact Match
        if user_id_clean in source_id_set:
            matched_orig_id = user_id_clean
        else:
            # LEVEL 2: Strip Hyphens/Spaces
            l2_key = _strip_separators(user_id_clean)
            l2_candidates = l2_index.get(l2_key, set())
            if len(l2_candidates) == 1:
                matched_orig_id = list(l2_candidates)[0]
            elif len(l2_candidates) > 1:
                collision_alerts.append((user_id_clean, 'hyphens/spaces', l2_candidates))
                continue
            else:
                # LEVEL 3: Strip Leading Zeros
                l3_key = _strip_all(user_id_clean)
                l3_candidates = l3_index.get(l3_key, set())
                if len(l3_candidates) == 1:
                    matched_orig_id = list(l3_candidates)[0]
                elif len(l3_candidates) > 1:
                    collision_alerts.append((user_id_clean, 'leading zeros', l3_candidates))
                    continue
        
        if matched_orig_id:
            # Found in source — pull all matching rows
            matching_rows = df_source[df_source[id_col_source] == matched_orig_id].copy()
            tab1_frames.append(matching_rows)
            ids_matched_in_source.add(matched_orig_id)
        else:
            # Placeholder row for unmatched Reference ID
            placeholder_data = {col: "" for col in df_source.columns}
            placeholder_data[id_col_source] = user_id_clean
            tab1_frames.append(pd.DataFrame([placeholder_data]))
            unmatched_ids.append(user_id_clean)
    
    # Leftovers for Tab 2
    df_remaining = df_source[~df_source[id_col_source].isin(ids_matched_in_source)].copy()
    
    # --- COLLISION ALERTS ---
    if collision_alerts:
        st.error(f"🔴 **{len(collision_alerts)} ID(s) have collisions** — Please provide the exact ID from the source file.")
        for user_id, level, candidates in collision_alerts:
            with st.expander(f"⚠️ Collision for '{user_id}' (ambiguous after stripping {level})", expanded=True):
                st.markdown(f"You entered **`{user_id}`** but multiple employees in the source file normalize to the same value. **Which one did you mean?**")
                # Build a detail table with names if available
                detail_rows = []
                for cand_id in sorted(candidates):
                    row_data = {"Source Employee ID (use this exact value)": cand_id}
                    if name_col:
                        name_matches = df_source[df_source[id_col_source] == cand_id][name_col].dropna().unique()
                        row_data["Employee Name"] = name_matches[0] if len(name_matches) > 0 else "—"
                    detail_rows.append(row_data)
                st.dataframe(pd.DataFrame(detail_rows), hide_index=True, use_container_width=True)
                st.info("👆 Copy the **exact Employee ID** from above and paste it in the manual input box to resolve this collision.")
    
    if unmatched_ids:
        with st.expander(f"🔻 {len(unmatched_ids)} ID(s) not found in source at any level", expanded=False):
            st.dataframe(pd.DataFrame({"Unmatched ID": unmatched_ids}), hide_index=True, use_container_width=True)
    
    # Build df_result from frames
    df_result = pd.concat(tab1_frames, ignore_index=True) if tab1_frames else pd.DataFrame(columns=df_source.columns)
    
    # Final column subset
    df_result = df_result[selected_cols]

    # 7. DATA CLEANING
    for col in df_result.columns:
        # Clear 00/00/0000 dates (Paycom exception)
        if df_result[col].astype(str).str.contains('00/00/0000', regex=False).any():
            df_result[col] = df_result[col].replace('00/00/0000', '')

    # Auto-detect and format all date-like columns to MM/DD/YYYY
    date_keywords = ['date', 'dob', 'birth', 'hire', 'termination', 'expir', 'expiration']
    date_like_cols = [
        col for col in df_result.columns
        if any(kw in str(col).lower() for kw in date_keywords)
    ]
    if date_like_cols:
        df_result = format_datetime_strings(df_result, date_like_cols)

    # Convert full state names to 2-letter abbreviations (License/Emergency/SIT-FIT reports)
    for state_col in ['Issued By', 'Lived In State Description', 'Worked in State Description']:
        if state_col in df_result.columns:
            df_result = convert_state_to_abbreviation(df_result, state_col)

    # Apply formatting and columns to remaining employees
    df_remaining_out = pd.DataFrame()
    if include_remaining and not df_remaining.empty:
        df_remaining_out = df_remaining[selected_cols].copy()
        if date_like_cols:
            df_remaining_out = format_datetime_strings(df_remaining_out, date_like_cols)
        for state_col in ['Issued By', 'Lived In State Description', 'Worked in State Description']:
            if state_col in df_remaining_out.columns:
                df_remaining_out = convert_state_to_abbreviation(df_remaining_out, state_col)

    # 8. RESULTS & DOWNLOAD
    st.markdown("---")
    if df_result.empty:
        st.warning("No employees found matching the sequence criteria.")
    else:
        st.success(f"Matched **{len(df_result)}** employees in the specified sequence.")
        st.dataframe(df_result.head(100), use_container_width=True, hide_index=True)
        
        col1, col2 = st.columns(2)
        # Excel
        buffer_xlsx = io.BytesIO()
        with pd.ExcelWriter(buffer_xlsx, engine='openpyxl') as writer:
            df_result.to_excel(writer, index=False, sheet_name="Sync Result")
            if include_remaining and not df_remaining_out.empty:
                df_remaining_out.to_excel(writer, index=False, sheet_name="Remaining IDs")
        buffer_xlsx.seek(0)
        col1.download_button("📥 Download Excel", buffer_xlsx.getvalue(), f"Selective_Census_{pd.Timestamp.now().strftime('%Y%m%d')}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
        # CSV
        buffer_csv = io.StringIO()
        df_result.to_csv(buffer_csv, index=False)
        col2.download_button("📄 Download CSV", buffer_csv.getvalue(), f"Selective_Census_{pd.Timestamp.now().strftime('%Y%m%d')}.csv", "text/csv")
