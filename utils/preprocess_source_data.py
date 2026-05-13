"""
Pre-processing auto-fix utility for Census Generators.
Applies opt-in corrections to source data AFTER sanity checks have been shown to the user.
The sanity check module (validate_source_data) is NOT modified.
"""
import re
import pandas as pd


def detect_fixable_issues(df, resolved_field_map):
    """
    Scan source data and count how many rows have fixable issues.
    Returns a dict of counts (0 means that fix category is not applicable).
    """
    counts = {
        'flsa_blank_count': 0,
        'email_blank_count': 0,
        'zip_fixable_count': 0,
        'hours_blank_count': 0,
        'hours_col_missing': False,
        'inactive_status_count': 0,
        'temporary_status_count': 0,
        'blank_dol_active_count': 0,
        'blank_dol_term_count': 0,
        'invalid_date_count': 0,
        'type_blank_count': 0
    }

    emp_id_col = resolved_field_map.get('Employee ID')
    type_col = resolved_field_map.get('Employment Type')
    pay_type_col = resolved_field_map.get('Pay Type')
    flsa_col = resolved_field_map.get('FLSA Classification')
    work_email_col = resolved_field_map.get('Work Email')
    personal_email_col = resolved_field_map.get('Personal Email')
    zip_col = resolved_field_map.get('Zip')
    hours_col = resolved_field_map.get('Working Hours')

    # Check if Working Hours column exists at all
    # NEW LOGIC: if not, check for "Standard Hours"
    std_hours_cand = None
    for cand in df.columns:
        if str(cand).strip().lower() == "standard hours":
            std_hours_cand = cand
            break

    if not hours_col or hours_col not in df.columns:
        if std_hours_cand:
            # We have Standard Hours. We will use that as base.
            # Only count blanks in Standard Hours as needing fixes.
            for _, row in df.iterrows():
                hrs_val = row.get(std_hours_cand)
                if pd.isna(hrs_val) or str(hrs_val).strip() == "" or str(hrs_val).strip().lower() == 'nan':
                    counts['hours_blank_count'] += 1
            # We will also need to rename/create the canonical column, which triggers the missing logic
            counts['hours_col_missing'] = True 
        else:
            counts['hours_col_missing'] = True
            counts['hours_blank_count'] = len(df)
    else:
        # Count blank working hours normally
        for _, row in df.iterrows():
            hrs_val = row.get(hours_col)
            if pd.isna(hrs_val) or str(hrs_val).strip() == "" or str(hrs_val).strip().lower() == 'nan':
                counts['hours_blank_count'] += 1

    # Count blank FLSA where Pay Type is set
    if flsa_col and flsa_col in df.columns and pay_type_col and pay_type_col in df.columns:
        for _, row in df.iterrows():
            flsa_val = row.get(flsa_col)
            pay_val = row.get(pay_type_col)
            flsa_is_blank = pd.isna(flsa_val) or str(flsa_val).strip() == ""
            pay_str = str(pay_val).strip().lower() if pd.notna(pay_val) and str(pay_val).strip() else ""
            if flsa_is_blank and pay_str and ("hourly" in pay_str or "hour" in pay_str or "salary" in pay_str or "salaried" in pay_str):
                counts['flsa_blank_count'] += 1

    # Count blank work emails where personal email exists
    if work_email_col and work_email_col in df.columns:
        for _, row in df.iterrows():
            we_val = row.get(work_email_col)
            if pd.isna(we_val) or str(we_val).strip() == "":
                if personal_email_col and personal_email_col in df.columns:
                    pe_val = row.get(personal_email_col)
                    if pd.notna(pe_val) and str(pe_val).strip():
                        counts['email_blank_count'] += 1

    # Count fixable zip codes
    if zip_col and zip_col in df.columns:
        for _, row in df.iterrows():
            zip_val = row.get(zip_col)
            if pd.notna(zip_val) and str(zip_val).strip():
                original = str(zip_val).strip()
                cleaned = original.split('-')[0].strip()
                cleaned = cleaned.split('.')[0].strip()
                digits_only = re.sub(r'[^0-9]', '', cleaned)
                if digits_only and digits_only != original:
                    counts['zip_fixable_count'] += 1
                    
    # Count Inactive Employee Statuses
    col_emp_status = None
    for cand in ['employee_status', 'employee status', 'employment status', 'status', 'ee status']:
        if cand in df.columns:
            col_emp_status = cand
            break
            
    if col_emp_status:
        for _, row in df.iterrows():
            val_emp = row.get(col_emp_status)
            if pd.notna(val_emp):
                if str(val_emp).strip().lower() == 'inactive':
                    counts['inactive_status_count'] += 1
                elif str(val_emp).strip().lower() == 'temporary':
                    counts['temporary_status_count'] += 1
                    
    col_dol = 'dol_status' if 'dol_status' in df.columns else None
    if col_dol and col_emp_status:
        for _, row in df.iterrows():
            val_dol = row.get(col_dol)
            if pd.isna(val_dol) or str(val_dol).strip() == "":
                emp_stat_str = str(row.get(col_emp_status)).strip().lower()
                if "term" in emp_stat_str:
                    counts['blank_dol_term_count'] += 1
                else:
                    counts['blank_dol_active_count'] += 1

    # Count invalid 00/00/0000 dates across all columns
    for col in df.columns:
        try:
            col_series = df.loc[:, col]
            if hasattr(col_series, 'astype') and col_series.astype(str).str.contains('00/00/0000', regex=False).any():
                counts['invalid_date_count'] += int(col_series.astype(str).str.count('00/00/0000').sum())
        except Exception:
            continue

    # Count blank Employment Type / Worker Category
    if type_col and type_col in df.columns:
        for _, row in df.iterrows():
            type_val = row.get(type_col)
            if pd.isna(type_val) or str(type_val).strip() == "":
                counts['type_blank_count'] += 1

    return counts


def apply_auto_fixes(df, resolved_field_map, fixes_to_apply=None):
    """
    Apply selected auto-corrections to the source DataFrame in-place.

    Args:
        df: Source DataFrame (mutated in-place)
        resolved_field_map: Dict mapping standard field names to resolved column names
        fixes_to_apply: Dict of booleans controlling which fixes to apply:
            {'fix_flsa': bool, 'fix_email': bool, 'fix_zip': bool, 'fix_hours': bool}
            If None, all fixes are applied.

    Returns:
        dict with keys: 'flsa_fills', 'email_fallbacks', 'zip_corrections', 'hours_fixes', 'inactive_fixes'
        Each value is a pd.DataFrame of corrections made.
    """
    if fixes_to_apply is None:
        fixes_to_apply = {
            'fix_flsa': True, 'fix_email': True, 'fix_zip': True, 'fix_hours': True, 
            'fix_inactive': True, 'fix_temporary': True, 
            'fix_blank_dol_active': True, 'fix_blank_dol_term': True,
            'fix_invalid_dates': True, 'fix_type_blanks': True
        }

    flsa_fills = []
    email_fallbacks = []
    zip_corrections = []
    hours_fixes = []
    inactive_fixes = []
    temporary_fixes = []
    dol_active_fixes = []
    dol_term_fixes = []
    invalid_date_fixes = []
    type_blank_fixes = []
    
    rows_to_drop = []

    # Resolve column references
    emp_id_col = resolved_field_map.get('Employee ID')
    type_col = resolved_field_map.get('Employment Type')
    pay_type_col = resolved_field_map.get('Pay Type')
    flsa_col = resolved_field_map.get('FLSA Classification')
    work_email_col = resolved_field_map.get('Work Email')
    personal_email_col = resolved_field_map.get('Personal Email')
    zip_col = resolved_field_map.get('Zip')
    hours_col = resolved_field_map.get('Working Hours')
    
    col_emp_status = None
    for cand in ['employee_status', 'employee status', 'employment status', 'status', 'ee status']:
        if cand in df.columns:
            col_emp_status = cand
            break

    def get_emp_ref(row, idx):
        ref = f"Row {idx + 2}"
        if emp_id_col and emp_id_col in df.columns:
            eid = row.get(emp_id_col)
            if pd.notna(eid) and str(eid).strip():
                ref = str(eid).strip()
        return ref

    # --- FIX: Working Hours (handle missing column first) ---
    if fixes_to_apply.get('fix_hours', False):
        # Always use a canonical column name for the fix
        canonical_hours_col = 'Working Hours Per Week'
        
        # Check for Standard Hours fallback
        std_hours_cand = None
        for cand in df.columns:
            if str(cand).strip().lower() == "standard hours":
                std_hours_cand = cand
                break
        
        if not hours_col or hours_col not in df.columns:
            if std_hours_cand:
                # Fallback to Standard Hours
                df[canonical_hours_col] = df[std_hours_cand].copy()
                resolved_field_map['Working Hours'] = canonical_hours_col
                hours_col = canonical_hours_col
                # Don't add to hours_fixes here; let the row-by-row blank check catch any missing values next.
            else:
                # Column is missing entirely and no Standard Hours — add it with "0" values
                df[canonical_hours_col] = "0"
                resolved_field_map['Working Hours'] = canonical_hours_col
                hours_col = canonical_hours_col
                hours_fixes.append({
                    'Employee ID': '(All Employees)',
                    'Original Hours': '(Column Missing)',
                    'Corrected Hours': '0'
                })
        else:
            # Column exists but may have blank values — create a new canonical column
            # Copy existing values, then fill blanks with 0
            df[canonical_hours_col] = df[hours_col].copy()
            resolved_field_map['Working Hours'] = canonical_hours_col
            hours_col = canonical_hours_col

    for idx, row in df.iterrows():
        emp_ref = get_emp_ref(row, idx)

        # --- FIX 1: Blank FLSA Classification → fill based on Pay Type ---
        if fixes_to_apply.get('fix_flsa', False):
            if (flsa_col and flsa_col in df.columns
                    and pay_type_col and pay_type_col in df.columns):
                flsa_val = row.get(flsa_col)
                pay_val = row.get(pay_type_col)

                flsa_is_blank = pd.isna(flsa_val) or str(flsa_val).strip() == ""
                pay_str = str(pay_val).strip().lower() if pd.notna(pay_val) and str(pay_val).strip() else ""

                if flsa_is_blank and pay_str:
                    if "hourly" in pay_str or "hour" in pay_str:
                        df.at[idx, flsa_col] = "Non-Exempt"
                        flsa_fills.append({
                            'Employee ID': emp_ref,
                            'Pay Type': str(pay_val).strip(),
                            'Assigned FLSA': 'Non-Exempt'
                        })
                    elif "salary" in pay_str or "salaried" in pay_str:
                        df.at[idx, flsa_col] = "Exempt"
                        flsa_fills.append({
                            'Employee ID': emp_ref,
                            'Pay Type': str(pay_val).strip(),
                            'Assigned FLSA': 'Exempt'
                        })

        # --- FIX 2: Blank Work Email → fill from Personal Email ---
        if fixes_to_apply.get('fix_email', False):
            if work_email_col and work_email_col in df.columns:
                we_val = row.get(work_email_col)
                if pd.isna(we_val) or str(we_val).strip() == "":
                    if personal_email_col and personal_email_col in df.columns:
                        pe_val = row.get(personal_email_col)
                        if pd.notna(pe_val) and str(pe_val).strip():
                            df.at[idx, work_email_col] = str(pe_val).strip()
                            email_fallbacks.append({
                                'Employee ID': emp_ref,
                                'Personal Email Used': str(pe_val).strip()
                            })

        # --- FIX 3: Zip Code Normalization ---
        if fixes_to_apply.get('fix_zip', False):
            if zip_col and zip_col in df.columns:
                zip_val = row.get(zip_col)
                if pd.notna(zip_val) and str(zip_val).strip():
                    original_zip = str(zip_val).strip()
                    cleaned = original_zip.split('-')[0].strip()
                    cleaned = cleaned.split('.')[0].strip()
                    digits_only = re.sub(r'[^0-9]', '', cleaned)
                    if digits_only and digits_only != original_zip:
                        df.at[idx, zip_col] = digits_only
                        zip_corrections.append({
                            'Employee ID': emp_ref,
                            'Original Zip': original_zip,
                            'Corrected Zip': digits_only
                        })

        # --- FIX 4: Blank Working Hours → set to 0 ---
        if fixes_to_apply.get('fix_hours', False):
            if hours_col and hours_col in df.columns:
                hrs_val = row.get(hours_col)
                if pd.isna(hrs_val) or str(hrs_val).strip() == "" or str(hrs_val).strip().lower() == 'nan':
                    df.at[idx, hours_col] = "0"
                    hours_fixes.append({
                        'Employee ID': emp_ref,
                        'Original Hours': str(hrs_val).strip() if pd.notna(hrs_val) else '(blank)',
                        'Corrected Hours': '0'
                    })

        # --- FIX 5: Inactive Status → Terminated ---
        if fixes_to_apply.get('fix_inactive', False):
            if col_emp_status:
                val_emp = row.get(col_emp_status)
                if pd.notna(val_emp) and str(val_emp).strip().lower() == 'inactive':
                    df.at[idx, col_emp_status] = 'Terminated'
                    inactive_fixes.append({
                        'Employee ID': emp_ref,
                        'Original Status': str(val_emp).strip(),
                        'Corrected Status': 'Terminated'
                    })

        # --- FIX 6: Temporary Status → Seasonal ---
        if fixes_to_apply.get('fix_temporary', False):
            if col_emp_status:
                val_emp = row.get(col_emp_status)
                if pd.notna(val_emp) and str(val_emp).strip().lower() == 'temporary':
                    df.at[idx, col_emp_status] = 'Seasonal'
                    temporary_fixes.append({
                        'Employee ID': emp_ref,
                        'Original Status': str(val_emp).strip(),
                        'Corrected Status': 'Seasonal'
                    })
                    
        # --- FIX 7: Blank DOL Status ---
        col_dol = 'dol_status' if 'dol_status' in df.columns else None
        if col_dol and col_emp_status:
            val_dol = row.get(col_dol)
            if pd.isna(val_dol) or str(val_dol).strip() == "":
                emp_stat_str = str(row.get(col_emp_status)).strip().lower()
                if "term" in emp_stat_str:
                    if fixes_to_apply.get('fix_blank_dol_term', False):
                        if idx not in rows_to_drop:
                            rows_to_drop.append(idx)
                        dol_term_fixes.append({
                            'Employee ID': emp_ref,
                            'Original DOL': '(blank)',
                            'Action': 'Deleted Row (Terminated Employee)'
                        })
                else:
                    if fixes_to_apply.get('fix_blank_dol_active', False):
                        df.at[idx, col_dol] = 'Full-Time'
                        dol_active_fixes.append({
                            'Employee ID': emp_ref,
                            'Original DOL': '(blank)',
                            'Action': 'Set to Full-Time (Active Employee)'
                        })

        # --- FIX 8: Invalid Dates (00/00/0000) ---
        if fixes_to_apply.get('fix_invalid_dates', False):
            for col in df.columns:
                try:
                    val = row.get(col)
                    if pd.notna(val) and '00/00/0000' in str(val):
                        fixed_val = str(val).replace('00/00/0000', '').strip()
                        df.at[idx, col] = fixed_val
                        invalid_date_fixes.append({
                            'Employee ID': emp_ref,
                            'Column': col,
                            'Original Value': str(val),
                            'Action': 'Blanked invalid date'
                        })
                except Exception:
                    continue

        # --- FIX 9: Blank Worker Category / Employment Type ---
        if fixes_to_apply.get('fix_type_blanks', False):
            if type_col and type_col in df.columns:
                type_val = row.get(type_col)
                if pd.isna(type_val) or str(type_val).strip() == "":
                    df.at[idx, type_col] = 'Part Time'
                    type_blank_fixes.append({
                        'Employee ID': emp_ref,
                        'Original Worker Category': '(blank)',
                        'Corrected Value': 'Part Time'
                    })

    if rows_to_drop:
        df.drop(list(set(rows_to_drop)), inplace=True)
        # We don't necessarily reset_index because it might mess up other references if they existed,
        # but the loop is over so it's fine. 
        df.reset_index(drop=True, inplace=True)

    return {
        'flsa_fills': pd.DataFrame(flsa_fills),
        'email_fallbacks': pd.DataFrame(email_fallbacks),
        'zip_corrections': pd.DataFrame(zip_corrections),
        'hours_fixes': pd.DataFrame(hours_fixes),
        'inactive_fixes': pd.DataFrame(inactive_fixes),
        'temporary_fixes': pd.DataFrame(temporary_fixes),
        'dol_active_fixes': pd.DataFrame(dol_active_fixes),
        'dol_term_fixes': pd.DataFrame(dol_term_fixes),
        'invalid_date_fixes': pd.DataFrame(invalid_date_fixes),
        'type_blank_fixes': pd.DataFrame(type_blank_fixes)
    }
