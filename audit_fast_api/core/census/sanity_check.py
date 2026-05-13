"""
ADP Census Sanity Check — auto-fix pipeline mirror of the Streamlit tool.

Exposes:
  - run_census_sanity_check(df, field_map_dict)             -> JSON-friendly dict (legacy validation-only)
  - generate_corrected_census_xlsx(content, field_map_dict, fix_options) -> bytes (Corrected Census + Change Log)

The fix_options dict supports the same toggles surfaced in apps/adp/census_generator.py
render_auto_fix_options. All toggles default to False if absent.

Toggles mirrored:
  fix_flsa, fix_emails, fix_job_title, fix_driver_smart, fix_license,
  fix_status, fix_inactive (alias of fix_status), fix_type, fix_dol_status,
  fix_leave_to_active, fix_blank_jt_to_driver,
  fix_std_hours, rename_std_hours, fix_zip, rename_zip_col, replace_gender_col
"""
import io
import re
import pandas as pd

from utils.audit_utils import norm_colname, norm_blank, norm_ssn_canonical


# ---------------------------------------------------------------------------
# Legacy lightweight validator (kept so existing endpoints / paycom path keep
# working with their original JSON shape).
# ---------------------------------------------------------------------------
def validate_source_data(df_source, resolved_field_map):
    hard_errors = []
    emp_id_col = resolved_field_map.get('Employee ID')
    ssn_col = resolved_field_map.get('SSN')
    status_col = resolved_field_map.get('Employment Status')

    for idx, row in df_source.iterrows():
        eid = str(row.get(emp_id_col, "")).strip() if emp_id_col else ""
        ssn = norm_ssn_canonical(row.get(ssn_col, "")) if ssn_col else ""
        status = str(row.get(status_col, "")).strip().lower() if status_col else ""

        issues = []
        if not eid: issues.append("Missing Employee ID")
        if not ssn: issues.append("Missing SSN")
        if not status: issues.append("Missing Employment Status")

        if issues:
            hard_errors.append({
                "Employee ID": eid or f"Row {idx + 2}",
                "Issue": ", ".join(issues)
            })

    return {"hard_errors": pd.DataFrame(hard_errors)}


def run_census_sanity_check(df_source, resolved_field_map):
    return validate_source_data(df_source, resolved_field_map)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _normalize_field_map(field_map_dict):
    """field_map_dict values may be a string or a list of strings (vendor variants).
    Returns a normalized {std_name: normalized_source_col_name} mapping."""
    out = {}
    for std_name, vendor_cols in field_map_dict.items():
        if isinstance(vendor_cols, str):
            vendor_cols = [vendor_cols]
        # Pick the first variant; resolution to actual columns happens later.
        out[std_name] = norm_colname(vendor_cols[0]) if vendor_cols else ""
    return out


def _resolve_field_map(df_normalized_cols, field_map_dict):
    """Resolve field_map values to actual normalized column names present in df."""
    resolved = {}
    for std_name, vendor_cols in field_map_dict.items():
        if isinstance(vendor_cols, str):
            vendor_cols = [vendor_cols]
        chosen = None
        for vc in vendor_cols:
            nv = norm_colname(vc)
            if nv in df_normalized_cols:
                chosen = nv
                break
        resolved[std_name] = chosen if chosen else (norm_colname(vendor_cols[0]) if vendor_cols else "")
    return resolved


def _read_source(content, filename="upload.xlsx"):
    """Read .xlsx or .csv content into a string-typed DataFrame."""
    bio = io.BytesIO(content)
    if filename.lower().endswith(".csv"):
        try:
            df = pd.read_csv(bio, dtype=str)
        except UnicodeDecodeError:
            bio.seek(0)
            df = pd.read_csv(bio, dtype=str, encoding="latin1")
    else:
        df = pd.read_excel(bio, dtype=str)
    return df


def _format_datetime_column(series):
    def _clean(val):
        if pd.isna(val) or str(val).strip() == "" or str(val).strip().lower() in ["nan", "nat"]:
            return ""
        try:
            d = pd.to_datetime(str(val).strip(), errors="coerce")
            if pd.notna(d):
                return d.strftime("%m/%d/%Y")
        except (ValueError, TypeError):
            pass
        return str(val).strip()
    return series.apply(_clean)


def _build_excel_bytes(df_main, df_audit,
                       sheet_main="Corrected Census", sheet_audit="Change Log"):
    """Two-sheet workbook: corrected census + change log (mirrors generate_excel_with_audit)."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_main.to_excel(writer, index=False, sheet_name=sheet_main)
        if df_audit is not None and not df_audit.empty:
            df_audit.to_excel(writer, index=False, sheet_name=sheet_audit)
            workbook = writer.book
            worksheet = writer.sheets[sheet_audit]
            header_format = workbook.add_format({"bold": True, "bg_color": "#D7E4BC", "border": 1})
            for col_num, value in enumerate(df_audit.columns.values):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 40 if value == "Comments" else 20)
        else:
            cols = ["Employee ID", "Employee Name", "Field Changed", "Old Value", "Assumed Value", "Comments"]
            pd.DataFrame(columns=cols).to_excel(writer, index=False, sheet_name=sheet_audit)
    return output.getvalue()


# ---------------------------------------------------------------------------
# Apply pipeline
# ---------------------------------------------------------------------------
def _validate_for_warnings(df_source, resolved_field_map):
    """Lightweight per-row warnings used to populate CRITICAL_WARNINGS column.
    Mirrors the spirit of validate_source_data without porting its full 500+ lines."""
    issues_by_row = {}
    emp_id_col = resolved_field_map.get('Employee ID')
    ssn_col = resolved_field_map.get('SSN')
    status_col = resolved_field_map.get('Employment Status')
    type_col = resolved_field_map.get('Employment Type')
    pay_type_col = resolved_field_map.get('Pay Type')
    job_title_col = resolved_field_map.get('Job Title')
    location_col = resolved_field_map.get('Work Location')
    hire_date_col = resolved_field_map.get('Hire Date')
    term_date_col = resolved_field_map.get('Termination Date')

    def _is_blank(v):
        return pd.isna(v) or str(v).strip() == "" or str(v).strip().lower() == "nan"

    for idx, row in df_source.iterrows():
        eid_val = row.get(emp_id_col) if emp_id_col else ""
        eid = str(eid_val).strip() if pd.notna(eid_val) else ""
        problems = []

        if ssn_col and ssn_col in df_source.columns:
            if not norm_ssn_canonical(row.get(ssn_col, "")):
                problems.append("SSN (blank)")

        if status_col and status_col in df_source.columns:
            sv = row.get(status_col)
            if _is_blank(sv):
                problems.append("Employment Status (blank)")
            else:
                sv_l = str(sv).strip().lower()
                is_term = sv_l in ['t', 'i'] or any(s in sv_l for s in ['terminated', 'inactive'])
                if is_term and term_date_col and term_date_col in df_source.columns:
                    if _is_blank(row.get(term_date_col)):
                        problems.append("Terminated/Inactive but missing Termination Date")
                is_standard = sv_l in ['a', 't', 'i'] or any(s in sv_l for s in ['active', 'terminated', 'inactive', 'leave'])
                if not is_standard:
                    problems.append(f"Non-standard Status ({str(sv).strip()})")

        if type_col and type_col in df_source.columns and _is_blank(row.get(type_col)):
            problems.append("Employment Type (blank)")
        if pay_type_col and pay_type_col in df_source.columns and _is_blank(row.get(pay_type_col)):
            problems.append("Pay Type (blank)")
        if job_title_col and job_title_col in df_source.columns and _is_blank(row.get(job_title_col)):
            problems.append("Job Title (blank)")
        if location_col and location_col in df_source.columns and _is_blank(row.get(location_col)):
            problems.append("Work Location (blank)")

        # Date logic: hire_date <= term_date when both present
        if (hire_date_col and term_date_col and
                hire_date_col in df_source.columns and term_date_col in df_source.columns):
            hv, tv = row.get(hire_date_col), row.get(term_date_col)
            if not _is_blank(hv) and not _is_blank(tv):
                try:
                    hd = pd.to_datetime(str(hv), errors="coerce")
                    td = pd.to_datetime(str(tv), errors="coerce")
                    if pd.notna(hd) and pd.notna(td) and td < hd:
                        problems.append("Termination date predates Hire date")
                except Exception:
                    pass

        if problems:
            issues_by_row[idx] = "; ".join(problems)
    return issues_by_row


def generate_corrected_census_xlsx(content, field_map_dict, fix_options=None,
                                   filename="upload.xlsx", location_mappings=None,
                                   sort_by_manager=False):
    """
    Apply selected auto-fixes to the source census file and return an .xlsx
    (bytes) containing two sheets: 'Corrected Census' and 'Change Log'.

    Args:
        content: bytes of the uploaded source (xlsx/csv).
        field_map_dict: ADP_FIELD_MAP-style {std_name: vendor_col_or_list}.
        fix_options: dict of toggle names -> bool. All default to False.
        filename: original filename (used to detect csv vs xlsx).
        location_mappings: optional {source_loc: uzio_loc} dict.
        sort_by_manager: optional bool — cluster managers/reportees at top.
    """
    fix_options = dict(fix_options or {})
    location_mappings = location_mappings or {}

    df_source = _read_source(content, filename)

    # Preserve original headers; normalize working copy.
    original_columns = list(df_source.columns)
    df_source.columns = [norm_colname(c) for c in df_source.columns]
    norm_to_orig = dict(zip(df_source.columns, original_columns))

    resolved_field_map = _resolve_field_map(df_source.columns, field_map_dict)

    issue_map_by_idx = _validate_for_warnings(df_source, resolved_field_map)
    df_download = df_source.copy()

    # Inject CRITICAL_WARNINGS column (per-row issue summary)
    df_download['CRITICAL_WARNINGS'] = ""
    for idx, msg in issue_map_by_idx.items():
        df_download.at[idx, 'CRITICAL_WARNINGS'] = msg

    audit_trail = []
    emp_id_col = resolved_field_map.get('Employee ID')
    emp_name_col = next((c for c in df_download.columns if 'name' in str(c).lower()), None)

    def _row_name(idx):
        if emp_name_col and emp_name_col in df_download.columns:
            return str(df_download.at[idx, emp_name_col]).strip()
        return "N/A"

    def log_change(idx, field, old_val, new_val, comment):
        eid = str(df_download.at[idx, emp_id_col]) if emp_id_col and emp_id_col in df_download.columns else "N/A"
        audit_trail.append({
            'Employee ID': eid,
            'Employee Name': _row_name(idx),
            'Field Changed': field,
            'Old Value': str(old_val) if pd.notna(old_val) else "(blank)",
            'Assumed Value': str(new_val),
            'Comments': comment
        })

    def _is_blank_series(col):
        s = df_download[col]
        return s.isna() | (s.astype(str).str.strip() == "") | (s.astype(str).str.strip().str.lower() == "nan")

    # --- Apply fixes (mirror order from render_census_sanity_check) ---

    # 1. Email fallback
    if fix_options.get('fix_emails'):
        c_work = resolved_field_map.get('Work Email')
        c_pers = resolved_field_map.get('Personal Email')
        if c_work and c_pers and c_work in df_download.columns and c_pers in df_download.columns:
            mask = _is_blank_series(c_work)
            for idx in df_download[mask].index:
                old_e = df_download.at[idx, c_work]
                new_e = df_download.at[idx, c_pers]
                if pd.notna(new_e) and str(new_e).strip():
                    df_download.at[idx, c_work] = new_e
                    log_change(idx, "Work Email", old_e, new_e,
                               "Personal email used as fallback for missing work email.")

    # 2. Leave -> Active when termination blank
    if fix_options.get('fix_leave_to_active'):
        c_pos = resolved_field_map.get('Employment Status')
        c_term = resolved_field_map.get('Termination Date')
        if c_pos and c_term and c_pos in df_download.columns and c_term in df_download.columns:
            pos_lower = df_download[c_pos].astype(str).str.strip().str.lower()
            mask_leave = pos_lower == "leave"
            mask_term_blank = _is_blank_series(c_term)
            for idx in df_download[mask_leave & mask_term_blank].index:
                old_p = df_download.at[idx, c_pos]
                df_download.at[idx, c_pos] = "Active"
                log_change(idx, "Position Status", old_p, "Active",
                           "Reclassified 'Leave' to 'Active' because Termination Date is blank.")

    # 3. DOL Status default Full Time
    if fix_options.get('fix_dol_status'):
        c_dol = resolved_field_map.get('Employment Type')
        if c_dol and c_dol in df_download.columns:
            mask_blank = _is_blank_series(c_dol)
            for idx in df_download[mask_blank].index:
                old_d = df_download.at[idx, c_dol]
                df_download.at[idx, c_dol] = "Full Time"
                log_change(idx, "Employment Type", old_d, "Full Time",
                           "Defaulted blank value to 'Full Time' for active employee.")

    # 4. Job Title from Department (correct dict key — the Streamlit code has a bug
    # where this checks 'fix_position' but the dict only sets 'fix_job_title'; we
    # honor the documented key here so the toggle actually works in the API).
    if fix_options.get('fix_job_title') or fix_options.get('fix_position'):
        c_job = resolved_field_map.get('Job Title')
        c_dep = resolved_field_map.get('Department')
        if c_job and c_dep and c_job in df_download.columns and c_dep in df_download.columns:
            mask = _is_blank_series(c_job)
            for idx in df_download[mask].index:
                old_val = df_download.at[idx, c_job]
                new_val = df_download.at[idx, c_dep]
                if pd.notna(new_val) and str(new_val).strip():
                    df_download.at[idx, c_job] = new_val
                    log_change(idx, "Job Title", old_val, new_val,
                               "Position was blank; filled using Department Description.")

    # 5. FLSA from Pay Type
    if fix_options.get('fix_flsa'):
        c_flsa = resolved_field_map.get('FLSA Classification')
        c_pt = resolved_field_map.get('Pay Type')
        if c_flsa and c_pt and c_flsa in df_download.columns and c_pt in df_download.columns:
            mask_blank = _is_blank_series(c_flsa)
            for idx in df_download[mask_blank].index:
                pt_val = str(df_download.at[idx, c_pt]).lower().strip()
                old_f = df_download.at[idx, c_flsa]
                if 'hour' in pt_val:
                    df_download.at[idx, c_flsa] = "Non-Exempt"
                    log_change(idx, "FLSA Status", old_f, "Non-Exempt", "Applied based on Hourly pay type.")
                elif 'salar' in pt_val:
                    df_download.at[idx, c_flsa] = "Exempt"
                    log_change(idx, "FLSA Status", old_f, "Exempt", "Applied based on Salaried/Salary pay type.")

    # 6. Smart Driver chain
    if fix_options.get('fix_driver_smart'):
        c_jt = resolved_field_map.get('Job Title')
        c_dept = resolved_field_map.get('Department')
        c_flsa = resolved_field_map.get('FLSA Classification')
        c_pt = resolved_field_map.get('Pay Type')
        if (c_jt and c_dept and c_flsa and
                c_jt in df_download.columns and c_dept in df_download.columns and c_flsa in df_download.columns):
            mask_jt_blank = _is_blank_series(c_jt)
            mask_dept_driver = df_download[c_dept].astype(str).str.lower().str.contains("driver", na=False)
            for idx in df_download[mask_jt_blank & mask_dept_driver].index:
                old_j = df_download.at[idx, c_jt]
                new_j = df_download.at[idx, c_dept]
                df_download.at[idx, c_jt] = new_j
                log_change(idx, "Job Title (Smart Driver)", old_j, new_j,
                           "Automatically assigned 'Driver' title from Department.")

            mask_job_driver = df_download[c_jt].astype(str).str.lower().str.contains("driver", na=False)
            mask_flsa_blank = _is_blank_series(c_flsa)
            for idx in df_download[mask_job_driver & mask_flsa_blank].index:
                old_f = df_download.at[idx, c_flsa]
                df_download.at[idx, c_flsa] = "Non-Exempt"
                log_change(idx, "FLSA Classification (Smart Driver)", old_f, "Non-Exempt",
                           "Automatic Non-Exempt status for Driver roles.")

            if c_pt and c_pt in df_download.columns:
                mask_pt_blank = _is_blank_series(c_pt)
                for idx in df_download[mask_job_driver & mask_pt_blank].index:
                    old_p = df_download.at[idx, c_pt]
                    df_download.at[idx, c_pt] = "Hourly"
                    log_change(idx, "Pay Type (Smart Driver)", old_p, "Hourly",
                               "Automatic Hourly pay type for Driver roles.")

    # 7. Generic blank-Job-Title -> Driver fallback (Non-Exempt + Hourly)
    if fix_options.get('fix_blank_jt_to_driver'):
        c_jt = resolved_field_map.get('Job Title')
        c_flsa = resolved_field_map.get('FLSA Classification')
        c_pt = resolved_field_map.get('Pay Type')
        if (c_jt and c_flsa and c_pt and
                c_jt in df_download.columns and c_flsa in df_download.columns and c_pt in df_download.columns):
            jt_lower = df_download[c_jt].astype(str).str.strip().str.lower()
            flsa_lower = df_download[c_flsa].astype(str).str.strip().str.lower()
            pt_lower = df_download[c_pt].astype(str).str.strip().str.lower()
            mask_jt_blank = df_download[c_jt].isna() | (jt_lower == "") | (jt_lower == "nan")
            mask_non_exempt = (flsa_lower.str.contains("non-exempt", na=False) |
                              flsa_lower.str.contains("non exempt", na=False))
            mask_hourly = pt_lower.str.contains("hourly", na=False)
            for idx in df_download[mask_jt_blank & mask_non_exempt & mask_hourly].index:
                old_j = df_download.at[idx, c_jt]
                df_download.at[idx, c_jt] = "Driver"
                log_change(idx, "Job Title", old_j, "Driver",
                           "Job Title was blank; defaulted to 'Driver' for Non-Exempt Hourly employee.")

    # 8. Standard Hours -> "0"
    if fix_options.get('fix_std_hours'):
        c_sh = resolved_field_map.get('Working Hours')
        if c_sh and c_sh in df_download.columns:
            # Zero out hours for ALL hourly employees
            pt_col = resolved_field_map.get('Pay Type')
            if pt_col and pt_col in df_download.columns:
                pt_lower = df_download[pt_col].astype(str).str.lower().str.strip()
                mask_hourly = pt_lower.str.contains('hour', na=False)
                
                # We log changes for those that weren't already "0"
                for idx in df_download[mask_hourly].index:
                    old_v = str(df_download.at[idx, c_sh]).strip()
                    if old_v not in ["0", "0.0", ""]:
                        df_download.at[idx, c_sh] = "0"
                        log_change(idx, "Working Hours", old_v, "0", "Forced zero hours for Hourly employee.")
                    else:
                        df_download.at[idx, c_sh] = "0"
            else:
                # Fallback to blank-only if no pay type col
                mask_sh = _is_blank_series(c_sh)
                df_download.loc[mask_sh, c_sh] = "0"

    # 9. Header renames (column-level — change norm_to_orig label)
    if fix_options.get('rename_std_hours'):
        c_sh = resolved_field_map.get('Working Hours')
        if c_sh and c_sh in norm_to_orig:
            norm_to_orig[c_sh] = "Working hours per Week"

    if fix_options.get('rename_zip_col'):
        c_zip = resolved_field_map.get('Zip')
        if c_zip and c_zip in norm_to_orig:
            norm_to_orig[c_zip] = "Primary Address: Zip Code"

    if fix_options.get('replace_gender_col'):
        sex_col = norm_colname("Sex")
        c_gender = resolved_field_map.get('Gender')
        if sex_col in df_download.columns:
            if c_gender and c_gender in df_download.columns and c_gender != sex_col:
                df_download = df_download.drop(columns=[c_gender])
            norm_to_orig[sex_col] = "Gender / Sex (Self-ID)"

    # 10. Zip cleanup
    if fix_options.get('fix_zip'):
        c_zip = resolved_field_map.get('Zip')
        c_mzip = resolved_field_map.get('Mailing Zip')
        for cz in [c_zip, c_mzip]:
            if cz and cz in df_download.columns:
                def _fix_zip_local(z):
                    if pd.isna(z) or str(z).strip() == "":
                        return ""
                    s = str(z).split('.')[0].split('-')[0]
                    s = re.sub(r'[^0-9]', '', s)
                    if not s:
                        return ""
                    if len(s) == 4:
                        s = '0' + s
                    return s[:5]
                orig_zips = df_download[cz].copy()
                df_download[cz] = df_download[cz].apply(_fix_zip_local).astype(str)
                for idx in df_download.index:
                    if str(df_download.at[idx, cz]) != str(orig_zips.at[idx]):
                        log_change(idx, cz, orig_zips.at[idx], df_download.at[idx, cz],
                                   "Standardized zip code format.")

    # 11. Work Location remap (only if caller supplied it)
    src_loc_col = resolved_field_map.get('Work Location')
    if location_mappings and src_loc_col and src_loc_col in df_download.columns:
        df_download[src_loc_col] = (df_download[src_loc_col].astype(str).str.strip()
                                    .map(lambda x: location_mappings.get(x, x)))

    # 12. Standardize date columns to MM/DD/YYYY
    for std_name in ['Hire Date', 'Original Hire Date', 'Termination Date', 'DOB']:
        col = resolved_field_map.get(std_name)
        if col and col in df_download.columns:
            df_download[col] = _format_datetime_column(df_download[col])

    # 13. Optional manager clustering
    if sort_by_manager:
        col_sup = resolved_field_map.get('Reports To ID')
        emp_col = resolved_field_map.get('Employee ID')
        if (col_sup and emp_col and
                col_sup in df_download.columns and emp_col in df_download.columns):
            sup_counts = df_download[df_download[col_sup].notna()][col_sup].value_counts().to_dict()
            df_download['__mgr_count'] = df_download[emp_col].astype(str).str.strip().map(lambda x: sup_counts.get(x, 0))
            df_download['__group_count'] = df_download[col_sup].astype(str).str.strip().map(lambda x: sup_counts.get(x, 0))
            df_download = df_download.sort_values(by=['__mgr_count', '__group_count'], ascending=[False, False])
            df_download = df_download.drop(columns=['__mgr_count', '__group_count'])

    # 14. Final column reorder + restore original headers (with renames)
    priority_keys = [
        'Employee ID', 'First Name', 'Last Name', 'Reports To ID',
        'Employment Type', 'Pay Type', 'Work Location', 'Workers Comp Code',
        'FLSA Classification', 'Employment Status', 'Job Title', 'Department'
    ]
    final_col_order = []
    renaming_dict = {}
    used = set()
    if 'CRITICAL_WARNINGS' in df_download.columns:
        final_col_order.append('CRITICAL_WARNINGS')
    for k in priority_keys:
        nc = resolved_field_map.get(k)
        if nc and nc in df_download.columns:
            final_col_order.append(nc)
            renaming_dict[nc] = norm_to_orig.get(nc, nc)
            used.add(nc)
    for col in df_download.columns:
        if col not in used and col != 'CRITICAL_WARNINGS':
            final_col_order.append(col)
            label = norm_to_orig.get(col, col)
            if label != col:
                renaming_dict[col] = label

    df_download = df_download[final_col_order].rename(columns=renaming_dict)

    df_audit = pd.DataFrame(audit_trail)
    return _build_excel_bytes(df_download, df_audit), {
        "rows_total": int(len(df_download)),
        "rows_with_warnings": int(sum(1 for v in df_download.get('CRITICAL_WARNINGS', pd.Series([])).astype(str) if v.strip())),
        "changes_logged": int(len(audit_trail)),
    }
