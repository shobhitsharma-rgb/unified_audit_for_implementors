import streamlit as st

def inject_premium_styles():
    """Injects global 'Editorial Ledger' CSS and Typography (Manrope/Inter)."""
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Manrope:wght@200;400;700;800&family=Inter:wght@300;400;600&display=swap');

        /* Typography: Apply to content but NOT to system icons/expander-arrows */
        .stMarkdown p, .stMarkdown li, .stMarkdown label, .stTable, .stDataFrame {
            font-family: 'Inter', sans-serif !important;
            color: #1b1c1c !important;
        }

        /* Headers with Manrope */
        h1, h2, h3, h4, h5, h6 {
            font-family: 'Manrope', sans-serif !important;
            font-weight: 700 !important;
            color: #050e39 !important;
            letter-spacing: -0.02em !important;
        }

        /* The Editorial Card Surface */
        .premium-card {
            background-color: #ffffff;
            border-radius: 12px;
            padding: 24px;
            margin-bottom: 20px;
            box-shadow: 0 4px 24px rgba(0, 0, 0, 0.04);
            border: 1px solid rgba(198, 197, 208, 0.2);
        }

        /* Action Hub Theme */
        .action-hub-error { background: linear-gradient(145deg, #fff5f5, #ffffff); border-left: 5px solid #ba1a1a; border-radius: 8px; padding: 20px; margin-bottom: 16px; }
        .action-hub-warning { background: linear-gradient(145deg, #fffaf3, #ffffff); border-left: 5px solid #e8881f; border-radius: 8px; padding: 20px; margin-bottom: 16px; }

        /* Glossy Primary Button - High Contrast Fix */
        div.stButton > button {
            background: linear-gradient(135deg, #050e39 0%, #1c244e 100%) !important;
            border: none !important;
            border-radius: 8px !important;
            padding: 12px 32px !important;
            transition: all 0.3s ease !important;
            box-shadow: 0 4px 12px rgba(5, 14, 57, 0.2) !important;
        }
        
        /* Force White Text on ALL button children (p, span, labels) */
        div.stButton > button * {
            color: #ffffff !important;
            font-weight: 700 !important;
            font-family: 'Manrope', sans-serif !important;
            text-shadow: 0 1px 2px rgba(0,0,0,0.2) !important;
        }

        /* Refined Expander (FIX: Don't break the SVG icons) */
        .stExpander {
            border: none !important;
            background-color: #f8f9fa !important;
            border-radius: 12px !important;
            margin-bottom: 16px !important;
            border: 1px solid rgba(0,0,0,0.05) !important;
        }
        
        /* Targeted Title text - Leave the summary arrow alone */
        div[data-testid="stExpanderSummary"] > div:last-child p {
            font-family: 'Manrope', sans-serif !important;
            font-weight: 700 !important;
            font-size: 1.05rem !important;
            color: #050e39 !important;
            margin: 0 !important;
        }

        /* Pill Badges */
        .pill-error {
            background-color: #ffdad6;
            color: #ba1a1a;
            padding: 2px 10px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }
        .pill-warning {
            background-color: #ffdcc1;
            color: #6c3a00;
            padding: 2px 10px;
            border-radius: 20px;
            font-size: 0.85rem;
            font-weight: 600;
        }
        </style>
    """, unsafe_allow_html=True)

def render_premium_header(title, subtitle=None):
    """Renders a styled header for content sections."""
    st.markdown(f"### {title}")
    if subtitle:
        st.markdown(f"<p style='color: #46464f; margin-top: -10px; margin-bottom: 20px;'>{subtitle}</p>", unsafe_allow_html=True)

def action_hub_container(type='error'):
    """Context manager for rendering audit findings in styled containers."""
    css_class = 'action-hub-error' if type == 'error' else 'action-hub-warning'
    return st.container() # In Streamlit, styling full containers requires CSS injection based on nesting or IDs, but for this MVP we'll use markdown blocks inside.

def render_finding_card(title, data_dict, type='error'):
    """Renders a high-fidelity card for audit findings (minified for Streamlit compatibility)."""
    bg_color = "#fff5f5" if type == 'error' else "#fffaf3"
    border_color = "#ba1a1a" if type == 'error' else "#e8881f"
    text_color = "#93000a" if type == 'error' else "#673700"
    
    # We build a single-line minified HTML string to avoid Streamlit markdown parsing issues
    html = f'<div style="background: {bg_color}; border-left: 5px solid {border_color}; border-radius: 8px; padding: 16px; margin-bottom: 16px;">'
    html += f'<div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 8px;">'
    html += f'<h5 style="color: {text_color}; border: none; padding: 0; margin: 0;">{title}</h5>'
    html += '</div>'
    html += '<div style="display: grid; grid-template-columns: repeat(2, 1fr); gap: 12px; margin-top: 12px;">'
    
    for label, value in data_dict.items():
        html += '<div>'
        html += f'<p style="font-size: 0.75rem; color: #46464f; margin: 0;">{label}</p>'
        html += f'<p style="font-size: 1rem; font-weight: 600; color: {text_color}; margin: 0;">{value}</p>'
        html += '</div>'

    html += '</div></div>'
    st.markdown(html, unsafe_allow_html=True)


def render_duplicate_column_error(dupes):
    """Hard-stop error shown when an uploaded census file has repeated column
    headers. Tells the user — in plain English — exactly which columns repeat,
    what to do, and that nothing has been generated."""
    st.markdown("""
<div style="background:#fff5f5; border-left:5px solid #ba1a1a; border-radius:8px; padding:16px 20px; margin:8px 0 4px 0;">
<h4 style="color:#93000a; margin:0 0 4px 0; border:none; padding:0;">⛔ This file can't be processed — it has repeated column names</h4>
<p style="color:#46464f; margin:0; font-size:0.9rem;">The sanity check has been stopped. No census file has been generated.</p>
</div>
""", unsafe_allow_html=True)
    dupe_list = "\n".join(f"- **{d}**" for d in dupes)
    st.markdown(
        f"The column name{'s' if len(dupes) != 1 else ''} below appear "
        f"**more than once** in your file:\n\n{dupe_list}\n\n"
        "A census file must have **every column name only once**. Please:\n\n"
        "1. Open the file in Excel.\n"
        "2. For each repeated name above, decide which copy of the column to keep.\n"
        "3. **Delete the extra column(s)** so each name is left only once.\n"
        "4. Save the file and upload it again here.\n\n"
        "The sanity check will start only after the duplicate columns are removed."
    )


# Census fields that MUST be present for the sanity check to run — keyed by the
# standard field name used in ADP_FIELD_MAP / PAYCOM_FIELD_MAP. If any is absent
# from an upload, the tool hard-stops (otherwise its checks are skipped silently).
REQUIRED_CENSUS_FIELDS = [
    'Employee ID', 'First Name', 'Last Name', 'SSN', 'DOB',
    'Employment Status', 'Employment Type', 'Hire Date',
    'Pay Type', 'FLSA Classification', 'Annual Salary', 'Working Hours',
    'Job Title', 'Address Line 1', 'City', 'Zip', 'State',
]

_FIELD_FRIENDLY = {
    'Employee ID': 'employee ID', 'First Name': 'first name',
    'Last Name': 'last name', 'SSN': 'Social Security Number',
    'DOB': 'date of birth', 'Employment Status': 'employment status (Active / Terminated)',
    'Employment Type': 'employment type (Full-Time / Part-Time)',
    'Hire Date': 'hire date', 'Pay Type': 'pay type (Hourly / Salary)',
    'FLSA Classification': 'Exempt / Non-Exempt status',
    'Annual Salary': 'annual salary', 'Working Hours': 'scheduled work hours',
    'Job Title': 'job title', 'Address Line 1': 'street address',
    'City': 'city', 'Zip': 'home zip code', 'State': 'state',
}


def render_missing_column_error(missing):
    """Hard-stop error shown when an uploaded census file is missing required
    columns. `missing` is a list of (expected_header, standard_field_name)."""
    st.markdown("""
<div style="background:#fff5f5; border-left:5px solid #ba1a1a; border-radius:8px; padding:16px 20px; margin:8px 0 4px 0;">
<h4 style="color:#93000a; margin:0 0 4px 0; border:none; padding:0;">⛔ This file can't be processed — required columns are missing</h4>
<p style="color:#46464f; margin:0; font-size:0.9rem;">The sanity check has been stopped. No census file has been generated.</p>
</div>
""", unsafe_allow_html=True)
    rows = "\n".join(
        f"- **{hdr}** — the {_FIELD_FRIENDLY.get(std, std)} column"
        for hdr, std in missing
    )
    many = len(missing) != 1
    st.markdown(
        f"The Census Sanity Check needs the column{'s' if many else ''} below, "
        f"but {'they were' if many else 'it was'} not found in your file:\n\n"
        f"{rows}\n\n"
        "Without these, important checks would be skipped without telling you. Please:\n\n"
        "1. Open your census export.\n"
        "2. Add the missing column(s), using the exact name(s) shown above.\n"
        "3. Fill in the value for every employee.\n"
        "4. Save the file and upload it again here.\n\n"
        "The sanity check will start only after every required column is present."
    )


def render_sanity_disclaimer():
    """Standing disclaimer shown at the top of both Census Sanity Check tools."""
    st.info(
        "**Note:** This tool was built and groomed over the course of a 4-month "
        "implementation. It handles the vast majority of cases automatically — but "
        "there may still be scenarios where human intervention and judgement are "
        "necessary. Always review the results and the Change Log before proceeding."
    )


def render_standardization_notice(include_column_renames=False):
    """Notice listing the file-wide standardizations applied to every corrected
    census download. These are also recorded as summary rows in the Change Log."""
    items = [
        "**Working hours** — set to **0** for every employee.",
        "**Dates** — hire, termination and birth dates reformatted to **MM/DD/YYYY**.",
        "**Column order** — key fields (Employee ID, Name, Pay Type, FLSA, etc.) moved to the front.",
        "**Row order** — when a reporting hierarchy is present, employee rows are grouped so each manager sits with their reportees.",
    ]
    if include_column_renames:
        items.append(
            "**Column headers** — the home-zip column header is standardized to "
            '"Primary Address: Zip Code", and the **Gender** column is populated '
            "from the **Sex** column."
        )
    bullets = "\n".join(f"- {it}" for it in items)
    st.info(
        "ℹ️ **Every corrected file is also standardized — no action needed.** "
        "These changes are applied automatically to all employees and are recorded "
        "in the Change Log of your download:\n\n"
        + bullets
    )


def _plain_english_issue(raw_issue):
    """Translate a raw validation issue string into plain English a non-payroll
    user can understand. Messages are GENERIC (no per-employee values such as a
    job title or state name) so every employee with the same kind of problem
    groups under a single line — the per-employee detail lives in the
    'View the full list of affected employees' table. Falls back to the raw
    text if no rule matches."""
    p = str(raw_issue).strip()
    if not p:
        return ""
    if "Duplicate SSN" in p:
        return "Same Social Security Number is used by more than one employee"
    if "SSN (blank)" in p:
        return "Missing Social Security Number"
    if "Employment Status (blank)" in p:
        return "Missing employment status (should be Active or Terminated)"
    if "Non-standard Status" in p:
        return "Unrecognized employment status (should be Active or Terminated)"
    if "Terminated but missing Termination Date" in p:
        return "Marked as Terminated but has no termination date"
    if "Employment Type (blank)" in p:
        return "Missing employment type (should be Full-Time or Part-Time)"
    if "Pay Type (blank)" in p:
        return "Missing pay type (should be Hourly or Salary)"
    if "Job Title (blank)" in p:
        return "Missing job title"
    if "Work Location (blank)" in p:
        return "Missing work location"
    if "Zip Code (blank)" in p:
        return "Missing home zip code"
    if "Zip Code" in p:
        return "Home zip code is not valid (needs to be 5 digits)"
    if "Annual Salary" in p:
        return "Salaried employee is missing an annual salary amount"
    if "Salaried Hourly-Only Exception" in p:
        return "Driver / Walker / Helper role is marked as salary — must be hourly pay"
    if "State" in p and "full name" in p:
        return "State should be the 2-letter code, e.g. NY instead of New York"
    if "predates date of hire" in p:
        return "Termination date is earlier than the hire date"
    if "Special characters in" in p:
        return "An emergency contact field contains unsupported special characters"
    return p


def render_validation_results(hard_errors, flsa_corrections, flsa_blanks,
                              anomalies, intern_corrections, email_fallbacks,
                              smart_driver_fixes=None, position_blanks=None,
                              dol_status_blanks=None, zip_fixes=None):
    """Render census validation results in a plain-English, two-section layout:
      1. 'Needs your attention'  — problems the user should review (red)
      2. 'Fixed automatically'   — corrections applied on download (green)
    Shows a success banner when nothing needs attention."""
    import pandas as pd
    from collections import defaultdict

    if smart_driver_fixes is None:
        smart_driver_fixes = pd.DataFrame()
    if position_blanks is None:
        position_blanks = pd.DataFrame()
    if dol_status_blanks is None:
        dol_status_blanks = pd.DataFrame()
    if zip_fixes is None:
        zip_fixes = pd.DataFrame()

    def _ids_str(df_in):
        if df_in is None or df_in.empty or 'Employee ID' not in df_in.columns:
            return ""
        ids = [str(x) for x in df_in['Employee ID'].unique().tolist()]
        shown = ", ".join(ids[:5])
        if len(ids) > 5:
            shown += f" (+{len(ids) - 5} more)"
        return shown

    def _n(df_in):
        # Count distinct employees, not rows — the validation engine can append
        # the same employee more than once to a finding list.
        if df_in is None or df_in.empty:
            return 0
        if 'Employee ID' in df_in.columns:
            return df_in['Employee ID'].nunique()
        return len(df_in)

    has_hard = hard_errors is not None and not hard_errors.empty
    auto_frames = [flsa_corrections, flsa_blanks, anomalies, intern_corrections,
                   email_fallbacks, smart_driver_fixes, position_blanks,
                   dol_status_blanks, zip_fixes]
    has_auto = any(f is not None and not f.empty for f in auto_frames)

    if not has_hard and not has_auto:
        st.success("✅ Your file looks good — no issues found. Click **Download Corrected Source** below to get your cleaned file.")
        return

    # --- SECTION 1: NEEDS YOUR ATTENTION ---
    if has_hard:
        issue_to_ids = defaultdict(list)
        for _, err in hard_errors.iterrows():
            eid = str(err['Employee ID'])
            for part in [p.strip() for p in str(err['Issue']).split(",") if p.strip()]:
                plain = _plain_english_issue(part)
                if plain and eid not in issue_to_ids[plain]:
                    issue_to_ids[plain].append(eid)

        n_emp = len(hard_errors['Employee ID'].unique())
        n_issues = len(issue_to_ids)
        st.markdown(f"""
<div style="background:#fff5f5; border-left:5px solid #ba1a1a; border-radius:8px; padding:16px 20px; margin:8px 0 4px 0;">
<h4 style="color:#93000a; margin:0 0 4px 0; border:none; padding:0;">⚠️ {n_issues} type{'s' if n_issues != 1 else ''} of issue need your attention</h4>
<p style="color:#46464f; margin:0; font-size:0.9rem;">Found across {n_emp} employee{'s' if n_emp != 1 else ''}. Please review these before uploading to Uzio — each one is also listed in the Change Log of your download.</p>
</div>
""", unsafe_allow_html=True)

        with st.container(height=350, border=True):
            for plain_issue, ids in sorted(issue_to_ids.items(), key=lambda kv: -len(kv[1])):
                shown = ", ".join(ids[:5])
                if len(ids) > 5:
                    shown += f" (+{len(ids) - 5} more)"
                st.markdown(f"**{plain_issue}** — {len(ids)} employee{'s' if len(ids) != 1 else ''}  \n&nbsp;&nbsp;&nbsp;Employee IDs: `{shown}`")

        with st.expander("🔍 View the full list of affected employees", expanded=False):
            st.dataframe(hard_errors, hide_index=True, use_container_width=True)

    # --- SECTION 2: FIXED AUTOMATICALLY (no action needed) ---
    # Every line here must describe what the download ACTUALLY does — no
    # "filled from department" claims for a vendor (ADP) that has no department
    # column. position_blanks / smart_driver_fixes carry a discriminator
    # (Resolution / Issue) so each real fix gets its own accurate sentence.
    def _subset(df_in, col, predicate):
        if df_in is None or df_in.empty or col not in df_in.columns:
            return pd.DataFrame()
        return df_in[df_in[col].astype(str).apply(predicate)]

    fixes = []
    if flsa_blanks is not None and not flsa_blanks.empty:
        n = _n(flsa_blanks)
        fixes.append(f"**Exempt / Non-Exempt status was blank** — we set it from each employee's pay type (paid hourly → Non-Exempt, salaried → Exempt). {n} employee{'s' if n != 1 else ''}: `{_ids_str(flsa_blanks)}`")

    # Smart-driver fixes split by what was actually blank.
    sd_flsa = _subset(smart_driver_fixes, 'Issue', lambda s: s.startswith('Blank FLSA'))
    sd_dept = _subset(smart_driver_fixes, 'Issue', lambda s: s.startswith('Blank Position'))
    if smart_driver_fixes is not None and not smart_driver_fixes.empty and 'Issue' not in smart_driver_fixes.columns:
        # Defensive fallback if an older caller passes frames without the discriminator.
        sd_flsa = smart_driver_fixes
    if sd_flsa is not None and not sd_flsa.empty:
        n = _n(sd_flsa)
        fixes.append(f"**A Driver / Walker had a blank overtime status or pay type** — we set it to Non-Exempt and Hourly based on the Driver job title. {n} employee{'s' if n != 1 else ''}: `{_ids_str(sd_flsa)}`")
    if sd_dept is not None and not sd_dept.empty:
        n = _n(sd_dept)
        fixes.append(f"**Job title was blank and the department is a Driver role** — we filled in the Driver title and set Non-Exempt + Hourly. {n} employee{'s' if n != 1 else ''}: `{_ids_str(sd_dept)}`")

    if intern_corrections is not None and not intern_corrections.empty:
        n = _n(intern_corrections)
        fixes.append(f'**Employment type "Intern" was changed to "Part-Time"**. {n} employee' + ('s' if n != 1 else '') + f": `{_ids_str(intern_corrections)}`")
    if email_fallbacks is not None and not email_fallbacks.empty:
        n = _n(email_fallbacks)
        fixes.append(f"**Work email was empty** — we used the employee's personal email instead. {n} employee{'s' if n != 1 else ''}: `{_ids_str(email_fallbacks)}`")

    # Blank job title split by how it was actually resolved.
    pb_dept = _subset(position_blanks, 'Resolution', lambda s: s == 'department')
    pb_driver = _subset(position_blanks, 'Resolution', lambda s: s == 'driver-default')
    if position_blanks is not None and not position_blanks.empty and 'Resolution' not in position_blanks.columns:
        pb_dept = position_blanks  # older caller without the discriminator
    if pb_dept is not None and not pb_dept.empty:
        n = _n(pb_dept)
        fixes.append(f"**Job title was blank** — we filled it in from the department name. {n} employee{'s' if n != 1 else ''}: `{_ids_str(pb_dept)}`")
    if pb_driver is not None and not pb_driver.empty:
        n = _n(pb_driver)
        fixes.append(f"**Job title was blank** — we set it to \"Driver\" (the employee is Non-Exempt and paid Hourly). {n} employee{'s' if n != 1 else ''}: `{_ids_str(pb_driver)}`")

    if dol_status_blanks is not None and not dol_status_blanks.empty:
        n = _n(dol_status_blanks)
        fixes.append(f'**Employment type was blank** — we set it to "Full Time". {n} employee' + ('s' if n != 1 else '') + f": `{_ids_str(dol_status_blanks)}`")
    if zip_fixes is not None and not zip_fixes.empty:
        n = _n(zip_fixes)
        fixes.append(f"**Home zip code wasn't 5 digits** — we padded a leading zero or trimmed extra digits to make it 5. {n} employee{'s' if n != 1 else ''}: `{_ids_str(zip_fixes)}`")

    if fixes:
        st.markdown("""
<div style="background:#f0faf4; border-left:5px solid #1a7a4a; border-radius:8px; padding:16px 20px; margin:16px 0 4px 0;">
<h4 style="color:#1a4a2a; margin:0 0 4px 0; border:none; padding:0;">✅ Fixed automatically — no action needed</h4>
<p style="color:#46464f; margin:0; font-size:0.9rem;">These corrections are applied to your file when you download it, and recorded in the Change Log.</p>
</div>
""", unsafe_allow_html=True)
        with st.container(height=260, border=True):
            for fix in fixes:
                st.markdown(f"- {fix}")

    # --- SECTION 3: PLEASE REVIEW (spotted but NOT changed) ---
    reviews = []
    if flsa_corrections is not None and not flsa_corrections.empty:
        n = _n(flsa_corrections)
        reviews.append(f"**Pay type and overtime setting don't match** — paid hourly but marked Exempt, or salaried but marked Non-Exempt. We left the data unchanged — please confirm which is correct. {n} employee{'s' if n != 1 else ''}: `{_ids_str(flsa_corrections)}`")
    # Status anomalies (On Leave / Inactive) — exclude any employee already
    # covered by the pay-type-mismatch line above so no one is listed twice.
    status_review = anomalies
    if (anomalies is not None and not anomalies.empty
            and flsa_corrections is not None and not flsa_corrections.empty
            and 'Employee ID' in anomalies.columns):
        fc_ids = set(flsa_corrections['Employee ID'])
        status_review = anomalies[~anomalies['Employee ID'].isin(fc_ids)]
    if status_review is not None and not status_review.empty:
        n = _n(status_review)
        reviews.append(f"**Employee status needs a second look** — e.g. marked On Leave or Inactive with no end date. We left the data unchanged — please review each one. {n} employee{'s' if n != 1 else ''}: `{_ids_str(status_review)}`")

    if reviews:
        st.markdown("""
<div style="background:#fffaf3; border-left:5px solid #e8881f; border-radius:8px; padding:16px 20px; margin:16px 0 4px 0;">
<h4 style="color:#6c3a00; margin:0 0 4px 0; border:none; padding:0;">👀 Please review before uploading</h4>
<p style="color:#46464f; margin:0; font-size:0.9rem;">The tool spotted these but did not change them — they need a person to decide. They're noted in the Change Log too.</p>
</div>
""", unsafe_allow_html=True)
        with st.container(height=180, border=True):
            for r in reviews:
                st.markdown(f"- {r}")
