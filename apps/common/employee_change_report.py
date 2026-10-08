"""Employee Profile Change Report — the same workbook Uzio exports, per employee.

Signs in with the user's OWN Uzio credentials and reads the employee history
through the read-only NeuronOps query endpoint, so every read is audited against
whoever is signed in. Nothing is written anywhere: no credential, no token and no
generated file is stored on the server or on disk — the token lives in this
browser session only and the workbooks are built in memory for download.

The workbook itself is produced by utils/change_report.py, which is a port of
Uzio's own report generator. Its docstring lists the two things that port cannot
reproduce (encrypted pay/SSN values, and the Work Schedule / Family lookups).
"""
import io
import zipfile
from datetime import datetime

import pandas as pd
import streamlit as st

from utils import change_report as cr
from utils import neuronops_client as ops
from utils.ui_components import _callout, render_premium_header

TOKEN_KEY = "ecr_token"
USER_KEY = "ecr_user"
HOST_KEY = "ecr_host"
MATCHES_KEY = "ecr_matches"


def _sign_in():
    render_premium_header("Sign in with your Uzio credentials",
                          "The same login you use for the onboarding API. Your password is "
                          "used once to get a token and is never stored.")
    with st.form("ecr_login"):
        col1, col2 = st.columns(2)
        username = col1.text_input("Username", key="ecr_username",
                                   placeholder="firstname.lastname@uzio.com")
        password = col2.text_input("Password", type="password", key="ecr_password")
        host = st.text_input("Environment", value=ops.DEFAULT_HOST,
                             help="Leave as-is for production.")
        submitted = st.form_submit_button("Sign in", type="primary")

    if submitted:
        if not username or not password:
            st.error("Enter both your username and password.")
            return
        with st.spinner("Signing in…"):
            try:
                st.session_state[TOKEN_KEY] = ops.login(username.strip(), password, host=host.strip())
            except ops.NeuronOpsError as e:
                st.error(str(e))
                return
        st.session_state[USER_KEY] = username.strip()
        st.session_state[HOST_KEY] = host.strip()
        st.rerun()


def _find_employees(token, host):
    render_premium_header("2. Find the employees",
                          "Start with the company's FEIN — the same Employee ID exists in "
                          "other companies, so without it you get strangers back. Leave the "
                          "Employee IDs box empty to list everyone in that company.")
    with st.form("ecr_search"):
        col1, col2 = st.columns([1, 2])
        fein = col1.text_input("Company FEIN", placeholder="863131339",
                               help="The 9-digit federal EIN. More than one is allowed, "
                                    "comma separated.")
        ids = col2.text_area("Employee IDs (optional)", height=110,
                             placeholder="1020, BH0KS5HPZ' + BS + 'n8OSU7337G",
                             help="Comma or newline separated. An ID Uzio has since "
                                  "replaced still finds the employee.")
        searched = st.form_submit_button("Search", type="primary")

    if not searched:
        return
    fein_list = [p.strip() for p in fein.replace("' + BS + 'n", ",").split(",") if p.strip()]
    id_list = [p.strip() for p in ids.replace("' + BS + 'n", ",").split(",") if p.strip()]
    if not fein_list and not id_list:
        st.warning("Enter the company FEIN, or at least one Employee ID.")
        return
    if not fein_list:
        st.info("No FEIN given, so this searches across every company. Check the Company "
                "column before you generate.")

    with st.spinner("Searching…"):
        try:
            if fein_list:
                companies = cr.find_companies(token, fein_list, host=host)
                found = {str(c.get("fein")) for c in companies}
                unknown = [f for f in fein_list if f not in found]
                if unknown:
                    st.error("No company in Uzio has FEIN " + ", ".join(unknown) + ".")
                if not companies:
                    st.session_state[MATCHES_KEY] = []
                    return
                st.caption("Company: " + " · ".join(
                    f"**{c.get('company_name')}** ({c.get('fein')})" for c in companies))
            st.session_state[MATCHES_KEY] = cr.find_employees(
                token, feins=fein_list, ids=id_list, host=host)
        except ops.NeuronOpsError as e:
            st.error(str(e))
            return

    matched = {str(m.get("ext_employee_code")) for m in st.session_state[MATCHES_KEY]}
    matched |= {str(m.get("found_via", "")).replace("was ", "")
                for m in st.session_state[MATCHES_KEY]}
    missing = [i for i in id_list if i not in matched]
    if missing:
        st.warning("No employee found for: " + ", ".join(missing))


def _selection_table(matches):
    frame = pd.DataFrame([{
        "Generate": False,
        "Employee Name": m.get("full_name") or "",
        "Employee ID": m.get("ext_employee_code") or "",
        "Matched on": m.get("found_via") or "",
        "Status": m.get("status_label") or "",
        "Date of Hire": cr.cell_value(m, "date", "date_of_hire"),
        "Termination Date": cr.cell_value(m, "date", "date_of_termination"),
        "Company": m.get("company_name") or "",
        "FEIN": m.get("fein") or "",
        "_code": m.get("employee_code"),
    } for m in matches])

    edited = st.data_editor(
        frame, hide_index=True, use_container_width=True, key="ecr_picker",
        column_config={
            "Generate": st.column_config.CheckboxColumn(required=True),
            "_code": None,
            "Matched on": st.column_config.TextColumn(
                help="Filled in when the Employee ID you searched for is an older one that "
                     "Uzio has since replaced."),
            "Company": st.column_config.TextColumn(
                help="Which company this employee belongs to — the way to tell apart two "
                     "employees who share an Employee ID."),
        },
        disabled=[c for c in frame.columns if c != "Generate"])
    return edited[edited["Generate"]]["_code"].tolist()


def _versions_table(versions):
    return pd.DataFrame([{
        "Version": cr.VERSION_STR + str(i + 1),
        "Who Changed": cr.who_changed(v),
        "Modified On": cr.modified_on(v.get("created_date")),
        "Source of Change": cr.SOURCE.get(str(v.get("source") or ""), str(v.get("source") or "")),
        "IP Address": v.get("ip_address") or "",
        "Employment Status": cr.cell_value(v, "status", "status"),
    } for i, v in enumerate(versions)])


def _generate(token, host, codes, matches):
    by_code = {m.get("employee_code"): m for m in matches}
    stamp = datetime.now()
    built = []
    skipped = []

    progress = st.progress(0.0, text="Reading change history…")
    for n, code in enumerate(codes, start=1):
        employee = by_code.get(code, {})
        label = employee.get("full_name") or employee.get("ext_employee_code") or code
        progress.progress(n / len(codes), text=f"Reading change history — {label}")
        try:
            workbook, versions = cr.build_for_employee(token, code, host=host)
        except ops.NeuronOpsError as e:
            st.error(f"{label}: {e}")
            return
        if workbook is None:
            skipped.append(label)
            continue
        stream = io.BytesIO()
        workbook.save(stream)
        name = cr.report_filename(stamp, employee.get("ext_employee_code") or label)
        built.append((label, name, stream.getvalue(), versions))
    progress.empty()

    if skipped:
        st.warning("No change history exists for: " + ", ".join(skipped))
    if not built:
        return

    st.markdown(_callout("ok", f"{len(built)} report(s) ready",
                         "Each file is the Change History sheet for one employee, newest "
                         "version in the first column."), unsafe_allow_html=True)

    if len(built) == 1:
        label, name, data, versions = built[0]
        st.download_button(f"Download — {label}", data=data, file_name=name, type="primary",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        bundle = io.BytesIO()
        with zipfile.ZipFile(bundle, "w", zipfile.ZIP_DEFLATED) as zf:
            for _, name, data, _ in built:
                zf.writestr(name, data)
        st.download_button(f"Download all {len(built)} reports (.zip)", data=bundle.getvalue(),
                           file_name=f"Employee Profile Change Reports_"
                                     f"{stamp.strftime('%Y-%m-%d-%H-%M-%S')}.zip",
                           mime="application/zip", type="primary")
        for label, name, data, _ in built:
            st.download_button(label, data=data, file_name=name, key=f"ecr_dl_{name}",
                               mime="application/vnd.openxmlformats-officedocument."
                                    "spreadsheetml.sheet")

    for label, _, _, versions in built:
        with st.expander(f"{label} — {len(versions)} version(s)"):
            st.dataframe(_versions_table(versions), hide_index=True, use_container_width=True)


def render_ui():
    st.title("Employee Profile Change Report")
    st.markdown(
        "Produces the **same workbook Uzio exports** — one file per employee, one column "
        "per version, newest first — so you can see who changed a profile, when, from "
        "which IP and through which channel (UI, census template, enrollment or payroll)."
    )

    token = st.session_state.get(TOKEN_KEY)
    if not token:
        st.markdown(_callout("warn", "Read-only, and audited under your own login",
                             "This tool runs a single SELECT against the reporting endpoint. It "
                             "cannot change anything, and the read is recorded against the "
                             "credentials you sign in with. Your password and token are kept in "
                             "this browser session only — nothing is saved on the server."),
                    unsafe_allow_html=True)
        _sign_in()
        return

    host = st.session_state.get(HOST_KEY) or ops.DEFAULT_HOST
    col1, col2 = st.columns([4, 1])
    col1.caption(f"Signed in as **{st.session_state.get(USER_KEY, '')}** · {host}")
    if col2.button("Sign out"):
        for key in (TOKEN_KEY, USER_KEY, HOST_KEY, MATCHES_KEY):
            st.session_state.pop(key, None)
        st.rerun()

    with st.expander("How to read the report"):
        st.markdown(
            "**The sheet is called `Change History`.** Column A is the section, column B the "
            "field, and then **one column per version of the profile — newest first**. The "
            "five lines above each version column say which version it is, **who changed it**, "
            "**when** (UTC), from **which IP**, and **through which channel** — User Interface, "
            "Census Template, Enrollment Update or Payroll Integration.\n\n"
            "**Read a row left to right to follow one field through time.** V1 is the oldest "
            "version — usually the migration itself — and the leftmost column is what the "
            "profile looks like now.\n\n"
            "**Rows that are highlighted green changed at some point.** A row whose value is "
            "the same in every version is left plain, so the highlighted rows are the ones "
            "worth reading for a delta or an RCA.\n\n"
            f"**`{cr.ENCRYPTED}` is a value this tool cannot read** — SSN and the pay "
            "fields are encrypted in the database. For a pay field you can still see *when* "
            "it changed without the number: the `(Effective Date - …)` beside it moves to the "
            "date the new amount applies from, in the version where someone changed it. "
            "Open that employee in Uzio if you need the figure itself.\n\n"
            "**`(Effective Date - …)` is not the same as the change date.** The change date is "
            "in the column header — when someone edited the profile. The effective date inside "
            "the cell is the date the value applies from, which the person editing chose and "
            "which can be in the past or the future.")

    with st.expander("What this report cannot show"):
        st.markdown(
            "- **Employee SSN, Hourly Pay Rate, Annual Salary, Bonus, Salary Commissions** — "
            "these are encrypted at rest. Uzio decrypts them inside the application; the "
            "reporting endpoint returns the encrypted value, so the cell reads "
            f"`{cr.ENCRYPTED}`. The effective date next to a pay field is real and is "
            "shown — and that date is how you spot a pay change: it moves to the date the "
            "new rate applies from, in the version where the rate was changed.\n"
            "- **Work Schedule** and the **Family (dependents)** section need lookups the "
            "reporting endpoint does not expose, so Work Schedule stays blank and the Family "
            "section is not written.\n"
            "- **Original DOH** is the one cell that is deliberately *not* the same as "
            "Uzio's export. Uzio's own report leaves it blank for everyone — its "
            "`getOriginalDOHString()` checks the formatted string instead of the date it "
            "is about to format, so the value never gets written. This tool shows the "
            "date the history row actually holds.\n"
            "- **Union Classification** is left out, exactly as Uzio leaves it out for an "
            "Amazon exchange — unless a version actually carries a value, in which case the "
            "row is written so nothing is hidden.")

    _find_employees(token, host)

    matches = st.session_state.get(MATCHES_KEY)
    if matches is None:
        return
    if not matches:
        st.info("No employees matched. Check the Employee ID, or search by name instead.")
        return

    render_premium_header("3. Pick who to generate for",
                          f"{len(matches)} employee(s) matched.")
    codes = _selection_table(matches)
    if not codes:
        st.info("Tick at least one employee above.")
        return
    if st.button(f"Generate {len(codes)} report(s)", type="primary"):
        _generate(token, host, codes, matches)
