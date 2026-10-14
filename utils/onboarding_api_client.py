"""Streamlit UI for the Uzio onboarding push (Census Sanity Check, root repo only).

Everything that talks to Uzio lives in utils/onboarding_core.py so this module and
the census-migration skill's scripts run identical code -- only the drawing is here.
The names re-exported below are kept so existing imports of this module keep working.

Credentials are never persisted: the password widget keeps the typed value in the
browser session only, and session_state holds the push response and the run's
error rows so they survive reruns -- never the password or the JWT.
"""
import pandas as pd
import streamlit as st

from utils.onboarding_core import (  # noqa: F401  (re-exported for existing callers)
    push_payment_method,
    DEFAULT_PROD_HOST, OnboardingAPIError, attach_run_log, automation_id,
    fetch_run_log, find_run_since, group_issues, ist, lost_answer, login,
    needs_recheck, new_result, parse_ts, push_employee_census, run_issues,
)
from utils.ui_components import _callout

from datetime import datetime, timezone


def _render_issue_groups(issues, label):
    """One line per distinct reason, with its count and the affected Employee IDs."""
    rows = group_issues(issues)
    st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)
    if len(issues) > len(rows):
        with st.expander(f"View all {len(issues)} {label} rows"):
            st.dataframe(pd.DataFrame(issues), hide_index=True, use_container_width=True)


def render_push_result(result: dict, key_prefix: str):
    """The push outcome plus what the API logged for that run (errors / warnings)."""
    lost = lost_answer(result)
    if result.get("push_error"):
        st.warning(f"⏳ Uzio didn't answer before the connection closed ({result['push_error']}).\n\n"
                   "A large census keeps processing on Uzio's side after that, so the tool looks "
                   "for the run this push started instead.")
    elif result["ok"]:
        st.success(f"✅ Census pushed successfully (HTTP {result['status']}).")
    elif lost:
        st.warning(f"⏳ The gateway stopped waiting (HTTP {result['status']}). A large census keeps "
                   "processing on Uzio's side, so the tool looks for the run this push started instead.")
    else:
        st.error(f"❌ API returned HTTP {result['status']}")
    if result.get("text"):
        with st.expander("Response details", expanded=not (result["ok"] or lost)):
            st.code(result["text"])

    run_id = result.get("run_id")
    if run_id is None:
        if result.get("run_error"):
            st.warning(f"Couldn't look up the run: {result['run_error']}")
        elif lost:
            st.error(f"No run was found for FEIN {result['fein']} started by {result['user']} since "
                     f"{ist(result['pushed_at'])} IST, so this census most likely never reached Uzio. "
                     "Check the onboarding logs before pushing again.")
        return

    st.markdown(f"#### 📋 What the API logged for run {run_id}")
    if result.get("found_by_lookup"):
        st.caption("Found by looking up your first run for this FEIN since the push started — "
                   "the push itself got no answer.")
    if result.get("run_error"):
        st.warning(f"Couldn't read the log for this run: {result['run_error']}\n\n"
                   f"The push itself is finished — its errors are stored under run **{run_id}**.")
        return
    row = result.get("run") or {}
    if not row.get("end_time"):
        started = parse_ts(row.get("start_time"))
        since = ""
        if started:
            mins = int((datetime.now(timezone.utc) - started).total_seconds() // 60)
            since = f" — started {ist(row.get('start_time'))} IST, {mins} min ago"
        st.info(f"Uzio is still processing this run{since}. Its errors are written when it finishes — "
                "click **Check the run log again** in a minute or two.")
        return

    errors = run_issues(row.get("error_messages"))
    warnings = run_issues(row.get("optional_validations"))
    if not errors and not warnings:
        st.success("No errors or warnings were logged for this run.")
        return
    if errors:
        n = len(errors)
        st.markdown(_callout(
            "error", f"❌ {n} record{'s' if n != 1 else ''} failed",
            "The API skipped these — they were not created or updated. Fix them in the source and push again."),
            unsafe_allow_html=True)
        _render_issue_groups(errors, "error")
    if warnings:
        n = len(warnings)
        st.markdown(_callout(
            "warn", f"⚠️ {n} warning{'s' if n != 1 else ''}",
            "These records went through, but the API flagged something on each one — please review.",
            margin="16px 0 4px 0"),
            unsafe_allow_html=True)
        _render_issue_groups(warnings, "warning")

    issues = ([{"Severity": "Error", **e} for e in errors]
              + [{"Severity": "Warning", **w} for w in warnings])
    st.download_button(
        "📥 Download errors and warnings (CSV)",
        pd.DataFrame(issues).to_csv(index=False).encode("utf-8"),
        file_name=f"onboarding_run_{run_id}_issues.csv", mime="text/csv",
        key=f"{key_prefix}_issues_dl",
    )


def render_push_to_uzio_section(vendor: str, data_key: str, jt_key_prefix: str, key_prefix: str):
    """Renders the 'Push Census to Uzio (Production)' section.

    vendor: "ADP" or "PAYCOM" -- sent to the login endpoint.
    data_key: session_state key holding the cleaned census bytes, set by the Census
        Sanity tool's download-button block ({"xlsx":..., "csv":..., "audit":...}).
    jt_key_prefix: the key_prefix passed to job_title_mapper.render_streamlit_section
        for this same file -- used to retrieve the mapping CSV it stashed.
    """
    st.markdown("---")
    st.markdown("### \U0001F680 Push Census to Uzio (Production)")
    st.warning(
        "This sends the corrected census straight to the **live Uzio onboarding API** "
        "and creates/updates real employee records for the employer below. "
        "Double-check the FEIN and the files before submitting."
    )

    cached = st.session_state.get(data_key, {})
    census_bytes = cached.get("csv")
    if not census_bytes:
        st.info("Run the sanity check above first — the corrected census isn't ready yet.")
        return

    jt_mapping = st.session_state.get(f"{jt_key_prefix}_job_title_mapping")
    if not jt_mapping:
        st.info("Complete the **Amazon Job Title Mapping** table above first — this API requires it.")
        return

    fein = st.text_input(
        "Employer FEIN", key=f"{key_prefix}_fein",
        help="Same FEIN used to log into Uzio for this employer.",
    )
    col_u, col_p = st.columns(2)
    with col_u:
        username = st.text_input("Uzio Username", key=f"{key_prefix}_user")
    with col_p:
        password = st.text_input(
            "Uzio Password", type="password", key=f"{key_prefix}_pass",
            help="Never stored -- used only for this request.",
        )

    include_lic = st.checkbox(
        "Also include a License & Emergency Contact file (optional)",
        key=f"{key_prefix}_inc_lic",
    )
    lic_file = None
    if include_lic:
        lic_file = st.file_uploader(
            "License & Emergency Contact file", type=["csv", "xlsx"],
            key=f"{key_prefix}_lic_upload",
        )

    confirm = st.checkbox(
        f"I confirm this pushes to **PRODUCTION** for FEIN `{fein or '____'}` and I've reviewed the files.",
        key=f"{key_prefix}_confirm",
    )

    result_key = f"{key_prefix}_last_push"
    if st.button("\U0001F680 Push to Uzio", key=f"{key_prefix}_btn", disabled=not confirm):
        if not (fein.strip() and username.strip() and password):
            st.error("FEIN, Username, and Password are all required.")
            return
        try:
            with st.spinner("Logging in to Uzio..."):
                token = login(username.strip(), password, fein.strip(), vendor)
        except OnboardingAPIError as e:
            st.session_state.pop(result_key, None)
            st.error(f"❌ {e}")
            return

        result = new_result(fein, username)
        try:
            with st.spinner("Uploading census... (a large file can take a few minutes)"):
                lic_bytes = lic_file.getvalue() if lic_file else None
                lic_name = lic_file.name if lic_file else None
                resp = push_employee_census(
                    token,
                    census_bytes, f"{vendor}_Census.csv",
                    jt_mapping["csv"], jt_mapping["filename"],
                    lic_bytes, lic_name,
                )
            result.update(ok=resp.ok, status=resp.status_code, text=resp.text[:3000],
                          run_id=automation_id(resp))
        except OnboardingAPIError as e:
            result["push_error"] = str(e)
        if result["run_id"] is not None or lost_answer(result):
            with st.spinner("Reading the API's log for this run..."):
                attach_run_log(result, token)
        st.session_state[result_key] = result

    result = st.session_state.get(result_key)
    if not result:
        return
    render_push_result(result, key_prefix)

    if needs_recheck(result) and st.button("🔄 Check the run log again", key=f"{key_prefix}_recheck"):
        try:
            token = login(username.strip(), password, fein.strip(), vendor)
        except OnboardingAPIError as e:
            result["run_error"] = str(e)
        else:
            attach_run_log(result, token)
        st.session_state[result_key] = result
        st.rerun()


def render_push_payment_method_section(vendor: str, data_key: str, key_prefix: str):
    """Renders the 'Push Payment Methods to Uzio (Production)' section.

    vendor: "ADP" or "PAYCOM" — sent to the login endpoint, and it also decides which
        file part the API expects (see push_payment_method).
    data_key: session_state key holding the corrected payment file bytes, set by the
        Payment Method Sanity tool's download block as {"csv": ..., "filename": ...}.

    Deliberately asks for nothing the census push asks for beyond the credentials:
    this endpoint takes one file, no job title mapping.
    """
    st.markdown("---")
    st.markdown("### \U0001F680 Push Payment Methods to Uzio (Production)")
    st.warning(
        "This sends the corrected direct deposit file straight to the **live Uzio "
        "onboarding API** and sets real payment methods for the employer below. "
        "Double-check the FEIN and the file before submitting."
    )

    cached = st.session_state.get(data_key, {})
    payment_bytes = cached.get("csv")
    if not payment_bytes:
        st.info("Run the sanity check above first — the corrected file isn't ready yet.")
        return

    fein = st.text_input(
        "Employer FEIN", key=f"{key_prefix}_fein",
        help="Same FEIN used to log into Uzio for this employer.",
    )
    col_u, col_p = st.columns(2)
    with col_u:
        username = st.text_input("Uzio Username", key=f"{key_prefix}_user")
    with col_p:
        password = st.text_input(
            "Uzio Password", type="password", key=f"{key_prefix}_pass",
            help="Never stored — used only for this request.",
        )

    confirm = st.checkbox(
        f"I confirm this pushes to **PRODUCTION** for FEIN `{fein or '____'}` and I've "
        "reviewed the file.", key=f"{key_prefix}_confirm",
    )

    result_key = f"{key_prefix}_last_push"
    if st.button("\U0001F680 Push Payment Methods", key=f"{key_prefix}_btn",
                 disabled=not confirm):
        if not (fein.strip() and username.strip() and password):
            st.error("FEIN, Username, and Password are all required.")
            return
        try:
            with st.spinner("Logging in to Uzio..."):
                token = login(username.strip(), password, fein.strip(), vendor)
        except OnboardingAPIError as e:
            st.session_state.pop(result_key, None)
            st.error(f"❌ {e}")
            return

        result = new_result(fein, username)
        try:
            with st.spinner("Uploading payment methods..."):
                resp = push_payment_method(
                    token, vendor, payment_bytes,
                    cached.get("filename") or f"{vendor}_Payment_Methods.csv",
                )
            result.update(ok=resp.ok, status=resp.status_code, text=resp.text[:3000],
                          run_id=automation_id(resp))
        except OnboardingAPIError as e:
            result["push_error"] = str(e)
        if result["run_id"] is not None or lost_answer(result):
            with st.spinner("Reading the API's log for this run..."):
                attach_run_log(result, token)
        st.session_state[result_key] = result

    result = st.session_state.get(result_key)
    if not result:
        return
    render_push_result(result, key_prefix)

    if needs_recheck(result) and st.button("🔄 Check the run log again",
                                           key=f"{key_prefix}_recheck"):
        try:
            token = login(username.strip(), password, fein.strip(), vendor)
        except OnboardingAPIError as e:
            result["run_error"] = str(e)
        else:
            attach_run_log(result, token)
        st.session_state[result_key] = result
        st.rerun()
