"""Onboarding API run logs, one Amazon DSP client at a time.

Every census, prior payroll, payment, tax or deduction push writes a row to
`onboarding_automation_history`. Reading it meant a jumpserver session and a
DBeaver window, which implementors do not have, so every "why did my run fail"
landed on someone with database access.

The question people actually arrive with is "what has been run for THIS client" —
they know the client's name and nothing else. So the tool starts from a dropdown
of every client on the Amazon exchange and answers, for the one picked: which
implementor ran which API, when, and how it went. Any run opens to its
per-employee errors.

Sign-in takes a username and a password and nothing else. The onboarding token
does need a FEIN, but the tool has the client list by then and mints the token
itself, so nobody has to know one.

Read-only throughout, audited under the credentials entered, and nothing is
stored on the server: both tokens live in this browser session only.
"""
import io

import pandas as pd
import streamlit as st

from utils import neuronops_client as neuron
from utils import onboarding_query as obq
from utils.ui_components import _callout, render_premium_header

# Every DSP client sits on this exchange; it is what makes the list "Amazon".
AMAZON_EXCHANGE_ID = "EX-20243277-1b50-4035-821d-d0fcd9b895a9"
ALL_CLIENTS = "— All Amazon clients (last 7 days) —"

TOKEN_KEY = "obl_token"
NEURON_KEY = "obl_neuron"
USER_KEY = "obl_user"
HOST_KEY = "obl_host"
CLIENTS_KEY = "obl_clients"
RUNS_KEY = "obl_runs"
RUNS_FOR_KEY = "obl_runs_for"


def _amazon_clients(token) -> list:
    """Every client on the Amazon exchange, by name."""
    rows = neuron.query(token, "select fein, company_name, live_status, is_test_client "
                        "from employer_organization where exchange_id = "
                        f"'{AMAZON_EXCHANGE_ID}' and deleted = 0 order by company_name")
    return [r for r in rows if r.get("fein") and r.get("company_name")]


def _mint_onboarding_token(username, password, clients, host):
    """The token needs some FEIN; the user should not have to know one.

    Any FEIN the caller has access to unlocks the whole log, so the client list is
    walked until one is accepted rather than asking for one.
    """
    last_error = None
    for client in clients[:8]:
        try:
            return obq.login(username, password, str(client["fein"]), host=host), None
        except obq.OnboardingQueryError as e:
            last_error = e
    return None, last_error


def _sign_in():
    render_premium_header("Sign in with your Uzio credentials",
                          "Username and password only — the client list and the log are "
                          "both unlocked from these. Your password is used once and is "
                          "never stored.")
    with st.form("obl_login"):
        col1, col2 = st.columns(2)
        username = col1.text_input("Username", placeholder="firstname.lastname@uzio.com")
        password = col2.text_input("Password", type="password")
        host = st.text_input("Environment", value=obq.DEFAULT_HOST)
        submitted = st.form_submit_button("Sign in", type="primary")

    if not submitted:
        return
    if not (username and password):
        st.error("Enter your username and password.")
        return

    with st.spinner("Signing in…"):
        try:
            neuron_token = neuron.login(username.strip(), password)
        except neuron.NeuronOpsError as e:
            st.error(str(e))
            return
        try:
            clients = _amazon_clients(neuron_token)
        except neuron.NeuronOpsError as e:
            st.error(f"Signed in, but the client list could not be read: {e}")
            return
        if not clients:
            st.error("No clients came back for the Amazon exchange.")
            return
        token, failure = _mint_onboarding_token(username.strip(), password, clients,
                                                host.strip())
    if not token:
        st.error(f"Signed in to the reporting database, but the onboarding log refused "
                 f"the same credentials: {failure}")
        return

    st.session_state.update({NEURON_KEY: neuron_token, TOKEN_KEY: token,
                             CLIENTS_KEY: clients, USER_KEY: username.strip(),
                             HOST_KEY: host.strip()})
    st.rerun()


def _runs_for_client(token, host, fein, limit=300):
    where = f"fein = '{obq.digits(fein)}'" if fein else None
    if not where:                      # the all-clients view, kept to a week
        from datetime import date, timedelta
        begin, finish = obq.day_bounds_utc(date.today() - timedelta(days=7), date.today())
        where = f"start_time >= '{begin}' and start_time < '{finish}'"
    return obq.query(token, f"select {obq.LIGHT_COLUMNS} from {obq.TABLE} "
                     f"where {where} order by id desc limit {int(limit)}", host=host)


def _who_ran_what(rows) -> pd.DataFrame:
    """One line per API, showing the most recent run of it and who ran it.

    This is the answer to the question people arrive with: what has been set up for
    this client, by whom, and when.
    """
    latest = {}
    for row in rows:                   # newest first, so the first hit per module wins
        summary = obq.summarize(row)
        for module in summary["modules"]:
            if module in latest:
                latest[module]["Runs"] += 1
                continue
            failed = int(summary["failed"].get(module) or 0)
            # The run's own status covers every module in it, so a run that failed
            # elsewhere would libel a module that went through cleanly. Only borrow it
            # when there is no result at all.
            status = summary["status"] if not row.get("end_time") else (
                "OK" if failed == 0 else f"FAIL {failed}")
            latest[module] = {
                "API": obq.MODULE_SHORT.get(module, module),
                "Last run by": (row.get("created_by") or "").split("@")[0],
                "When (IST)": obq.ist(row.get("start_time"), "%d-%b-%Y %H:%M"),
                "Run": row.get("id"),
                "Total": summary["totals"].get(module),
                "OK": summary["passed"].get(module),
                "Fail": summary["failed"].get(module),
                "Status": status,
                "Runs": 1,
            }
    order = ["API", "Last run by", "When (IST)", "Total", "OK", "Fail", "Status",
             "Runs", "Run"]
    return pd.DataFrame(list(latest.values()))[order] if latest else pd.DataFrame()


def _runs_frame(rows, names=None) -> pd.DataFrame:
    out = []
    for row in rows:
        summary = obq.summarize(row)
        entry = {"Run": row.get("id"),
                 "Started (IST)": obq.ist(row.get("start_time")),
                 "Ran by": (row.get("created_by") or "").split("@")[0],
                 "APIs": ", ".join(obq.MODULE_SHORT.get(m, m) for m in summary["modules"]),
                 "Vendor": (row.get("vendor") or "").upper().replace("PAYCOM", "Paycom"),
                 "Total": summary["total"], "OK": summary["ok"], "Fail": summary["fail"],
                 "Took": summary["duration"], "Status": summary["status"]}
        if names is not None:
            fein = str(row.get("fein") or "")
            entry = {"Client": names.get(fein, fein), **entry}
        out.append(entry)
    return pd.DataFrame(out)


def _run_detail(token, host, run_id, client_label):
    rows = obq.query(token, f"select * from {obq.TABLE} where id = {int(run_id)}", host=host)
    if not rows:
        st.error(f"No run with id {run_id}.")
        return
    row = rows[0]
    summary = obq.summarize(row)

    st.markdown(_callout("ok" if summary["status"] == "OK" else "warn",
                         f"Run {run_id} — {client_label}",
                         f"{(row.get('vendor') or '').upper()} · started "
                         f"{obq.ist(row.get('start_time'), '%d-%b-%Y %H:%M:%S')} IST · "
                         f"by {row.get('created_by') or '?'} · took "
                         f"{summary['duration'] or '—'} · {summary['status']}"),
                unsafe_allow_html=True)
    if not row.get("end_time"):
        st.warning("This run has no end time and no result: it is still going, or it "
                   "died without writing one.")

    if summary["totals"]:
        st.dataframe(pd.DataFrame([{
            "API": obq.MODULE_SHORT.get(m, m),
            "Total": summary["totals"].get(m),
            "OK": summary["passed"].get(m),
            "Fail": summary["failed"].get(m),
        } for m in summary["totals"]]), hide_index=True, use_container_width=True)

    errors = obq.issue_rows(row.get("error_messages"))
    warnings = obq.issue_rows(row.get("optional_validations"))
    if not errors and not warnings:
        st.success("No per-employee errors or warnings were written for this run.")

    for label, issues, tone in (
            ("Errors — these employees did not go through", errors, "error"),
            ("Warnings — these went through, but check them", warnings, "warn")):
        if not issues:
            continue
        st.markdown(_callout(tone, f"{label}  ({len(issues)} rows)",
                             "Grouped by reason — the same reason hit by many employees "
                             "is one line here; the full list is below and in the "
                             "download."), unsafe_allow_html=True)
        st.dataframe(pd.DataFrame(obq.group_issues(issues)), hide_index=True,
                     use_container_width=True)
        with st.expander(f"Every row ({len(issues)})"):
            st.dataframe(pd.DataFrame(issues), hide_index=True, use_container_width=True)

    if errors or warnings:
        frame = pd.DataFrame([{**i, "Kind": kind} for kind, group in
                              (("Error", errors), ("Warning", warnings)) for i in group])
        book = io.BytesIO()
        with pd.ExcelWriter(book, engine="xlsxwriter") as writer:
            frame.to_excel(writer, sheet_name="Issues", index=False)
            if errors:
                pd.DataFrame(obq.group_issues(errors)).to_excel(
                    writer, sheet_name="Errors grouped", index=False)
        col1, col2 = st.columns(2)
        col1.download_button(f"Download run {run_id} issues (.xlsx)", book.getvalue(),
                             file_name=f"Onboarding_run_{run_id}_issues.xlsx",
                             mime="application/vnd.openxmlformats-officedocument."
                                  "spreadsheetml.sheet", type="primary")
        # Plain UTF-8, no BOM: these CSVs get fed back into other tools.
        col2.download_button(f"Download run {run_id} issues (.csv)",
                             frame.to_csv(index=False).encode("utf-8"),
                             file_name=f"Onboarding_run_{run_id}_issues.csv",
                             mime="text/csv")


def render_ui():
    st.title("Onboarding API Run Logs")
    st.markdown(
        "Pick an Amazon DSP client and see **which implementor ran which API, when, and "
        "how it went** — then open any run for its per-employee errors. Reads the "
        "onboarding API's own log, so nobody needs a jumpserver and DBeaver."
    )

    token = st.session_state.get(TOKEN_KEY)
    if not token:
        st.markdown(_callout("warn", "Read-only, and audited under your own login",
                             "One SELECT against the reporting endpoints — it cannot "
                             "change anything, and the read is recorded against the "
                             "credentials you sign in with. Nothing is saved on the "
                             "server; the tokens live in this browser session only."),
                    unsafe_allow_html=True)
        _sign_in()
        return

    host = st.session_state.get(HOST_KEY) or obq.DEFAULT_HOST
    clients = st.session_state.get(CLIENTS_KEY) or []
    col1, col2 = st.columns([4, 1])
    col1.caption(f"Signed in as **{st.session_state.get(USER_KEY, '')}** · {host} · "
                 f"{len(clients)} Amazon clients")
    if col2.button("Sign out"):
        for key in (TOKEN_KEY, NEURON_KEY, USER_KEY, HOST_KEY, CLIENTS_KEY, RUNS_KEY,
                    RUNS_FOR_KEY):
            st.session_state.pop(key, None)
        st.rerun()

    labels = {ALL_CLIENTS: None}
    names = {}
    for client in clients:
        fein = str(client["fein"])
        mark = "  ·  test client" if client.get("is_test_client") else ""
        labels[f"{client['company_name']}  ({fein}){mark}"] = fein
        names[fein] = client["company_name"]

    col1, col2 = st.columns([3, 1])
    picked = col1.selectbox("Client", list(labels),
                            help="Start typing to search. Every client on the Amazon "
                                 "exchange is here.")
    only_failures = col2.checkbox("Failures only")
    if st.button("Show runs", type="primary"):
        with st.spinner("Reading the log…"):
            try:
                st.session_state[RUNS_KEY] = _runs_for_client(token, host, labels[picked])
                st.session_state[RUNS_FOR_KEY] = picked
            except obq.OnboardingQueryError as e:
                st.error(str(e))
                return

    rows = st.session_state.get(RUNS_KEY)
    if rows is None:
        return
    shown_for = st.session_state.get(RUNS_FOR_KEY, picked)
    if only_failures:
        rows = [r for r in rows if (obq.summarize(r)["fail"] or 0) > 0]
    if not rows:
        st.info(f"No runs found for {shown_for}."
                + (" Untick “Failures only” to see the successful ones."
                   if only_failures else " This client has never been pushed to."))
        return

    client_fein = labels.get(shown_for)
    if client_fein:
        render_premium_header(f"What has been run for {names.get(client_fein, shown_for)}",
                              "One line per API — the most recent run of it, who ran it, "
                              "and how many runs there have been in total.")
        summary = _who_ran_what(rows)
        if not summary.empty:
            st.dataframe(summary, hide_index=True, use_container_width=True)

    render_premium_header(f"All {len(rows)} run(s), newest first",
                          "Pick one below to see its per-employee errors.")
    frame = _runs_frame(rows, None if client_fein else names)
    st.dataframe(frame, hide_index=True, use_container_width=True)

    options = {f"{r['Run']} — {r['APIs'] or '?'} — {r['Started (IST)']} — "
               f"by {r['Ran by']} — {r['Status']}": r["Run"] for _, r in frame.iterrows()}
    chosen = st.selectbox("Open a run", list(options))
    if st.button("Show this run's errors", type="primary"):
        with st.spinner("Reading that run…"):
            try:
                _run_detail(token, host, options[chosen],
                            names.get(client_fein, shown_for) if client_fein
                            else "all clients")
            except obq.OnboardingQueryError as e:
                st.error(str(e))
