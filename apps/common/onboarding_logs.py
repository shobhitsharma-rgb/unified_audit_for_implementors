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

TOKEN_KEY = "obl_token"
NEURON_KEY = "obl_neuron"
USER_KEY = "obl_user"
HOST_KEY = "obl_host"
CLIENTS_KEY = "obl_clients"
RUNS_KEY = "obl_runs"
RUNS_FOR_KEY = "obl_runs_for"
SHOW_KEY = "obl_show"


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
    return obq.query(token, f"select {obq.LIGHT_COLUMNS} from {obq.TABLE} "
                     f"where fein = '{obq.digits(fein)}' order by id desc "
                     f"limit {int(limit)}", host=host)


def _api_status(rows) -> list:
    """One record per API: did it go through, when, and who ran it.

    Keyed on the most recent run of each API, because that is the one that counts -
    an earlier failure that has since been re-run is history, not a problem.
    """
    latest = {}
    for row in rows:                   # newest first, so the first hit per API wins
        summary = obq.summarize(row)
        for module in summary["modules"]:
            if module in latest:
                latest[module]["runs"] += 1
                continue
            failed = int(summary["failed"].get(module) or 0)
            total = int(summary["totals"].get(module) or 0)
            latest[module] = {
                "api": obq.MODULE_SHORT.get(module, module),
                "module": module,
                "run": row.get("id"),
                "by": (row.get("created_by") or "").split("@")[0],
                "when": obq.ist(row.get("start_time"), "%d-%b %H:%M"),
                "total": total, "ok": int(summary["passed"].get(module) or 0),
                "fail": failed,
                # No end time means no result was ever written, which is not the same
                # as a clean run and must not be shown as one.
                "unfinished": not row.get("end_time"),
                # Finished having processed nobody: also not a pass.
                "empty": bool(row.get("end_time")) and total == 0,
                "runs": 1,
            }
    records = list(latest.values())
    records.sort(key=lambda r: (not (r["fail"] or r["unfinished"] or r["empty"]),
                                r["api"]))
    return records


def _runs_frame(rows) -> pd.DataFrame:
    """The full history, for the expander at the bottom."""
    out = []
    for row in rows:
        summary = obq.summarize(row)
        out.append({"Run": row.get("id"),
                    "Started (IST)": obq.ist(row.get("start_time")),
                    "Ran by": (row.get("created_by") or "").split("@")[0],
                    "APIs": ", ".join(obq.MODULE_SHORT.get(m, m)
                                      for m in summary["modules"]),
                    "Total": summary["total"], "OK": summary["ok"],
                    "Fail": summary["fail"], "Took": summary["duration"],
                    "Status": summary["status"]})
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


def _errors_for(token, host, run_id, api, client_label):
    """One run of one API: what failed in it, grouped, with the detail tucked away."""
    rows = obq.query(token, f"select * from {obq.TABLE} where id = {int(run_id)}",
                     host=host)
    if not rows:
        st.error(f"Run {run_id} is no longer in the log.")
        return
    row = rows[0]
    errors = [i for i in obq.issue_rows(row.get("error_messages"))
              if i["Module"] == api or not i["Module"]]
    warnings = [i for i in obq.issue_rows(row.get("optional_validations"))
                if i["Module"] == api or not i["Module"]]

    st.caption(f"{client_label} · {api} · run {run_id} · "
               f"{obq.ist(row.get('start_time'), '%d-%b-%Y %H:%M')} IST · "
               f"by {(row.get('created_by') or '?').split('@')[0]}")

    if not errors:
        st.info("This run wrote no per-employee failures for this API.")
    else:
        st.markdown(f"**{len(errors)} employees did not go through.** Same reason, one "
                    "line — fix the reason and they all clear.")
        st.dataframe(pd.DataFrame(obq.group_issues(errors))[
            ["Employees", "Reason", "e.g."]], hide_index=True, use_container_width=True)

    frame = pd.DataFrame([{**i, "Kind": kind} for kind, group in
                          (("Error", errors), ("Warning", warnings)) for i in group])
    if not frame.empty:
        book = io.BytesIO()
        with pd.ExcelWriter(book, engine="xlsxwriter") as writer:
            frame.to_excel(writer, sheet_name="Issues", index=False)
            if errors:
                pd.DataFrame(obq.group_issues(errors)).to_excel(
                    writer, sheet_name="Grouped", index=False)
        st.download_button(f"Download the full list ({len(frame)} rows)", book.getvalue(),
                           file_name=f"{api}_run_{run_id}_issues.xlsx",
                           key=f"dl_{run_id}_{api}",
                           mime="application/vnd.openxmlformats-officedocument."
                                "spreadsheetml.sheet", type="primary")

    if errors:
        with st.expander(f"Every employee ({len(errors)})"):
            st.dataframe(pd.DataFrame(errors)[["Employee ID", "Row", "Reason"]],
                         hide_index=True, use_container_width=True)
    if warnings:
        with st.expander(f"Warnings — these went through anyway ({len(warnings)})"):
            st.dataframe(pd.DataFrame(obq.group_issues(warnings))[
                ["Employees", "Reason", "e.g."]], hide_index=True,
                use_container_width=True)


def in_flight(rows):
    """Runs that never wrote a result, split into "still going" and "gave up".

    Until a run finishes it has no response_body, so it has no TotalMap, so it
    belongs to no API and produces no card. Without this it would be invisible on
    the main screen - and a push that is still going is exactly the thing someone
    refreshing this page wants to know about. The log does not say which API is in
    flight; that only arrives with the result.

    Anything older than two hours has not been running for two hours, it died.
    """
    from datetime import datetime, timezone

    now = datetime.now(timezone.utc)
    going, died = [], []
    for row in rows:
        if row.get("end_time"):
            continue
        started = obq.parse_ts(row.get("start_time"))
        minutes = int((now - started).total_seconds() // 60) if started else 0
        entry = {"run": row.get("id"), "minutes": minutes,
                 "by": (row.get("created_by") or "").split("@")[0],
                 "when": obq.ist(row.get("start_time"), "%d-%b %H:%M")}
        (going if minutes <= 120 else died).append(entry)
    return going, died


def api_runs(rows, module):
    """Every run that included `module`, newest first, with that module's own counts.

    Returns the table to show and {label -> (run id, failures)} for the picker. Kept
    out of the Streamlit function so the status of each attempt can be tested.
    """
    table, options = {}, {}
    for row in rows:                                   # newest first
        summary = obq.summarize(row)
        if module not in summary["modules"]:
            continue
        failed = int(summary["failed"].get(module) or 0)
        total = int(summary["totals"].get(module) or 0)
        if not row.get("end_time"):
            status = summary["status"]
        elif total == 0:
            status = obq.NOTHING_PROCESSED
        else:
            status = "OK" if failed == 0 else f"FAIL {failed}"
        when = obq.ist(row.get("start_time"), "%d-%b-%Y %H:%M")
        by = (row.get("created_by") or "").split("@")[0]
        table[row.get("id")] = {"Run": row.get("id"), "When (IST)": when, "Ran by": by,
                                "Total": summary["totals"].get(module),
                                "OK": summary["passed"].get(module),
                                "Fail": summary["failed"].get(module), "Status": status}
        options[f"{row.get('id')} — {when} — by {by} — {status}"] = (row.get("id"), failed)
    return list(table.values()), options


def _api_section(token, host, rows, module, api, client_name):
    """Every run of ONE API, oldest attempts included, any of them openable.

    An API that was run fifteen times has fifteen stories, not one: which attempt
    fixed what, and what was still broken in the one before it. The headline card
    is the latest run; this is all of them.
    """
    table, options = api_runs(rows, module)
    st.markdown(f"#### {api} — {len(table)} run(s)")
    st.dataframe(pd.DataFrame(table), hide_index=True, use_container_width=True)

    # Land on the newest run that actually failed - that is the one being chased.
    failed_first = next((i for i, (_, failed) in enumerate(options.values()) if failed), 0)
    chosen = st.selectbox("Open a run", list(options), index=failed_first,
                          key=f"pick_{module}")
    _errors_for(token, host, options[chosen][0], api, client_name)


def render_ui():
    st.title("Onboarding API Run Logs")
    st.markdown(
        "Pick a client and see **which APIs went through and which did not** — and for "
        "the ones that did not, exactly which employees failed and why, ready to "
        "download and fix."
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
    col1.caption(f"Signed in as **{st.session_state.get(USER_KEY, '')}** · "
                 f"{len(clients)} Amazon clients")
    if col2.button("Sign out"):
        for key in (TOKEN_KEY, NEURON_KEY, USER_KEY, HOST_KEY, CLIENTS_KEY, RUNS_KEY,
                    RUNS_FOR_KEY, SHOW_KEY):
            st.session_state.pop(key, None)
        st.rerun()

    labels = {}
    for client in clients:
        fein = str(client["fein"])
        mark = "  ·  test client" if client.get("is_test_client") else ""
        labels[f"{client['company_name']}  ({fein}){mark}"] = (fein, client["company_name"])

    with st.expander("How to read this"):
        st.markdown(
            "- **One line per API, showing its latest run** — that is the state the "
            "client is in now. The button beside it opens **every run of that API**, so "
            "if Census was run fifteen times you can read each of the fifteen: what "
            "failed, what the next attempt fixed, and what was still broken.\n"
            "- **Green means every employee in that run went through.** Red says how "
            "many did not: \"58 of 1618 employees failed\" means 1560 are in Uzio and 58 "
            "are not.\n"
            "- Inside, pick any run and you get **its** reasons, not the rows. The same "
            "reason usually hits "
            "many employees at once, so it is one line with a count — fix that one thing "
            "and all of them clear on the next run. The download has every employee if "
            "you need to work through them.\n"
            "- **Warnings are not failures.** Those employees went through; Uzio just "
            "filled something in for them (a blank amount defaulted to 0, say). They sit "
            "in an expander because they rarely need action.\n"
            "- **A run that is still going** appears as a banner at the top, not as a "
            "line — the log does not record which API it was until it finishes. Press "
            "the button again a few minutes later and it will have joined the list.\n"
            "- **\"Nothing processed\"** means a run finished but no employee went in: "
            "an empty file, or nobody matched. **Runs that never wrote a result** died "
            "partway and are called out separately. Neither is a pass.\n"
            "- Times are **IST**, and the name beside each line is whoever ran it.")

    picked = st.selectbox("Client", list(labels),
                          help="Start typing to search. Every client on the Amazon "
                               "exchange is here.")
    if st.button("Show me how the APIs went", type="primary"):
        with st.spinner("Reading the log…"):
            try:
                st.session_state[RUNS_KEY] = _runs_for_client(token, host, labels[picked][0])
                st.session_state[RUNS_FOR_KEY] = picked
                st.session_state.pop(SHOW_KEY, None)
            except obq.OnboardingQueryError as e:
                st.error(str(e))
                return

    rows = st.session_state.get(RUNS_KEY)
    if rows is None:
        return
    shown_for = st.session_state.get(RUNS_FOR_KEY, picked)
    client_name = labels.get(shown_for, ("", shown_for))[1]
    if not rows:
        st.info(f"Nothing has ever been pushed for {client_name}.")
        return

    records = _api_status(rows)
    bad = [r for r in records if r["fail"] or r["unfinished"] or r["empty"]]
    render_premium_header(
        client_name,
        ("Every API went through." if not bad else
         f"{len(bad)} of {len(records)} APIs need attention."))

    going, died = in_flight(rows)
    for entry in going:
        minutes = entry["minutes"]
        age = "just now" if minutes < 1 else f"{minutes} minute(s) ago"
        st.warning(f"⏳ **A run is still going** — run {entry['run']}, started {age} by "
                   f"{entry['by']}. The log only says which API it was once it "
                   f"finishes, so it is not in the list below yet. Press **Show me how "
                   f"the APIs went** again in a few minutes.")
    if died:
        st.info(f"{len(died)} earlier run(s) never wrote a result — "
                + ", ".join(f"run {e['run']} ({e['when']}, {e['by']})" for e in died[:3])
                + (" and others" if len(died) > 3 else "")
                + ". They died partway, so whatever they did is not counted below.")

    for record in records:
        col1, col2 = st.columns([5, 1])
        if record["unfinished"]:
            col1.warning(f"**{record['api']}** — never finished. Started "
                         f"{record['when']} IST by {record['by']}.")
        elif record["empty"]:
            col1.info(f"**{record['api']}** — the last run processed no employees at "
                      f"all. Run on {record['when']} IST by {record['by']}.")
        elif record["fail"]:
            col1.error(f"**{record['api']}** — **{record['fail']} of {record['total']} "
                       f"employees failed**. Run on {record['when']} IST by "
                       f"{record['by']}.")
        else:
            col1.success(f"**{record['api']}** — all {record['total']} went through. "
                         f"Run on {record['when']} IST by {record['by']}.")
        if record["runs"] > 1:
            label = f"See all {record['runs']} runs"
        elif record["fail"] or record["unfinished"] or record["empty"]:
            label = "See why"
        else:
            label = None                   # one clean run needs no second look
        if label and col2.button(label, key=f"open_{record['module']}"):
            st.session_state[SHOW_KEY] = record["module"]

    chosen = st.session_state.get(SHOW_KEY)
    if chosen:
        st.divider()
        api = obq.MODULE_SHORT.get(chosen, chosen)
        with st.spinner(f"Reading the {api} runs…"):
            try:
                _api_section(token, host, rows, chosen, api, client_name)
            except obq.OnboardingQueryError as e:
                st.error(str(e))

    with st.expander(f"Every run for this client ({len(rows)})"):
        st.caption("Including the earlier attempts. The cards above only show the most "
                   "recent run of each API.")
        st.dataframe(_runs_frame(rows), hide_index=True, use_container_width=True)
