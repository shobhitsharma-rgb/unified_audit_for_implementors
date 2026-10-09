"""Onboarding API run logs — the jumpserver/DBeaver view, on screen.

Every census, prior payroll, payment, tax and deduction push writes a row to
`onboarding_automation_history`. Today the only way to read it is a jumpserver
session and a DBeaver window, which implementors do not have, so every "why did
my run fail" lands on someone else's desk. This shows the same thing from the
tool: who ran it, when, how many went through, how many failed, and exactly
which employees failed for what reason.

Sign-in mints two tokens from ONE set of credentials:
  * the onboarding token (needs any FEIN you have access to) — the log itself
  * the NeuronOps token — only to turn a FEIN into a client name

Everything is read-only and audited under the credentials entered. Nothing is
stored on the server: both tokens live in this browser session only.
"""
import io
from datetime import date, timedelta

import pandas as pd
import streamlit as st

from utils import neuronops_client as neuron
from utils import onboarding_query as obq
from utils.ui_components import _callout, render_premium_header

TOKEN_KEY = "obl_token"
NEURON_KEY = "obl_neuron"
USER_KEY = "obl_user"
HOST_KEY = "obl_host"
RUNS_KEY = "obl_runs"
NAMES_KEY = "obl_names"

VENDORS = {"All": None, "ADP": "ADP", "Paycom": "PAYCOM"}
STATUSES = ["All", "Failures only", "Still running / no result"]


def _sign_in():
    render_premium_header("Sign in with your Uzio credentials",
                          "The same login you use for the onboarding API. Your password "
                          "is used once to mint a token and is never stored.")
    with st.form("obl_login"):
        col1, col2, col3 = st.columns(3)
        username = col1.text_input("Username", placeholder="firstname.lastname@uzio.com")
        password = col2.text_input("Password", type="password")
        fein = col3.text_input("Any FEIN you have access to", placeholder="863131339",
                               help="The token is minted against one employer, but the "
                                    "log it reads covers every client.")
        host = st.text_input("Environment", value=obq.DEFAULT_HOST)
        submitted = st.form_submit_button("Sign in", type="primary")

    if not submitted:
        return
    if not (username and password and fein):
        st.error("Enter your username, password and a FEIN.")
        return
    with st.spinner("Signing in…"):
        try:
            st.session_state[TOKEN_KEY] = obq.login(username.strip(), password,
                                                    obq.digits(fein) or fein.strip(),
                                                    host=host.strip())
        except obq.OnboardingQueryError as e:
            st.error(str(e))
            return
        # Client names come from a different database, so this one is optional:
        # without it the log still reads, it just shows FEINs.
        try:
            st.session_state[NEURON_KEY] = neuron.login(username.strip(), password)
        except neuron.NeuronOpsError:
            st.session_state[NEURON_KEY] = ""
    st.session_state[USER_KEY] = username.strip()
    st.session_state[HOST_KEY] = host.strip()
    st.rerun()


def _client_names(feins) -> dict:
    """{fein -> company name}, as far as they can be resolved."""
    known = st.session_state.setdefault(NAMES_KEY, {})
    wanted = {obq.digits(f) for f in feins if obq.digits(f)} - set(known)
    for fein in wanted & set(obq.SANDBOX_FEINS):
        known[fein] = obq.SANDBOX_FEINS[fein]
    wanted -= set(obq.SANDBOX_FEINS)
    token = st.session_state.get(NEURON_KEY)
    if wanted and token:
        quoted = neuron.sql_in_list(wanted)
        if quoted:
            try:
                for row in neuron.query(token, "select fein, company_name from "
                                        f"employer_organization where fein in ({quoted})"):
                    known[str(row.get("fein"))] = row.get("company_name") or ""
            except neuron.NeuronOpsError:
                pass                      # names are a nicety, the log is the point
    for fein in wanted:
        known.setdefault(fein, "")
    return known


def _fetch_runs(token, host, start_day, end_day, vendor, fein, user, limit):
    where = []
    begin, finish = obq.day_bounds_utc(start_day, end_day)
    where.append(f"start_time >= '{begin}' and start_time < '{finish}'")
    if vendor:
        where.append(f"upper(vendor) = '{vendor}'")
    if fein:
        where.append(f"fein = '{fein}'")
    if user:
        where.append(f"lower(created_by) like lower('%{user}%')")
    sql = (f"select {obq.LIGHT_COLUMNS} from {obq.TABLE} where " + " and ".join(where)
           + f" order by id desc limit {int(limit)}")
    return obq.query(token, sql, host=host)


def _runs_frame(rows, names) -> pd.DataFrame:
    out = []
    for row in rows:
        summary = obq.summarize(row)
        fein = str(row.get("fein") or "")
        out.append({
            "Run": row.get("id"),
            "Client": names.get(obq.digits(fein), "") or "(unknown)",
            "FEIN": fein,
            "Started (IST)": obq.ist(row.get("start_time")),
            "Vendor": (row.get("vendor") or "").upper().replace("PAYCOM", "Paycom"),
            "Modules": ", ".join(obq.MODULE_SHORT.get(m, m) for m in summary["modules"]),
            "Total": summary["total"],
            "OK": summary["ok"],
            "Fail": summary["fail"],
            "Ran by": (row.get("created_by") or "").split("@")[0],
            "Took": summary["duration"],
            "Status": summary["status"],
        })
    return pd.DataFrame(out)


def _filters():
    today = date.today()
    with st.form("obl_filters"):
        col1, col2, col3 = st.columns([2, 1, 1])
        span = col1.date_input("Started between (IST)",
                               value=(today - timedelta(days=7), today), max_value=today)
        vendor = VENDORS[col2.selectbox("Vendor", list(VENDORS))]
        status = col3.selectbox("Show", STATUSES)
        col4, col5, col6 = st.columns([2, 2, 1])
        client = col4.text_input("Client name or FEIN", placeholder="Express Package, or 863131339")
        user = col5.text_input("Ran by", placeholder="tierra")
        limit = col6.number_input("Max runs", 10, 1000, 200, step=10)
        go = st.form_submit_button("Show runs", type="primary")
    if not go:
        return None
    start_day, end_day = span if isinstance(span, (tuple, list)) and len(span) == 2 else (span, span)
    return {"start": start_day, "end": end_day, "vendor": vendor, "status": status,
            "client": client.strip(), "user": obq.sql_literal(user), "limit": limit}


def _resolve_client(text):
    """A typed client -> FEIN. A name needs the NeuronOps token; a FEIN never does."""
    if not text:
        return "", None
    fein = obq.digits(text)
    if fein:
        return fein, None
    token = st.session_state.get(NEURON_KEY)
    if not token:
        return "", ("Searching by client name needs the second sign-in, which did not "
                    "go through. Enter the 9-digit FEIN instead.")
    safe = obq.sql_literal(text)
    try:
        found = neuron.query(token, "select fein, company_name from employer_organization "
                             f"where lower(company_name) like lower('%{safe}%') and deleted = 0")
    except neuron.NeuronOpsError as e:
        return "", str(e)
    feins = {str(r.get("fein")): r.get("company_name") for r in found if r.get("fein")}
    if not feins:
        return "", f"No client matching “{text}”."
    if len(feins) > 1:
        listed = ", ".join(f"{name} ({fein})" for fein, name in list(feins.items())[:6])
        return "", f"That matches more than one client: {listed}. Use the FEIN."
    return next(iter(feins)), None


def _run_detail(token, host, run_id, names):
    rows = obq.query(token, f"select * from {obq.TABLE} where id = {int(run_id)}", host=host)
    if not rows:
        st.error(f"No run with id {run_id}.")
        return
    row = rows[0]
    summary = obq.summarize(row)
    fein = str(row.get("fein") or "")
    name = names.get(obq.digits(fein), "") or "(unknown client)"

    st.markdown(_callout("ok" if summary["status"] == "OK" else "warn",
                         f"Run {run_id} — {name} ({fein})",
                         f"{(row.get('vendor') or '').upper()} · started "
                         f"{obq.ist(row.get('start_time'), '%d-%b-%Y %H:%M:%S')} IST · "
                         f"by {row.get('created_by') or '?'} · took {summary['duration'] or '—'} · "
                         f"{summary['status']}"), unsafe_allow_html=True)
    if not row.get("end_time"):
        st.warning("This run has no end time and no result: it is still going, or it died "
                   "without writing one.")

    if summary["totals"]:
        st.dataframe(pd.DataFrame([{
            "Module": obq.MODULE_SHORT.get(m, m),
            "Total": summary["totals"].get(m),
            "OK": summary["passed"].get(m),
            "Fail": summary["failed"].get(m),
        } for m in summary["totals"]]), hide_index=True, use_container_width=True)

    errors = obq.issue_rows(row.get("error_messages"))
    warnings = obq.issue_rows(row.get("optional_validations"))
    if not errors and not warnings:
        st.success("No per-employee errors or warnings were written for this run.")

    for label, issues, tone in (("Errors — these employees did not go through", errors, "error"),
                                ("Warnings — these went through, but check them", warnings, "warn")):
        if not issues:
            continue
        st.markdown(_callout(tone, f"{label}  ({len(issues)} rows)",
                             "Grouped by reason — the same reason hit by many employees is "
                             "one line here; the full list is below and in the download."),
                    unsafe_allow_html=True)
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
                             file_name=f"Onboarding_run_{run_id}_issues.csv", mime="text/csv")


def render_ui():
    st.title("Onboarding API Run Logs")
    st.markdown(
        "Every census, prior payroll, payment, tax or deduction push writes a row to "
        "`onboarding_automation_history`. This reads it directly — **who ran it, when, "
        "how many went through, how many failed and why** — so nobody needs a "
        "jumpserver and DBeaver to answer that."
    )

    token = st.session_state.get(TOKEN_KEY)
    if not token:
        st.markdown(_callout("warn", "Read-only, and audited under your own login",
                             "One SELECT against the reporting endpoint — it cannot change "
                             "anything, and the read is recorded against the credentials you "
                             "sign in with. Nothing is saved on the server; your password is "
                             "used once and the tokens live in this browser session only."),
                    unsafe_allow_html=True)
        _sign_in()
        return

    host = st.session_state.get(HOST_KEY) or obq.DEFAULT_HOST
    col1, col2 = st.columns([4, 1])
    names_on = "client names on" if st.session_state.get(NEURON_KEY) else "FEIN only"
    col1.caption(f"Signed in as **{st.session_state.get(USER_KEY, '')}** · {host} · {names_on}")
    if col2.button("Sign out"):
        for key in (TOKEN_KEY, NEURON_KEY, USER_KEY, HOST_KEY, RUNS_KEY, NAMES_KEY):
            st.session_state.pop(key, None)
        st.rerun()

    if not st.session_state.get(NEURON_KEY):
        st.info("Client names could not be switched on (the second sign-in was refused), so "
                "runs show their FEIN. Everything else works.")

    chosen = _filters()
    if chosen:
        fein, problem = _resolve_client(chosen["client"])
        if problem:
            st.error(problem)
        else:
            with st.spinner("Reading the log…"):
                try:
                    rows = _fetch_runs(token, host, chosen["start"], chosen["end"],
                                       chosen["vendor"], fein, chosen["user"], chosen["limit"])
                except obq.OnboardingQueryError as e:
                    st.error(str(e))
                    rows = None
            if rows is not None:
                if chosen["status"] == "Failures only":
                    rows = [r for r in rows if (obq.summarize(r)["fail"] or 0) > 0]
                elif chosen["status"] == "Still running / no result":
                    rows = [r for r in rows if not r.get("end_time")]
                st.session_state[RUNS_KEY] = rows

    rows = st.session_state.get(RUNS_KEY)
    if rows is None:
        return
    if not rows:
        st.info("No runs matched. Widen the dates, or clear a filter.")
        return

    names = _client_names({r.get("fein") for r in rows})
    frame = _runs_frame(rows, names)
    render_premium_header(f"{len(frame)} run(s)",
                          "Newest first. Pick a run below to see its per-employee errors.")
    st.dataframe(frame, hide_index=True, use_container_width=True)

    labels = {f"{r['Run']} — {r['Client']} — {r['Modules'] or '?'} — {r['Started (IST)']} "
              f"— {r['Status']}": r["Run"] for _, r in frame.iterrows()}
    picked = st.selectbox("Open a run", list(labels), index=0)
    if st.button("Show this run's errors", type="primary"):
        with st.spinner("Reading that run…"):
            try:
                _run_detail(token, host, labels[picked], names)
            except obq.OnboardingQueryError as e:
                st.error(str(e))
