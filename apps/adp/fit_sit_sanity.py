import io
from collections import Counter, OrderedDict

import pandas as pd
import streamlit as st

from utils import state_filing_status as sfs

APP_TITLE = "ADP FIT/SIT Sanity Check"

# =========================================================
# ADP FIT/SIT Sanity Check
# - Input: single ADP FIT/SIT export (.csv / .xlsx)
# - Fills blanks in Dependents and Non-Resident Alien with fixed defaults.
# - Checks State Marital Status Description against what Uzio's onboarding API
#   actually accepts for that employee's WORKED-IN state, and asks the user to
#   map anything the API would reject. It used to fill every blank with the
#   literal "Single", which is invalid in several states -- Missouri takes only
#   "Single or Married Spouse Works or Married Filing Separate" -- so those rows
#   failed the import with "Invalid State filing status: 'Single' for state: MO".
# - Everything else is handled downstream by the API.
# =========================================================

STATUS_COL = "State Marital Status Description"

# The API binds this column to worksInState, and worksInState is what
# ADPStateTaxWithholdingValidator passes to parseStateFilingStatus. There is
# deliberately no fallback to "State Tax Code": the two disagree for multi-state
# employees, and classifying against a different state than the API will use
# would let this tool bless a value the API then rejects.
STATE_COL = "Worked in State Code"

DEFAULTS = {
    "Dependents": "0",
    "Non-Resident Alien": "No",
}

LEAVE_AS_IS = "— leave as is —"

# Row buckets, in the order the summary lists them.
ACCEPTED = "accepted"
PUNCTUATION = "punctuation"
NEEDS_DECISION = "decide"
NO_TABLE = "no_table"
NO_SIT = "no_sit"
NO_STATE = "no_state"

BUCKET_LABELS = OrderedDict([
    (ACCEPTED, "Already accepted by Uzio"),
    (PUNCTUATION, "Punctuation corrected automatically"),
    (NEEDS_DECISION, "Needs your mapping"),
    (NO_TABLE, "State has no Uzio filing-status list — left alone"),
    (NO_SIT, "No state income tax — left alone"),
    (NO_STATE, "No worked-in state on the row — left alone"),
])


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    return s == "" or s.lower() == "nan"


def _clean(v) -> str:
    return "" if _is_blank(v) else str(v).strip()


def _read_file(uploaded) -> pd.DataFrame:
    uploaded.seek(0)
    name = (uploaded.name or "").lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded, dtype=str)
    else:
        df = pd.read_excel(uploaded, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    return df


def _find_col(df: pd.DataFrame, target: str) -> str:
    """Exact match first, then case-insensitive."""
    if target in df.columns:
        return target
    target_lower = target.casefold()
    for c in df.columns:
        if c.casefold() == target_lower:
            return c
    return ""


# ---------------------------------------------------------------- classify

def classify(df: pd.DataFrame) -> dict:
    """Bucket every row's State Marital Status Description against the API.

    Returns the plan the UI renders and `apply_fixes` consumes. Nothing is
    changed here -- classification stays read-only so the user sees the facts
    before anything is decided.
    """
    status_col = _find_col(df, STATUS_COL)
    if not status_col:
        raise ValueError(
            "Could not find the '{}' column in the file.".format(STATUS_COL)
        )
    state_col = _find_col(df, STATE_COL)

    buckets = {b: [] for b in BUCKET_LABELS}
    punctuation = {}          # row index -> the API's spelling
    decision_rows = OrderedDict()   # (state, value) -> [row index, ...]

    if state_col:
        for idx, row in df.iterrows():
            state = _clean(row.get(state_col)).upper()
            value = _clean(row.get(status_col))

            if not state:
                buckets[NO_STATE].append(idx)
            elif sfs.is_no_sit(state):
                buckets[NO_SIT].append(idx)
            elif not sfs.has_table(state):
                buckets[NO_TABLE].append(idx)
            elif sfs.is_accepted(state, value):
                # Georgia maps "" to GA_SINGLE, so a blank there is accepted.
                buckets[ACCEPTED].append(idx)
            else:
                canonical = sfs.canonical_label(state, value) if value else None
                if canonical:
                    buckets[PUNCTUATION].append(idx)
                    punctuation[idx] = canonical
                else:
                    buckets[NEEDS_DECISION].append(idx)
                    decision_rows.setdefault((state, value), []).append(idx)

    # Biggest first, so the user's attention goes where the rows are.
    decisions = sorted(
        decision_rows,
        key=lambda k: (-len(decision_rows[k]), k[0], k[1]),
    )

    return {
        "status_col": status_col,
        "state_col": state_col,
        "buckets": buckets,
        "counts": Counter({b: len(v) for b, v in buckets.items()}),
        "punctuation": punctuation,
        "decision_rows": decision_rows,
        "decisions": decisions,
        "total_rows": len(df),
    }


# ---------------------------------------------------------------- apply

def apply_fixes(df: pd.DataFrame, plan: dict, decisions=None):
    """Write the defaults and the user's filing-status mappings into a copy.

    `decisions` maps (state, value) -> chosen label. A pair the user left empty
    is skipped: those rows are neither filled nor altered.
    """
    decisions = decisions or {}
    status_col = plan["status_col"]
    df_fixed = df.copy()

    id_col = _find_col(df, "Associate ID")
    first_col = _find_col(df, "Legal First Name")
    last_col = _find_col(df, "Legal Last Name")
    state_col = plan["state_col"]

    change_rows = []
    fill_counts = Counter()
    touched = set()

    def log(idx, row, column, old, new, reason):
        touched.add(idx)
        change_rows.append({
            "Associate ID": _clean(row.get(id_col)) if id_col else "",
            "Employee Name": " ".join(
                p for p in (
                    _clean(row.get(first_col)) if first_col else "",
                    _clean(row.get(last_col)) if last_col else "",
                ) if p
            ),
            "State": _clean(row.get(state_col)).upper() if state_col else "",
            "Column": column,
            "Old Value": old or "(blank)",
            "Filled With": new,
            "Reason": reason,
        })

    # 1. The two unconditional defaults, unchanged from before.
    for target, default in DEFAULTS.items():
        col = _find_col(df, target)
        if not col:
            raise ValueError(
                "Could not find the '{}' column in the file.".format(target)
            )
        for idx, row in df.iterrows():
            if _is_blank(row.get(col)):
                df_fixed.at[idx, col] = default
                fill_counts[target] += 1
                log(idx, row, target, "", default,
                    "Blank filled with the standard default")

    # 2. Punctuation repairs -- a correction, not a guess: the value already
    #    resolves to exactly one accepted label once punctuation is ignored.
    for idx, canonical in plan["punctuation"].items():
        row = df.loc[idx]
        old = _clean(row.get(status_col))
        df_fixed.at[idx, status_col] = canonical
        fill_counts[STATUS_COL] += 1
        state = _clean(row.get(state_col)).upper() if state_col else ""
        log(idx, row, STATUS_COL, old, canonical,
            "Punctuation corrected to Uzio's spelling ({})".format(state))

    # 3. The user's mappings.
    review_rows = []
    for (state, value) in plan["decisions"]:
        idxs = plan["decision_rows"][(state, value)]
        choice = (decisions.get((state, value)) or "").strip()
        if choice:
            for idx in idxs:
                row = df.loc[idx]
                df_fixed.at[idx, status_col] = choice
                fill_counts[STATUS_COL] += 1
                reason = (
                    "Blank filled from your mapping ({})".format(state)
                    if not value else
                    "'{}' is not accepted for {}, remapped by you".format(value, state)
                )
                log(idx, row, STATUS_COL, value, choice, reason)
        review_rows.append({
            "State": state,
            "Value In File": value or "(blank)",
            "Employees": len(idxs),
            "Mapped To": choice or "(left as is)",
            "Result": "Fixed" if choice
                      else "Still rejected by Uzio — no mapping chosen",
        })

    # Rows we deliberately did not touch, so the omission is on the record too.
    for idx in plan["buckets"][NO_STATE]:
        row = df.loc[idx]
        review_rows.append({
            "State": "(blank)",
            "Value In File": _clean(row.get(status_col)) or "(blank)",
            "Employees": 1,
            "Mapped To": "(left as is)",
            "Result": "No worked-in state on the row — cannot be checked",
        })

    changes_df = pd.DataFrame(
        change_rows,
        columns=["Associate ID", "Employee Name", "State", "Column",
                 "Old Value", "Filled With", "Reason"],
    )
    review_df = pd.DataFrame(
        review_rows,
        columns=["State", "Value In File", "Employees", "Mapped To", "Result"],
    )

    counts = plan["counts"]
    unresolved = sum(
        len(plan["decision_rows"][k]) for k in plan["decisions"]
        if not (decisions.get(k) or "").strip()
    )
    summary_df = pd.DataFrame({
        "Metric": [
            "Total rows",
            "Rows with at least one value written",
            "Dependents blanks filled",
            "Non-Resident Alien blanks filled",
            "State filing statuses written",
            "State filing statuses already accepted by Uzio",
            "State filing statuses left alone (no list for that state)",
            "State filing statuses left alone (no state income tax)",
            "Rows with no worked-in state",
            "Rows Uzio would still reject (no mapping chosen)",
        ],
        "Value": [
            len(df),
            len(touched),
            fill_counts["Dependents"],
            fill_counts["Non-Resident Alien"],
            fill_counts[STATUS_COL],
            counts[ACCEPTED],
            counts[NO_TABLE],
            counts[NO_SIT],
            counts[NO_STATE],
            unresolved,
        ],
    })

    # Stringify everything to keep long numeric strings (e.g. amounts, IDs)
    # from being emitted in exponential notation in either output.
    df_fixed_clean = df_fixed.fillna("").astype(str)
    df_fixed_clean = df_fixed_clean.replace({"nan": "", "NaN": "", "None": ""})

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        changes_df.to_excel(writer, sheet_name="Changes", index=False)
        review_df.to_excel(writer, sheet_name="Filing Status Review", index=False)
        df_fixed_clean.to_excel(writer, sheet_name="Corrected_Source", index=False)

    # Bare UTF-8 (NO BOM). Downstream APIs match the first header literally; a
    # utf-8-sig BOM smuggles U+FEFF in front of it and the column lookup silently
    # misses. Excel users should open the XLSX export instead.
    csv_bytes = df_fixed_clean.to_csv(index=False).encode("utf-8")

    return out.getvalue(), csv_bytes, summary_df, changes_df, review_df


def run_sanity(adp_file, decisions=None):
    """Read, classify and apply in one call — for non-interactive callers.

    Without `decisions` nothing is mapped, so rows the API would reject are left
    exactly as they are and reported in the Filing Status Review sheet.
    """
    df = _read_file(adp_file)
    plan = classify(df)
    return apply_fixes(df, plan, decisions)


# ---------------------------------------------------------------- UI

def _render_plan(plan: dict) -> None:
    counts = plan["counts"]
    st.subheader("What's in the file")

    if not plan["state_col"]:
        st.warning(
            "This file has no **{}** column, so state filing statuses cannot be "
            "checked — they are left exactly as they are. Dependents and "
            "Non-Resident Alien are still filled.".format(STATE_COL)
        )
        return

    breakdown = pd.DataFrame(
        [{"State filing status": BUCKET_LABELS[b], "Employees": counts[b]}
         for b in BUCKET_LABELS if counts[b]],
    )
    st.dataframe(breakdown, hide_index=True, use_container_width=True)

    if counts[NO_STATE]:
        st.info(
            "{} row(s) have no **{}**, so there is no state to check them "
            "against. They are left untouched.".format(counts[NO_STATE], STATE_COL)
        )


def _render_mapping(plan: dict) -> dict:
    """One dropdown per distinct (state, value). Returns the user's choices."""
    decisions = {}
    if not plan["decisions"]:
        return decisions

    st.subheader("Map the values Uzio does not accept")
    st.caption(
        "Uzio accepts a different set of filing statuses in each state. Nothing "
        "is pre-selected — a plausible-looking guess is worse than none here, "
        "because it invites being accepted unread. A dropdown you leave alone "
        "means those rows are skipped, not filled."
    )

    for i, (state, value) in enumerate(plan["decisions"]):
        n = len(plan["decision_rows"][(state, value)])
        shown = value or "(blank)"
        options = [LEAVE_AS_IS] + sfs.accepted_labels(state)
        choice = st.selectbox(
            "**{}** — `{}` · {} employee(s)".format(state, shown, n),
            options,
            index=0,
            key="adp_fitsit_map_{}".format(i),
        )
        if choice != LEAVE_AS_IS:
            decisions[(state, value)] = choice
    return decisions


def render_ui():
    st.title(APP_TITLE)
    st.markdown(
        """
**Purpose**: Fill blanks in the ADP FIT/SIT report with the values Uzio expects,
so the file is API-ready.

| Column | What happens |
|---|---|
| Dependents | blanks filled with `0` |
| Non-Resident Alien | blanks filled with `No` |
| State Marital Status Description | checked against what Uzio accepts **for that employee's worked-in state**; anything it would reject is yours to map |

Everything else in the file is left untouched — the API handles the rest.
"""
    )

    client_name = st.text_input("Client Name", value="Client", key="adp_fitsit_client")
    adp_file = st.file_uploader(
        "Upload ADP FIT/SIT Report (.csv / .xlsx)",
        type=["csv", "xlsx", "xls"],
        key="adp_fitsit_upload",
    )

    if not adp_file:
        for k in ("adp_fitsit_plan", "adp_fitsit_df", "adp_fitsit_out"):
            st.session_state.pop(k, None)
        return

    # A new upload invalidates everything the previous one produced.
    signature = (adp_file.name, adp_file.size)
    if st.session_state.get("adp_fitsit_sig") != signature:
        st.session_state["adp_fitsit_sig"] = signature
        for k in ("adp_fitsit_plan", "adp_fitsit_df", "adp_fitsit_out"):
            st.session_state.pop(k, None)

    if st.button("Check File", type="primary", key="adp_fitsit_run"):
        try:
            with st.spinner("Checking filing statuses..."):
                df = _read_file(adp_file)
                plan = classify(df)
        except Exception as e:
            st.error("Failed: {}".format(e))
            st.exception(e)
            return
        st.session_state["adp_fitsit_df"] = df
        st.session_state["adp_fitsit_plan"] = plan
        st.session_state.pop("adp_fitsit_out", None)

    plan = st.session_state.get("adp_fitsit_plan")
    if plan is None:
        return

    _render_plan(plan)
    decisions = _render_mapping(plan)

    build_label = ("Apply Mappings & Build File" if plan["decisions"]
                   else "Build Corrected File")
    if st.button(build_label, type="primary", key="adp_fitsit_build"):
        try:
            with st.spinner("Building..."):
                st.session_state["adp_fitsit_out"] = apply_fixes(
                    st.session_state["adp_fitsit_df"], plan, decisions
                )
        except Exception as e:
            st.error("Failed: {}".format(e))
            st.exception(e)
            return

    out = st.session_state.get("adp_fitsit_out")
    if out is None:
        return

    xlsx_bytes, csv_bytes, summary_df, changes_df, review_df = out
    st.success("Sanity check complete.")

    st.subheader("Summary")
    st.dataframe(summary_df, hide_index=True, use_container_width=True)

    if not review_df.empty:
        st.subheader("Filing Status Review")
        st.dataframe(review_df, hide_index=True, use_container_width=True)

    if changes_df.empty:
        st.info("Nothing needed changing — the file is already clean.")
    else:
        st.subheader("Changes")
        with st.container(height=400, border=True):
            st.dataframe(changes_df, hide_index=True, use_container_width=True)

    timestamp = pd.Timestamp.now().strftime("%d_%m_%Y_%H%M")
    xlsx_name = "{}_ADP_FIT_SIT_Sanity_{}.xlsx".format(client_name, timestamp)
    csv_name = "{}_ADP_FIT_SIT_Corrected_{}.csv".format(client_name, timestamp)

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            label="📊 Download Full Report (.xlsx)",
            data=xlsx_bytes,
            file_name=xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            key="adp_fitsit_dl_xlsx",
        )
    with col2:
        st.download_button(
            label="📄 Download Corrected Source (.csv)",
            data=csv_bytes,
            file_name=csv_name,
            mime="text/csv",
            key="adp_fitsit_dl_csv",
        )


if __name__ == "__main__":
    st.set_page_config(page_title=APP_TITLE, layout="centered", initial_sidebar_state="collapsed")
    render_ui()
