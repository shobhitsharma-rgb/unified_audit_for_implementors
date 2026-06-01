import io
import pandas as pd
import streamlit as st

APP_TITLE = "ADP FIT/SIT Sanity Check"

# =========================================================
# ADP FIT/SIT Sanity Check
# - Input: single ADP FIT/SIT export (.csv / .xlsx)
# - Fills blanks in three columns with hardcoded defaults:
#     1) Dependents                          -> 0
#     2) Non-Resident Alien                  -> No
#     3) State Marital Status Description    -> Single
# - Everything else is handled downstream by the API.
# =========================================================

DEFAULTS = {
    "Dependents": "0",
    "Non-Resident Alien": "No",
    "State Marital Status Description": "Single",
}


def _is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    return s == "" or s.lower() == "nan"


def _read_file(uploaded) -> pd.DataFrame:
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


def run_sanity(adp_file):
    df = _read_file(adp_file)

    # Resolve column names (defensive — should be exact)
    resolved: dict[str, str] = {}
    missing: list[str] = []
    for target in DEFAULTS:
        col = _find_col(df, target)
        if col:
            resolved[target] = col
        else:
            missing.append(target)

    if missing:
        raise ValueError(
            "Could not find required column(s) in the file: " + ", ".join(missing)
        )

    df_fixed = df.copy()

    # Pick out the ID + name columns for the change log (best-effort)
    id_col = _find_col(df, "Associate ID")
    first_col = _find_col(df, "Legal First Name")
    last_col = _find_col(df, "Legal Last Name")

    change_rows: list[dict] = []
    fill_counts: dict[str, int] = {t: 0 for t in DEFAULTS}

    for idx, row in df.iterrows():
        for target, default in DEFAULTS.items():
            col = resolved[target]
            if _is_blank(row.get(col)):
                df_fixed.at[idx, col] = default
                fill_counts[target] += 1

                emp_id = str(row.get(id_col, "")).strip() if id_col else ""
                fname = str(row.get(first_col, "")).strip() if first_col else ""
                lname = str(row.get(last_col, "")).strip() if last_col else ""
                emp_name = f"{fname} {lname}".strip()

                change_rows.append({
                    "Associate ID": emp_id,
                    "Employee Name": emp_name,
                    "Column": target,
                    "Filled With": default,
                })

    changes_df = pd.DataFrame(
        change_rows,
        columns=["Associate ID", "Employee Name", "Column", "Filled With"],
    )

    summary_df = pd.DataFrame({
        "Metric": [
            "Total rows",
            "Rows with at least one blank filled",
            "Dependents blanks filled",
            "Non-Resident Alien blanks filled",
            "State Marital Status Description blanks filled",
            "Total blanks filled",
        ],
        "Value": [
            len(df),
            changes_df["Associate ID"].nunique() if not changes_df.empty else 0,
            fill_counts["Dependents"],
            fill_counts["Non-Resident Alien"],
            fill_counts["State Marital Status Description"],
            sum(fill_counts.values()),
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
        df_fixed_clean.to_excel(writer, sheet_name="Corrected_Source", index=False)

    # Bare UTF-8 (NO BOM). Downstream APIs match the first header literally; a
    # utf-8-sig BOM smuggles U+FEFF in front of it and the column lookup silently
    # misses. Excel users should open the XLSX export instead.
    csv_bytes = df_fixed_clean.to_csv(index=False).encode("utf-8")

    return out.getvalue(), csv_bytes, summary_df, changes_df


def render_ui():
    st.title(APP_TITLE)
    st.markdown(
        """
**Purpose**: Fill blanks in the ADP FIT/SIT report with the defaults Uzio
expects, so the file is API-ready.

**Defaults applied**:
| Column | Default |
|---|---|
| Dependents | `0` |
| Non-Resident Alien | `No` |
| State Marital Status Description | `Single` |

Everything else in the file is left untouched — downstream API handles the rest.
"""
    )

    client_name = st.text_input("Client Name", value="Client", key="adp_fitsit_client")
    adp_file = st.file_uploader(
        "Upload ADP FIT/SIT Report (.csv / .xlsx)",
        type=["csv", "xlsx", "xls"],
        key="adp_fitsit_upload",
    )

    if not adp_file:
        return

    if st.button("Run Sanity Check", type="primary", key="adp_fitsit_run"):
        try:
            with st.spinner("Filling blanks..."):
                xlsx_bytes, csv_bytes, summary_df, changes_df = run_sanity(adp_file)
        except Exception as e:
            st.error(f"Failed: {e}")
            st.exception(e)
            return

        st.success("Sanity check complete.")

        st.subheader("Summary")
        st.dataframe(summary_df, hide_index=True, use_container_width=True)

        if changes_df.empty:
            st.info("No blanks found in the three target columns — file is already clean.")
        else:
            st.subheader("Changes")
            st.dataframe(changes_df, hide_index=True, use_container_width=True)

        timestamp = pd.Timestamp.now().strftime("%d_%m_%Y_%H%M")
        xlsx_name = f"{client_name}_ADP_FIT_SIT_Sanity_{timestamp}.xlsx"
        csv_name = f"{client_name}_ADP_FIT_SIT_Corrected_{timestamp}.csv"

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
