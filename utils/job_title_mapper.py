"""Shared job-title mapping utility.

Maps free-form DSP job titles (from ADP / Paycom census exports) to Amazon's
30-row standard catalog. Used by the Streamlit Census Sanity tool
to emit a side-car `job_title_mapping.csv` (DSP Job Title | Amazon Job Title)
alongside the cleaned census.

Vendor fallback chains (per-row, when the primary field is blank):
  - ADP:    Job Title Description -> Department Description
  - Paycom: Position -> Business_Title -> Job_Title_Description -> Department_Desc
"""
import io
import os
import re
from pathlib import Path
import pandas as pd
import streamlit as st

CATALOG_PATH = Path(__file__).parent.parent / "templates" / "amazon_job_titles.csv"

VENDOR_FALLBACK_CHAIN = {
    "adp": [
        {"map_key": "Job Title",  "raw_col": "Job Title Description"},
        {"map_key": "Department", "raw_col": "Department Description"},
    ],
    "paycom": [
        {"map_key": "Job Title",  "raw_col": "Position"},
        {"raw_col": "Business_Title"},
        {"raw_col": "Job_Title_Description"},
        {"map_key": "Department", "raw_col": "Department_Desc"},
    ],
}

def load_amazon_catalog() -> list[str]:
    """Load standard Amazon titles for dropdowns."""
    try:
        df = pd.read_csv(CATALOG_PATH, dtype=str).fillna("")
        titles = df["Job Title"].str.strip().tolist()
        return [t for t in titles if t]
    except Exception as e:
        st.error(f"Error loading Amazon Job Title catalog: {e}")
        return []

def _norm(s) -> str:
    if s is None: return ""
    try:
        if pd.isna(s): return ""
    except: pass
    out = re.sub(r"\s+", " ", str(s)).strip()
    return "" if out.lower() == "nan" else out

def _find_column(df: pd.DataFrame, target: str):
    if not target: return None
    t = _norm(target).lower()
    for col in df.columns:
        if _norm(col).lower() == t: return col
    return None

def extract_dsp_titles(df: pd.DataFrame, vendor: str, resolved_field_map: dict | None = None) -> list[str]:
    """Return distinct, non-empty DSP titles after applying the vendor fallback chain."""
    vendor = vendor.lower()
    chain = VENDOR_FALLBACK_CHAIN.get(vendor, VENDOR_FALLBACK_CHAIN["adp"])

    cols: list[str] = []
    for step in chain:
        actual = None
        if "map_key" in step and resolved_field_map and resolved_field_map.get(step["map_key"]):
            actual = resolved_field_map[step["map_key"]]
        if not actual and "raw_col" in step:
            actual = _find_column(df, step["raw_col"])
        if actual and actual in df.columns and actual not in cols:
            cols.append(actual)

    if not cols: return []

    titles: set[str] = set()
    for _, row in df[cols].iterrows():
        for c in cols:
            v = _norm(row[c])
            if v:
                titles.add(v)
                break
    return sorted(list(titles), key=str.lower)

def render_streamlit_section(df, vendor: str, resolved_field_map=None, key_prefix: str = "jtmap"):
    """Renders the 'Map Job Titles to Amazon Catalog' UI block."""
    st.markdown("---")
    st.markdown("### 🏷️ Amazon Job Title Mapping")
    st.caption(
        "Generate a side-car CSV mapping each distinct DSP job title to "
        "Amazon's standard catalog. The cleaned census file remains unchanged."
    )

    distinct_dsp_titles = extract_dsp_titles(df, vendor, resolved_field_map)
    if not distinct_dsp_titles:
        st.info("No job titles found in this file to map.")
        return

    standard_amazon_titles = load_amazon_catalog()
    if not standard_amazon_titles:
        return

    # Auto-mapping logic (exact match)
    mapping_data = []
    for dsp_t in distinct_dsp_titles:
        matched_amz = ""
        # Check for exact match (case-insensitive)
        for amz_t in standard_amazon_titles:
            if dsp_t.strip().lower() == amz_t.strip().lower():
                matched_amz = amz_t
                break
        mapping_data.append({"DSP Job Title": dsp_t, "Amazon Job Title": matched_amz})

    df_mapping = pd.DataFrame(mapping_data)

    st.write(f"Found **{len(distinct_dsp_titles)}** unique job titles in the census file.")
    st.markdown("Please verify or select the correct **Amazon Job Title** for each DSP title below:")

    # Data Editor for manual mapping
    edited_df = st.data_editor(
        df_mapping,
        column_config={
            "DSP Job Title": st.column_config.TextColumn("DSP Job Title", disabled=True),
            "Amazon Job Title": st.column_config.SelectboxColumn(
                "Amazon Job Title",
                options=standard_amazon_titles,
                required=True,
                help="Select the standard Amazon title that best matches this DSP title."
            )
        },
        hide_index=True,
        use_container_width=True,
        key=f"{key_prefix}_editor"
    )

    # Download Button
    stamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
    csv_bytes = edited_df.to_csv(index=False).encode("utf-8")
    
    st.download_button(
        label="📥 Download Job Title Mapping (CSV)",
        data=csv_bytes,
        file_name=f"{vendor}_Job_Title_Mapping_{stamp}.csv",
        mime="text/csv",
        key=f"{key_prefix}_dl_btn",
        help="Download the mapping file with 'DSP Job Title' and 'Amazon Job Title' columns."
    )
