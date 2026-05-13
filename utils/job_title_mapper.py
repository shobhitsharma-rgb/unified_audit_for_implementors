"""Shared job-title mapping utility.

Maps free-form DSP job titles (from ADP / Paycom census exports) to Amazon's
30-row standard catalog using Claude. Used by the Streamlit Census Sanity tool
to emit a side-car `job_title_mapping.csv` (DSP Job Title | Amazon Job Title)
alongside the cleaned census.

Vendor fallback chains (per-row, when the primary field is blank):
  - ADP:    Job Title Description -> Department Description
  - Paycom: Position -> Business_Title -> Job_Title_Description -> Department_Desc
"""
from __future__ import annotations

import io
import json
import os
import re
from pathlib import Path

import pandas as pd

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

CLAUDE_MODEL = "claude-3-haiku-20240307"


def load_amazon_catalog() -> pd.DataFrame:
    df = pd.read_csv(CATALOG_PATH, dtype=str).fillna("")
    return df[df["Job Title"].str.strip() != ""].reset_index(drop=True)


def _norm(s) -> str:
    if s is None:
        return ""
    try:
        if pd.isna(s):
            return ""
    except (TypeError, ValueError):
        pass
    out = re.sub(r"\s+", " ", str(s)).strip()
    return "" if out.lower() == "nan" else out


def _find_column(df: pd.DataFrame, target: str):
    if not target:
        return None
    t = _norm(target).lower()
    for col in df.columns:
        if _norm(col).lower() == t:
            return col
    return None


def extract_dsp_titles(
    df: pd.DataFrame,
    vendor: str,
    resolved_field_map: dict | None = None,
) -> list[str]:
    """Return distinct, non-empty DSP titles after applying the vendor fallback chain.

    `resolved_field_map` is the dict produced by `preprocess_{adp,paycom}_file`
    in the Streamlit codebase: canonical key -> actual source column name. When
    omitted (MCP path), columns are looked up by their canonical names directly.
    """
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

    if not cols:
        return []

    titles: set[str] = set()
    for _, row in df[cols].iterrows():
        for c in cols:
            v = _norm(row[c])
            if v:
                titles.add(v)
                break

    return sorted(titles, key=str.lower)


def _build_prompt(distinct_titles: list[str], catalog: pd.DataFrame) -> tuple[str, str]:
    catalog_lines = [
        f"- {row['Job Title']} (Category: {row['Job Category']}, Code: {row['Job Code']})"
        for _, row in catalog.iterrows()
    ]
    system = (
        "You map free-form employee job titles from a Delivery Service Partner (DSP) census file "
        "to Amazon's 30-row standard DSP job-title catalog.\n\n"
        "STANDARD AMAZON CATALOG (the Amazon Job Title MUST be exactly one of these strings):\n"
        + "\n".join(catalog_lines)
        + "\n\nMapping rules:\n"
        "- Pick the single closest semantic match. \"Sr. Operations Mgr\" -> Operations Manager. "
        "\"Delivery Driver\" / \"DA\" / \"Associate\" -> Driver. \"Helper\" / \"Jumper\" -> Helper. "
        "\"Walker\" / \"Foot delivery\" -> Walker. Step-van drivers -> Driver-Step Van.\n"
        "- Department-only fallbacks (e.g. \"Operations\", \"Warehouse\", \"Dispatch\") map to the closest catalog title.\n"
        "- If the input is clearly outside DSP scope (CEO, Software Engineer, etc.), use \"Non-DSP Related\".\n"
        "- Return STRICTLY a JSON array. No prose, no markdown fences."
    )
    user = (
        "Map each DSP title below. Return JSON: "
        "[{\"dsp_title\": \"<input>\", \"amazon_title\": \"<one catalog value>\"}, ...]\n\n"
        "DSP titles:\n" + "\n".join(f"- {t}" for t in distinct_titles)
    )
    return system, user


def map_titles_with_claude(
    distinct_titles: list[str],
    catalog: pd.DataFrame,
    api_key: str,
) -> pd.DataFrame:
    """One batched call. Returns DataFrame[DSP Job Title, Amazon Job Title]."""
    if not distinct_titles:
        return pd.DataFrame(columns=["DSP Job Title", "Amazon Job Title"])

    from anthropic import Anthropic

    client = Anthropic(api_key=api_key)
    system, user = _build_prompt(distinct_titles, catalog)
    resp = client.messages.create(
        model=CLAUDE_MODEL,
        max_tokens=4096,
        system=[
            {"type": "text", "text": system, "cache_control": {"type": "ephemeral"}}
        ],
        messages=[{"role": "user", "content": user}],
    )
    text = "".join(b.text for b in resp.content if getattr(b, "type", "") == "text").strip()
    text = re.sub(r"^```(?:json)?|```$", "", text, flags=re.MULTILINE).strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"Claude did not return valid JSON: {e}\n--- response ---\n{text[:1000]}")

    valid_amazon = set(catalog["Job Title"].tolist())
    rows = []
    seen_inputs = set()
    for item in data:
        dsp = _norm(item.get("dsp_title"))
        amz = _norm(item.get("amazon_title"))
        if not dsp or dsp in seen_inputs:
            continue
        if amz not in valid_amazon:
            amz = "Non-DSP Related" if "Non-DSP Related" in valid_amazon else amz
        rows.append({"DSP Job Title": dsp, "Amazon Job Title": amz})
        seen_inputs.add(dsp)

    for t in distinct_titles:
        if t not in seen_inputs:
            rows.append({"DSP Job Title": t, "Amazon Job Title": ""})

    return pd.DataFrame(rows, columns=["DSP Job Title", "Amazon Job Title"])


def to_csv_bytes(mapping: pd.DataFrame) -> bytes:
    buf = io.StringIO()
    mapping.to_csv(buf, index=False)
    return buf.getvalue().encode("utf-8")


def get_anthropic_api_key() -> str | None:
    try:
        import streamlit as st
        if "ANTHROPIC_API_KEY" in st.secrets:
            return st.secrets["ANTHROPIC_API_KEY"]
    except Exception:
        pass
    return os.environ.get("ANTHROPIC_API_KEY")


def render_streamlit_section(df, vendor: str, resolved_field_map=None, key_prefix: str = "jtmap"):
    """Renders the 'Map Job Titles to Amazon Catalog' UI block. Used by both
    ADP and Paycom census sanity check screens."""
    import hashlib
    import streamlit as st

    st.markdown("---")
    st.markdown("### 🏷️ Amazon Job Title Mapping")
    st.caption(
        "Generate a side-car CSV mapping each distinct DSP job title to "
        "Amazon's standard catalog. The cleaned census file is unchanged."
    )

    distinct = extract_dsp_titles(df, vendor, resolved_field_map)
    if not distinct:
        st.info("No job-title or fallback (department) values found in this file — nothing to map.")
        return

    st.write(f"Found **{len(distinct)} distinct DSP titles** (after Job Title → Department fallback).")

    # Cache key tied to the actual input titles — a new file invalidates the cache.
    sig = hashlib.md5("|".join(distinct).encode()).hexdigest()[:12]
    cache_key = f"{key_prefix}_jt_mapping_{sig}"

    if cache_key not in st.session_state:
        if not st.button("Generate Job Title Mapping", key=f"{key_prefix}_jt_btn"):
            return

        api_key = get_anthropic_api_key()
        if not api_key:
            st.error(
                "ANTHROPIC_API_KEY not found. Set it in `.streamlit/secrets.toml` or "
                "the `ANTHROPIC_API_KEY` environment variable, then click again."
            )
            return

        with st.spinner(f"Mapping {len(distinct)} titles via Claude..."):
            try:
                catalog = load_amazon_catalog()
                mapping_df = map_titles_with_claude(distinct, catalog, api_key)
            except Exception as e:
                st.error(f"Mapping failed: {e}")
                return

        st.session_state[cache_key] = mapping_df

    mapping_df = st.session_state[cache_key]

    st.success(f"Mapped {len(mapping_df)} titles.")
    st.dataframe(mapping_df, hide_index=True, use_container_width=True)
    stamp = pd.Timestamp.now().strftime('%Y%m%d_%H%M')
    st.download_button(
        label="📥 Download Job Title Mapping (CSV)",
        data=to_csv_bytes(mapping_df),
        file_name=f"{vendor}_job_title_mapping_{stamp}.csv",
        mime="text/csv",
        key=f"{key_prefix}_jt_dl",
    )
    if st.button("Regenerate", key=f"{key_prefix}_jt_regen"):
        del st.session_state[cache_key]
        st.rerun()
