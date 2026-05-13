import streamlit as st

st.set_page_config(page_title="Unified Audit for Implementors", page_icon="⚙️", layout="wide")

# Sidebar navigation
st.sidebar.title("Navigation")
platform = st.sidebar.radio("Select Platform", ["ADP", "Paycom"])

if platform == "ADP":
    st.sidebar.subheader("ADP Tools")
    adp_tool = st.sidebar.radio("Select Tool", ["Census Sanity", "Census Audit"], key="adp_nav")
    
    if adp_tool == "Census Sanity":
        from apps.adp.census_generator import render_census_sanity_check
        render_census_sanity_check()
    elif adp_tool == "Census Audit":
        from apps.adp.census_audit import render_ui
        render_ui()

elif platform == "Paycom":
    st.sidebar.subheader("Paycom Tools")
    paycom_tool = st.sidebar.radio("Select Tool", ["Census Sanity", "Census Audit"], key="paycom_nav")
    
    if paycom_tool == "Census Sanity":
        from apps.paycom.census_generator import render_census_sanity_check
        render_census_sanity_check()
    elif paycom_tool == "Census Audit":
        from apps.paycom.census_audit import render_ui
        render_ui()
