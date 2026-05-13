import streamlit as st

st.set_page_config(page_title="Unified Audit for Implementors", page_icon="⚙️", layout="wide")

# Sidebar styling
st.markdown("""
<style>
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #050e39 0%, #0a1128 100%) !important;
    }
    [data-testid="stSidebar"] p, 
    [data-testid="stSidebar"] span, 
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3 {
        color: #ffffff !important;
        font-family: 'Inter', sans-serif !important;
    }
    [data-testid="stSidebar"] div[role="radiogroup"] label {
        color: #ffffff !important;
    }
    [data-testid="stSidebarUserContent"] hr {
        border-color: rgba(255, 255, 255, 0.1) !important;
    }
</style>
""", unsafe_allow_html=True)

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
