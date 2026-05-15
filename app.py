import streamlit as st

st.set_page_config(page_title="Unified Audit for Implementors", page_icon="⚙️", layout="wide")

# Sidebar styling
st.markdown("""
<style>
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #050e39 0%, #0a1128 100%) !important;
    }
    
    /* Target text elements but EXCLUDE Material Icons (which use spans/i with specific classes) */
    [data-testid="stSidebar"] p, 
    [data-testid="stSidebar"] h1, 
    [data-testid="stSidebar"] h2, 
    [data-testid="stSidebar"] h3,
    [data-testid="stSidebar"] label {
        color: #ffffff !important;
    }
    
    /* Ensure radio buttons inside the sidebar remain visible */
    [data-testid="stSidebar"] div[role="radiogroup"] label p {
        color: #ffffff !important;
    }
    
    /* Subtle divider */
    [data-testid="stSidebarUserContent"] hr {
        border-color: rgba(255, 255, 255, 0.1) !important;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar navigation
st.sidebar.title("Data Migration Assistant")
platform = st.sidebar.radio("Select Platform", ["ADP", "Paycom", "Universal Tools"])

if platform == "ADP":
    st.sidebar.subheader("ADP Tools")
    adp_tool = st.sidebar.radio("Select Tool", ["Census Sanity", "Census Audit", "Payment Audit", "Emergency Audit", "License Audit"], key="adp_nav")
    
    if adp_tool == "Census Sanity":
        from apps.adp.census_generator import render_census_sanity_check
        render_census_sanity_check()
    elif adp_tool == "Census Audit":
        from apps.adp.census_audit import render_ui
        render_ui()
    elif adp_tool == "Payment Audit":
        from apps.adp.payment_audit import render_ui
        render_ui()
    elif adp_tool == "Emergency Audit":
        from apps.adp.emergency_audit import render_ui
        render_ui()
    elif adp_tool == "License Audit":
        from apps.adp.license_audit import render_ui
        render_ui()

elif platform == "Paycom":
    st.sidebar.subheader("Paycom Tools")
    paycom_tool = st.sidebar.radio("Select Tool", ["Census Sanity", "Census Audit", "Payment Audit", "Emergency Audit"], key="paycom_nav")

    if paycom_tool == "Census Sanity":
        from apps.paycom.census_generator import render_census_sanity_check
        render_census_sanity_check()
    elif paycom_tool == "Census Audit":
        from apps.paycom.census_audit import render_ui
        render_ui()
    elif paycom_tool == "Payment Audit":
        from apps.paycom.payment_audit import render_ui
        render_ui()
    elif paycom_tool == "Emergency Audit":
        from apps.paycom.emergency_audit import render_ui
        render_ui()

elif platform == "Universal Tools":
    st.sidebar.subheader("Universal Tools")
    univ_tool = st.sidebar.radio("Select Tool", ["Selective Employee Extractor"], key="univ_nav")
    
    if univ_tool == "Selective Employee Extractor":
        from apps.common.employee_extractor import render_employee_extractor
        render_employee_extractor()
