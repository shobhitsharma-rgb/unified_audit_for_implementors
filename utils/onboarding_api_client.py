"""Thin client for the Uzio onboarding-service (production) API.

Replicates the "Prod-Amazon Onboarding Automation" Postman collection: login to get
a JWT, then push files to the vendor-agnostic Employee Census endpoint. Scope is
intentionally limited to Employee Census for now -- the same login() can be reused
to wire the remaining endpoints (tax withholding, payment method, deductions,
contributions, workers comp, SOC code, prior payroll, job title) later.

Nothing here persists credentials to disk: login() takes them as plain arguments,
Streamlit's password widget keeps the typed value only in the browser session, and
nothing is written to logs, files, or session_state beyond the request itself.
"""
import streamlit as st
import requests

DEFAULT_PROD_HOST = "https://app.uzio.com"


class OnboardingAPIError(Exception):
    """Raised when the onboarding API rejects a login or returns an unusable response."""


def login(username: str, password: str, fein: str, vendor: str,
          prod_host: str = DEFAULT_PROD_HOST, timeout: int = 30) -> str:
    """POST {prod_host}/app/onboarding/token -> JWT string.

    vendor must match the source census's vendor: "ADP" or "PAYCOM".
    """
    url = f"{prod_host.rstrip('/')}/app/onboarding/token"
    payload = {"username": username, "password": password, "fein": fein, "vendor": vendor}
    try:
        resp = requests.post(url, json=payload, timeout=timeout)
    except requests.RequestException as e:
        raise OnboardingAPIError(f"Could not reach {url}: {e}") from e

    if not resp.ok:
        raise OnboardingAPIError(f"Login failed (HTTP {resp.status_code}): {resp.text[:500]}")

    token = resp.text.strip()
    # Some deployments wrap the token in JSON instead of returning it bare.
    if token.startswith("{"):
        try:
            data = resp.json()
            token = data.get("token") or data.get("access_token") or data.get("jwt") or ""
        except ValueError:
            token = ""

    if not token or token.count(".") != 2:
        raise OnboardingAPIError(f"Login succeeded but no JWT found in the response: {resp.text[:300]}")
    return token


def push_employee_census(jwt_token: str,
                          census_bytes: bytes, census_filename: str,
                          job_title_mapping_bytes: bytes, job_title_mapping_filename: str,
                          license_ec_bytes: bytes = None, license_ec_filename: str = None,
                          prod_host: str = DEFAULT_PROD_HOST, timeout: int = 120):
    """POST {prod_host}/app/onboarding/api/employee/census (multipart).

    job_title_mapping_file is REQUIRED by this endpoint (confirmed against the
    reference "Prod-Amazon Onboarding Automation" Postman collection).
    license_and_emergency_contact_file is optional there (disabled in the saved
    example) and stays optional here.

    Returns the raw requests.Response so the caller can inspect status/body.
    """
    url = f"{prod_host.rstrip('/')}/app/onboarding/api/employee/census"
    headers = {"Accept": "application/json", "AuthorizationHeader": jwt_token}
    files = {
        "employee_census_adv_file": (census_filename, census_bytes),
        "job_title_mapping_file": (job_title_mapping_filename, job_title_mapping_bytes),
    }
    if license_ec_bytes is not None:
        files["license_and_emergency_contact_file"] = (license_ec_filename, license_ec_bytes)
    try:
        return requests.post(url, headers=headers, files=files, timeout=timeout)
    except requests.RequestException as e:
        raise OnboardingAPIError(f"Could not reach {url}: {e}") from e


def render_push_to_uzio_section(vendor: str, data_key: str, jt_key_prefix: str, key_prefix: str):
    """Renders the 'Push Census to Uzio (Production)' section.

    vendor: "ADP" or "PAYCOM" -- sent to the login endpoint.
    data_key: session_state key holding the cleaned census bytes, set by the Census
        Sanity tool's download-button block ({"xlsx":..., "csv":..., "audit":...}).
    jt_key_prefix: the key_prefix passed to job_title_mapper.render_streamlit_section
        for this same file -- used to retrieve the mapping CSV it stashed.
    """
    st.markdown("---")
    st.markdown("### \U0001F680 Push Census to Uzio (Production)")
    st.warning(
        "This sends the corrected census straight to the **live Uzio onboarding API** "
        "and creates/updates real employee records for the employer below. "
        "Double-check the FEIN and the files before submitting."
    )

    cached = st.session_state.get(data_key, {})
    census_bytes = cached.get("csv")
    if not census_bytes:
        st.info("Run the sanity check above first — the corrected census isn't ready yet.")
        return

    jt_mapping = st.session_state.get(f"{jt_key_prefix}_job_title_mapping")
    if not jt_mapping:
        st.info("Complete the **Amazon Job Title Mapping** table above first — this API requires it.")
        return

    fein = st.text_input(
        "Employer FEIN", key=f"{key_prefix}_fein",
        help="Same FEIN used to log into Uzio for this employer.",
    )
    col_u, col_p = st.columns(2)
    with col_u:
        username = st.text_input("Uzio Username", key=f"{key_prefix}_user")
    with col_p:
        password = st.text_input(
            "Uzio Password", type="password", key=f"{key_prefix}_pass",
            help="Never stored -- used only for this request.",
        )

    include_lic = st.checkbox(
        "Also include a License & Emergency Contact file (optional)",
        key=f"{key_prefix}_inc_lic",
    )
    lic_file = None
    if include_lic:
        lic_file = st.file_uploader(
            "License & Emergency Contact file", type=["csv", "xlsx"],
            key=f"{key_prefix}_lic_upload",
        )

    confirm = st.checkbox(
        f"I confirm this pushes to **PRODUCTION** for FEIN `{fein or '____'}` and I've reviewed the files.",
        key=f"{key_prefix}_confirm",
    )

    if st.button("\U0001F680 Push to Uzio", key=f"{key_prefix}_btn", disabled=not confirm):
        if not (fein.strip() and username.strip() and password):
            st.error("FEIN, Username, and Password are all required.")
            return
        try:
            with st.spinner("Logging in to Uzio..."):
                token = login(username.strip(), password, fein.strip(), vendor)
            with st.spinner("Uploading census..."):
                lic_bytes = lic_file.getvalue() if lic_file else None
                lic_name = lic_file.name if lic_file else None
                resp = push_employee_census(
                    token,
                    census_bytes, f"{vendor}_Census.csv",
                    jt_mapping["csv"], jt_mapping["filename"],
                    lic_bytes, lic_name,
                )
            if resp.ok:
                st.success(f"✅ Census pushed successfully (HTTP {resp.status_code}).")
                if resp.text:
                    with st.expander("Response details"):
                        st.code(resp.text[:3000])
            else:
                st.error(f"❌ API returned HTTP {resp.status_code}")
                with st.expander("Response details"):
                    st.code(resp.text[:3000])
        except OnboardingAPIError as e:
            st.error(f"❌ {e}")
