"""Paycom Payment Method Sanity Check.

Paycom keeps the bank details inside the census itself, which is why the onboarding
API's payment-method endpoint wants `employee_census_adv_file` for Paycom where ADP
sends a separate direct deposit export. So this tool takes the Paycom census and
checks the payment columns in it.

Every rule below is the rule the API applies, read off
`validator/PaymentMethodValidator.java` (Paycom's own validator class is an empty
stub, so it inherits every default method there):

  * `Has_D._Deposit` is the gate. False means no direct deposit and nothing else is
    required — the API records the employee as CHECK.
  * The net account is the remainder account: `Net_Acct_Code`, `Net_Rout_Code` and
    `Net_Type_Code` are all required when direct deposit is on.
  * Up to four distribution accounts, `Dist_1..4`. A group counts as "present" when
    any of its type / account / routing is filled, and then ALL of deposit method,
    type, account and routing are required.
  * Routing numbers are padded to nine digits with leading zeros and must pass the
    ABA checksum (weights 3,7,1 repeating).
  * Account numbers are stripped to digits and must not be all zeros.
  * Account types must be Savings, Checking, rapid! PayCard — or Vault, which the
    API lets through for Paycom specifically.
  * Percentage mode is decided by `Dist_1_Amount` containing a "%". In that mode the
    distribution amounts are totalled and may not exceed 100.

Two of those the API does to the data itself, so the tool does them too and records
them: padding a short routing number, and stripping non-digits out of an account
number. Everything else is reported, never guessed at — a wrong bank account is
worse than a blocked upload.
"""
import io
import re

import pandas as pd
import streamlit as st

from utils.file_io import read_report

APP_TITLE = "Paycom Payment Method Sanity Check"

ID_COL = "Employee_Code"
DD_COL = "Has_D._Deposit"
NET_ACCT, NET_ROUT, NET_TYPE = "Net_Acct_Code", "Net_Rout_Code", "Net_Type_Code"
DIST_RANGE = (1, 2, 3, 4)
# The API accepts these four; "Vault" is let through for Paycom specifically.
ACCOUNT_TYPES = {"savings", "checking", "rapid! paycard", "vault"}
MAX_PERCENTAGE = 100


def dist_cols(n: int) -> dict:
    return {"method": f"Dist_{n}_Deposit_Method", "type": f"Dist_{n}_Type_Code",
            "acct": f"Dist_{n}_Acct_Code", "rout": f"Dist_{n}_Rout_Code",
            "amount": f"Dist_{n}_Amount"}


def _text(value) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    text = str(value).strip()
    return "" if text.lower() in ("nan", "none") else text


def is_direct_deposit(value) -> bool:
    """`Has_D._Deposit` as the API reads it — anything but a true means no."""
    return _text(value).strip().lower() in ("true", "yes", "y", "1")


def pad_routing(value) -> str:
    """Nine digits, left-padded with zeros — exactly what the API stores."""
    text = _text(value)
    if not re.search(r"\d", text):
        return text
    digits = re.sub(r"\D", "", text)
    return digits.zfill(9) if len(digits) <= 9 else digits


def routing_is_valid(routing: str) -> bool:
    """ABA checksum: weights 3,7,1 repeating, total must divide by 10."""
    if not re.fullmatch(r"\d{9}", routing or ""):
        return False
    weights = (3, 7, 1, 3, 7, 1, 3, 7, 1)
    total = sum(int(d) * w for d, w in zip(routing, weights))
    return total % 10 == 0


def clean_account(value) -> str:
    return re.sub(r"\D", "", _text(value))


def account_is_valid(account: str) -> bool:
    """Digits, and not all zeros — the API's ^\\d*[1-9]+\\d*$."""
    return bool(re.fullmatch(r"\d*[1-9]+\d*", account or ""))


def percentage_total(row) -> tuple:
    """(is_percent_mode, total). Dollar amounts are skipped, as the API skips them."""
    first = _text(row.get(dist_cols(1)["amount"]))
    if "%" not in first:
        return False, 0.0
    total = 0.0
    for n in DIST_RANGE:
        amount = _text(row.get(dist_cols(n)["amount"]))
        if not amount or "$" in amount:
            continue
        try:
            total += float(amount.replace("%", "").strip())
        except ValueError:
            continue
    return True, total


def run_sanity(census_file):
    """(corrected dataframe, issues dataframe). The frame is a copy; nothing in
    place is touched, so the uploaded file is never the thing being edited."""
    df = read_report(census_file, dtype=str)
    df.columns = [str(c).strip() for c in df.columns]
    missing = [c for c in (ID_COL, DD_COL, NET_ACCT, NET_ROUT) if c not in df.columns]
    if missing:
        raise ValueError(
            "This does not look like a Paycom census — missing " + ", ".join(missing)
            + ". The payment columns live in the census export, not a separate file.")

    fixed = df.copy()
    issues = []

    def log(emp, rule, severity, issue, action):
        issues.append({"Employee ID": emp, "Rule": rule, "Severity": severity,
                       "Issue": issue, "Action": action})

    for idx, row in df.iterrows():
        emp = _text(row.get(ID_COL)) or f"(row {idx + 2})"

        if not is_direct_deposit(row.get(DD_COL)):
            flag = _text(row.get(DD_COL))
            if any(_text(row.get(dist_cols(n)["acct"])) for n in DIST_RANGE) or \
                    _text(row.get(NET_ACCT)):
                log(emp, "Direct deposit flag", "Review",
                    f"Has_D._Deposit is '{flag or '(blank)'}' but bank details are filled in.",
                    "The API will set this employee to CHECK and ignore the accounts. "
                    "Set the flag to TRUE if they should be paid by direct deposit.")
            continue

        # --- the net (remainder) account -------------------------------------
        for col, label in ((NET_ACCT, "Net account"), (NET_ROUT, "Net routing"),
                           (NET_TYPE, "Net account type")):
            if col in df.columns and not _text(row.get(col)):
                log(emp, "Net account", "Error",
                    f"{label} ({col}) is blank but direct deposit is on.",
                    "The API rejects this employee. Fill it in from Paycom.")

        net_acct_raw = _text(row.get(NET_ACCT))
        if net_acct_raw:
            cleaned = clean_account(net_acct_raw)
            if cleaned != net_acct_raw:
                fixed.at[idx, NET_ACCT] = cleaned
                log(emp, "Net account", "Fixed",
                    f"Net account '{net_acct_raw}' had characters that are not digits.",
                    f"Stripped to '{cleaned}' — the API does the same before storing it.")
            if not account_is_valid(cleaned):
                log(emp, "Net account", "Error",
                    f"Net account '{net_acct_raw}' is not a usable account number.",
                    "The API rejects an account that is empty or all zeros.")

        net_rout_raw = _text(row.get(NET_ROUT))
        if net_rout_raw:
            padded = pad_routing(net_rout_raw)
            if padded != net_rout_raw:
                fixed.at[idx, NET_ROUT] = padded
                log(emp, "Net routing", "Fixed",
                    f"Net routing '{net_rout_raw}' was not nine digits.",
                    f"Padded to '{padded}' — the API pads it the same way.")
            if not routing_is_valid(padded):
                log(emp, "Net routing", "Error",
                    f"Net routing '{padded}' fails the ABA checksum.",
                    "The API rejects this employee. Check the number against the bank.")

        net_type = _text(row.get(NET_TYPE))
        if net_type and net_type.lower() not in ACCOUNT_TYPES:
            log(emp, "Net account type", "Error",
                f"Net account type '{net_type}' is not one Uzio accepts.",
                "Use Savings, Checking, rapid! PayCard or Vault.")

        # --- the four distribution accounts ----------------------------------
        for n in DIST_RANGE:
            cols = dist_cols(n)
            present = [c for c in (cols["type"], cols["acct"], cols["rout"])
                       if c in df.columns and _text(row.get(c))]
            if not present:
                continue                      # the group is simply not used

            for key, label in (("method", "deposit method"), ("type", "account type"),
                               ("acct", "account number"), ("rout", "routing number")):
                col = cols[key]
                if col not in df.columns or not _text(row.get(col)):
                    log(emp, f"Distribution {n}", "Error",
                        f"Account {n} is partly filled in but its {label} ({col}) is blank.",
                        "The API requires all four together. Complete it or clear the "
                        "whole account.")

            acct_raw = _text(row.get(cols["acct"]))
            if acct_raw:
                cleaned = clean_account(acct_raw)
                if cleaned != acct_raw:
                    fixed.at[idx, cols["acct"]] = cleaned
                    log(emp, f"Distribution {n}", "Fixed",
                        f"Account {n} number '{acct_raw}' had characters that are not digits.",
                        f"Stripped to '{cleaned}' — the API does the same.")
                if not account_is_valid(cleaned):
                    log(emp, f"Distribution {n}", "Error",
                        f"Account {n} number '{acct_raw}' is not usable.",
                        "The API rejects an account that is empty or all zeros.")

            rout_raw = _text(row.get(cols["rout"]))
            if rout_raw:
                padded = pad_routing(rout_raw)
                if padded != rout_raw:
                    fixed.at[idx, cols["rout"]] = padded
                    log(emp, f"Distribution {n}", "Fixed",
                        f"Account {n} routing '{rout_raw}' was not nine digits.",
                        f"Padded to '{padded}'.")
                if not routing_is_valid(padded):
                    log(emp, f"Distribution {n}", "Error",
                        f"Account {n} routing '{padded}' fails the ABA checksum.",
                        "Check the number against the bank.")

            type_raw = _text(row.get(cols["type"]))
            if type_raw and type_raw.lower() not in ACCOUNT_TYPES:
                log(emp, f"Distribution {n}", "Error",
                    f"Account {n} type '{type_raw}' is not one Uzio accepts.",
                    "Use Savings, Checking, rapid! PayCard or Vault.")

        # --- percentage split ------------------------------------------------
        is_percent, total = percentage_total(row)
        if is_percent and total > MAX_PERCENTAGE:
            log(emp, "Percentage split", "Error",
                f"The distribution percentages add up to {total:g}%.",
                "The API rejects anything over 100%. Correct the split in Paycom.")

    columns = ["Employee ID", "Rule", "Severity", "Issue", "Action"]
    return fixed, pd.DataFrame(issues, columns=columns)


def _report_bytes(fixed: pd.DataFrame, issues: pd.DataFrame) -> bytes:
    book = io.BytesIO()
    with pd.ExcelWriter(book, engine="xlsxwriter") as writer:
        issues.to_excel(writer, sheet_name="Issues", index=False)
        errors = issues[issues["Severity"] == "Error"] if not issues.empty else issues
        if not errors.empty:
            errors.to_excel(writer, sheet_name="Blocking", index=False)
        fixed.to_excel(writer, sheet_name="Corrected Census", index=False)
    return book.getvalue()


def render_ui():
    st.title(f"💳 {APP_TITLE}")
    st.markdown(
        "Paycom keeps bank details **inside the census**, so upload the same census "
        "export here. This checks the payment columns against the rules the Uzio "
        "onboarding API applies, and can push them straight to production."
    )
    with st.expander("What it checks"):
        st.markdown(
            "- **`Has_D._Deposit`** decides everything. If it is not true the API pays "
            "the employee by cheque and ignores any bank details in the row — so a row "
            "with accounts but the flag off is flagged for you.\n"
            "- **Net account** (`Net_Acct_Code`, `Net_Rout_Code`, `Net_Type_Code`) is "
            "required when direct deposit is on. That is the remainder account.\n"
            "- **Up to four distribution accounts.** A group counts as used as soon as "
            "any of its type / account / routing is filled, and then all four of its "
            "fields are required.\n"
            "- **Routing numbers** are padded to nine digits and must pass the ABA "
            "checksum. **Account numbers** must be digits and not all zeros. **Account "
            "types** must be Savings, Checking, rapid! PayCard or Vault.\n"
            "- **Percentages** (when `Dist_1_Amount` carries a `%`) may not add up to "
            "more than 100.\n\n"
            "Two things are corrected for you, because the API does exactly the same "
            "before storing: padding a short routing number, and stripping non-digits "
            "out of an account number. Everything else is reported — a wrong bank "
            "account is worse than a blocked upload.")

    client_name = st.text_input("Client name (for the file names)", "Client",
                                key="pc_pm_client")
    census_file = st.file_uploader("Upload the Paycom census export",
                                   type=["csv", "xlsx", "xls"], key="pc_pm_upload")
    if not census_file:
        st.info("Upload the Paycom census to begin.")
        return

    # Keep the parsed result against the uploaded file, so typing in the Push
    # section below does not re-parse the whole census on every keystroke.
    results_key = "pc_pm_sanity_results"
    signature = (census_file.name, getattr(census_file, "size", None))
    cached = st.session_state.get(results_key)
    if not cached or cached.get("signature") != signature:
        try:
            fixed, issues = run_sanity(census_file)
        except ValueError as e:
            st.error(str(e))
            return
        except Exception as e:                  # a broken export should say so plainly
            st.error(f"Could not read that file: {e}")
            return
        st.session_state[results_key] = {"signature": signature, "fixed": fixed,
                                         "issues": issues}
    cached = st.session_state[results_key]
    fixed, issues = cached["fixed"], cached["issues"]

    errors = issues[issues["Severity"] == "Error"] if not issues.empty else issues
    fixes = issues[issues["Severity"] == "Fixed"] if not issues.empty else issues
    reviews = issues[issues["Severity"] == "Review"] if not issues.empty else issues

    col1, col2, col3 = st.columns(3)
    col1.metric("Employees", len(fixed))
    col2.metric("Blocking errors", int(errors["Employee ID"].nunique()) if not errors.empty else 0)
    col3.metric("Auto-corrected", int(fixes["Employee ID"].nunique()) if not fixes.empty else 0)

    if issues.empty:
        st.success("✅ Every payment row passes the rules the API applies.")
    else:
        if not errors.empty:
            st.error(f"**{errors['Employee ID'].nunique()} employee(s) will be rejected "
                     "by the API.** Fix these in Paycom and export again.")
            st.dataframe(errors, hide_index=True, use_container_width=True)
        if not reviews.empty:
            st.warning("**Worth a look before pushing.**")
            st.dataframe(reviews, hide_index=True, use_container_width=True)
        if not fixes.empty:
            st.info("**Corrected for you** — the same corrections the API makes itself.")
            st.dataframe(fixes, hide_index=True, use_container_width=True)

    timestamp = pd.Timestamp.now().strftime("%d_%m_%Y_%H%M")
    csv_name = f"{client_name}_Paycom_Payment_Method_Corrected_{timestamp}.csv"
    # Plain UTF-8, no BOM: the onboarding API matches the first header literally.
    csv_bytes = fixed.to_csv(index=False).encode("utf-8")

    st.session_state["pc_pm_cached_files"] = {"csv": csv_bytes, "filename": csv_name}

    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button("📊 Download Full Report (.xlsx)", _report_bytes(fixed, issues),
                           file_name=f"{client_name}_Paycom_Payment_Method_Sanity_{timestamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", key="pc_pm_dl_xlsx")
    with col2:
        st.download_button("📄 Download Corrected Census (.csv)", csv_bytes,
                           file_name=csv_name, mime="text/csv", key="pc_pm_dl_csv")
    with col3:
        st.download_button("📝 Download Change Log (.csv)",
                           issues.to_csv(index=False).encode("utf-8"),
                           file_name=f"{client_name}_Paycom_Payment_Method_ChangeLog_{timestamp}.csv",
                           mime="text/csv", key="pc_pm_dl_changelog")

    # --- Push to Uzio (Production) Section ---
    from utils.onboarding_api_client import render_push_payment_method_section
    render_push_payment_method_section(vendor="PAYCOM", data_key="pc_pm_cached_files",
                                       key_prefix="pc_pm_push2uzio")
