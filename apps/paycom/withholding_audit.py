import streamlit as st
import pandas as pd
import io
import re
from datetime import date

# Inlined config — formerly loaded from three on-disk files
# (Mapping.xlsx / key_mapping.yml / filing status_code.txt) at runtime.
# All three are now Python literals in withholding_data.py so the tool
# works on a fresh clone with zero external file dependencies.
# To update: edit scratch/generate_paycom_data_module.py source files and
# re-run that script, then commit the regenerated withholding_data.py.
from apps.paycom.withholding_data import (
    MAPPING_ROWS,
    FILING_STATUS_MAP,
    LABELS_BY_STATE,
)

# =========================================================
# Paycom to UZIO Federal/State Withholding Audit Tool (FIT/SIT)
# =========================================================

APP_TITLE = "Paycom to UZIO Withholding Audit Tool (FIT/SIT)"

ACTIVE_STATUSES = {"active", "on leave"}

DETAIL_COLUMNS = [
    "Employee ID",
    "Paycom Status",
    "Paycom State",
    "Paycom First Name",
    "Paycom Last Name",
    "UZIO First Name",
    "UZIO Last Name",
    "Field Label",
    "Paycom Column",
    "Paycom Value",
    "UZIO Field Key",
    "UZIO Stored Value",
    "Paycom Normalized",
    "UZIO Normalized / UI",
    "Rule Applied"
]

NORMALIZATION_NOTES = [
    "1. Filing Status stored in UZIO as DB value (e.g. FEDERAL_SINGLE) is mapped to UI label (e.g. Single). Match is substring and punct-insensitive.",
    "2. Boolean: Yes/Y/1/True => True; No/N/0/False => False.",
    "3. Amounts: UZIO stores in cents (divided by 100). Paycom stores in dollars.",
    "4. Blank handling: Numerics blank=0, Booleans blank=unknown.",
    "5. Fields compared strictly to what is available in Mapping file.",
]

def _norm_col(c):
    if c is None:
        return ""
    return str(c).strip().replace("\n", " ").strip()

def _pick_first(cols, candidates):
    lower_map = {c.lower(): c for c in cols}
    for cand in candidates:
        if cand.lower() in lower_map:
            return lower_map[cand.lower()]
    return None

def _autodetect_paycom_cols(df):
    """Find the canonical columns we use to drive the audit.

    State priority is Work Location FIRST. Per the user's note: SIT records
    in UZIO are stored against the work-location state, not the home state.
    An employee living in MI but working in OH has their OH SIT_* values in
    UZIO; matching against MI would always show a false mismatch.

    Falls back to other state-ish columns only when the Paycom file doesn't
    have a Work Location State column populated for this client.
    """
    cols = list(df.columns)
    emp_id = _pick_first(cols, ["Employee_Code", "Employee Code", "Employee ID", "Employee_ID", "Emp_ID", "employee_id", "EE Code"]) or cols[0]
    status = _pick_first(cols, ["Employee_Status", "Status", "Employee Status"]) or cols[0]
    state = _pick_first(cols, [
        # 1. Work-location state — primary match key for SIT.
        "Work_Location_State", "Work Location State",
        "Works-in_State", "Works in State", "Work_State", "Work State",
        # 2. SUI state — usually matches work location for most clients.
        "SUI_State",
        # 3. Home / residence state — only if no work-location column exists.
        "Lives-in_State", "State", "Home_State", "Home State",
        "Primary_State/Province", "State_Abbreviation", "Paycom State",
    ])
    first_name = _pick_first(cols, [
        "Legal_Firstname", "Legal First Name", "Legal_First_Name",
        "First_Name", "First Name", "Employee_First_Name", "FirstName", "EE Name",
    ])
    last_name = _pick_first(cols, [
        "Legal_Lastname", "Legal Last Name", "Legal_Last_Name",
        "Last_Name", "Last Name", "Employee_Last_Name", "LastName",
    ])
    return emp_id, status, state, first_name, last_name

def _load_mapping_df() -> pd.DataFrame:
    """Return the inlined Paycom -> UZIO field mapping as a DataFrame.

    Columns: 'Uzio Field Key', 'PayCom Column', 'Comments'.
    Previously read from Mapping.xlsx; now sourced from
    apps/paycom/withholding_data.MAPPING_ROWS.
    """
    df = pd.DataFrame(MAPPING_ROWS, columns=["Uzio Field Key", "PayCom Column", "Comments"])
    df = df[(df["Uzio Field Key"] != "") & (df["PayCom Column"] != "")]
    return df.drop_duplicates(subset=["Uzio Field Key", "PayCom Column"], keep="first").reset_index(drop=True)

def _pivot_uzio_long_to_wide(df_long):
    """Pivot UZIO long-format into two wide views:

      federal_wide — index = employee_id, columns = FIT_* / FICA_* / FUTA_*
                     (rows with empty state_code or scope = FEDERAL).
      state_wide   — index = (employee_id, state_code), columns = SIT_* etc.
                     (rows where state_code is populated).

    The previous flat pivot collapsed all per-state rows into one row per
    employee with `aggfunc='first'`, which silently picked an arbitrary
    state's value for SIT fields. For a multi-state employee (e.g. lives in
    MI, works in OH) this caused systematic SIT mismatches against the
    wrong state's record. See user-reported A0PO / A06U cases.
    """
    required = {"employee_id", "withholding_field_key", "withholding_field_value"}
    missing = required - set(df_long.columns)
    if missing:
        raise ValueError(f"UZIO CSV missing required columns: {sorted(missing)}")

    uz = df_long.copy()
    # pandas read_csv(dtype=str) turns true NaN into the literal string "nan"
    # — both `.fillna("")` and a simple `!= ""` check miss that. Normalize first.
    def _clean_str(s):
        s = s.fillna("").astype(str).str.strip()
        return s.where(~s.str.lower().isin(["nan", "none", "nat"]), "")
    for c in required:
        uz[c] = _clean_str(uz[c])
    if "state_code" in uz.columns:
        uz["state_code"] = _clean_str(uz["state_code"]).str.upper()
    else:
        uz["state_code"] = ""

    is_state = uz["state_code"] != ""

    fed = uz[~is_state]
    if not fed.empty:
        federal_wide = fed.pivot_table(
            index="employee_id", columns="withholding_field_key",
            values="withholding_field_value", aggfunc="first",
        ).reset_index()
    else:
        federal_wide = pd.DataFrame(columns=["employee_id"])

    st = uz[is_state]
    if not st.empty:
        state_wide = st.pivot_table(
            index=["employee_id", "state_code"], columns="withholding_field_key",
            values="withholding_field_value", aggfunc="first",
        ).reset_index()
    else:
        state_wide = pd.DataFrame(columns=["employee_id", "state_code"])

    # Names (only need from federal-side; same per employee).
    name_cols = [c for c in ("employee_first_name", "employee_last_name") if c in uz.columns]
    if name_cols:
        names = uz.groupby("employee_id")[name_cols].first().reset_index()
        federal_wide = federal_wide.merge(names, on="employee_id", how="left")

    federal_wide.columns = [str(c) for c in federal_wide.columns]
    state_wide.columns = [str(c) for c in state_wide.columns]
    return federal_wide, state_wide

def _load_labels_by_state() -> dict:
    """Return the inlined field-label-per-jurisdiction dict.
    Was loaded from key_mapping.yml; now from withholding_data.LABELS_BY_STATE.
    """
    return LABELS_BY_STATE


def _load_filing_status_code() -> dict:
    """Return the inlined UZIO filing-status enum -> label map.
    Was loaded from filing status_code.txt; now from withholding_data.FILING_STATUS_MAP.
    """
    return FILING_STATUS_MAP

def _norm_text(s):
    s = "" if pd.isna(s) else str(s)
    s = s.strip().lower()
    s = re.sub(r"[^a-z0-9\s]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _parse_bool(raw):
    if pd.isna(raw): return None
    s = str(raw).strip().lower()
    if s == "" or s == "nan": return None
    if s in {"yes", "y", "true", "1", "t", "on"}: return True
    if s in {"no", "n", "false", "0", "f", "off"}: return False
    return None

def _parse_number(raw):
    if pd.isna(raw): return 0.0
    s = str(raw).strip()
    if s == "" or s.lower() == "nan": return 0.0
    s = s.replace(",", "").replace("$", "")
    try:
        if s.startswith("(") and s.endswith(")"):
            return -float(s[1:-1])
        return float(s)
    except:
        return 0.0

def _resolve_uz_key_for_row(uz_key_raw: str, work_state: str) -> str:
    """Resolve a mapping row's 'Uzio Field Key' string to the actual UZIO key
    we should look up, applying state-specific business rules where they exist.

    Why this exists: a few rows in Mapping.xlsx encode special rules in the
    `Comments` column rather than as separate mapping rows. We honor those
    rules here. The mapping file is the source of truth — when a new rule
    appears in Comments, add the corresponding case here.

    Documented rules:
      - "SIT_TOTAL_ALLOWANCES / SIT_TOTAL_ALLOWANCES_VALUE"
            Comment: "Use SIT_TOTAL_ALLOWANCES_VALUES if State is IA,
                      Map with Personal Exemptions"
            Behavior: state == 'IA' -> SIT_TOTAL_ALLOWANCES_VALUE
                      otherwise     -> SIT_TOTAL_ALLOWANCES

    All other keys are returned as-is (single field key, no transformation).
    """
    key = (uz_key_raw or "").strip()
    state = (work_state or "").strip().upper()

    # Iowa special case for total allowances.
    if "SIT_TOTAL_ALLOWANCES" in key and "SIT_TOTAL_ALLOWANCES_VALUE" in key:
        return "SIT_TOTAL_ALLOWANCES_VALUE" if state == "IA" else "SIT_TOTAL_ALLOWANCES"

    # No rule registered → use the raw key (which is the normal case for
    # all the single-key rows in Mapping.xlsx).
    return key


def _infer_type(uzio_key, paycom_col):
    """Infer a comparison type from the UZIO field key.

    Order matters. Booleans must be checked FIRST because field keys like
    FIT_WITHHOLDING_EXEMPTION, SIT_WITHHOLDING_EXEMPTION, and
    FIT_HIGHER_WITHHOLDING all contain the substring 'WITHHOLDING' — which
    the amount branch matches on — but they're actually boolean Yes/No flags.
    Tagging them as 'amount' parses 'Yes'/'true' as numeric 0 on both sides,
    so the values always match silently, and the tool reports no mismatch.
    """
    k = (uzio_key or "").upper()
    pc = (paycom_col or "").lower()
    if k in {"FIT_FILING_STATUS", "SIT_FILING_STATUS"}:
        return "filing_status"
    # Booleans — check before amount so EXEMPT/HIGHER fields aren't miscast.
    if any(x in k for x in ["EXEMPT", "FLAG", "HIGHER", "NON_RESIDENT", "RESIDENT", "CERTIFICATE", "MULTIPLE_JOBS"]):
        return "boolean"
    if ("$" in pc) or any(x in k for x in ["OTHER_INCOME", "ADDL", "WITHHOLDING", "CREDIT", "DEDUCTION", "OVERRIDE"]):
        return "amount"
    if any(x in k for x in ["ALLOWANCE", "EXEMPTION", "NUMBER", "TOTAL", "COUNT"]):
        return "integer"
    return "string"

def _field_label_for(uzio_key, state, labels_by_state):
    key = (uzio_key or "").strip()
    st_code = (state or "").strip().upper()
    if st_code and st_code in labels_by_state and key in labels_by_state[st_code]: return labels_by_state[st_code][key]
    if "FED" in labels_by_state and key in labels_by_state["FED"]: return labels_by_state["FED"][key]
    for mp in labels_by_state.values():
        if key in mp: return mp[key]
    return key.replace("_", " ").title()

# --- Comparisons ---
def _compare_filing_status(pay_raw, uz_code, filing_map):
    pc = "" if pd.isna(pay_raw) else str(pay_raw).strip()
    uz = "" if pd.isna(uz_code) else str(uz_code).strip()
    if pc == "" and uz == "": return True, "", "", "Both blank"
    if uz == "": return False, pc, "", "Value missing in UZIO"
    ui = filing_map.get(uz)
    if ui is None: return False, pc, "", "Filing status code not found in mapping file"
    pc_n = _norm_text(pc)
    ui_n = _norm_text(ui)
    if pc_n == ui_n or (pc_n and pc_n in ui_n) or (ui_n and ui_n in pc_n): return True, pc, ui, "Matched normalized UI string"
    return False, pc, ui, "Filing Status mismatch"

def _compare_amount(pay_raw, uz_raw):
    pc = _parse_number(pay_raw)
    uz = _parse_number(uz_raw)
    uz_dollars = round(uz / 100.0, 2) if uz_raw else 0.0
    pc_dollars = round(pc, 2)
    # Tolerance is HALF a cent. Previously was `< 0.01` which combined with
    # float-precision noise (e.g. 10.0 - 10.01 ~ 0.0099999) let exact-1-cent
    # differences silently pass as matches. Tighten to 0.005 so genuine
    # 1-cent disagreements register as mismatches.
    return (abs(pc_dollars - uz_dollars) < 0.005), str(pc_dollars), str(uz_dollars), "Divide UZIO by 100 (Cents), blank=0"

def _compare_integer(pay_raw, uz_raw):
    pc_i = int(_parse_number(pay_raw))
    uz_i = int(_parse_number(uz_raw))
    return (pc_i == uz_i), str(pc_i), str(uz_i), "Integer match, blank=0"

def _compare_boolean(pay_raw, uz_raw):
    pc_b = _parse_bool(pay_raw)
    uz_b = _parse_bool(uz_raw)
    if pc_b is None and uz_b is None: return True, "", "", "Both blank"
    if pc_b is None or uz_b is None: return False, str(pc_b), str(uz_b), "Blank vs Value"
    return (pc_b == uz_b), str(pc_b), str(uz_b), "Boolean match"

def _compare_string(pay_raw, uz_raw):
    pc = "" if pd.isna(pay_raw) else str(pay_raw).strip()
    uz = "" if pd.isna(uz_raw) else str(uz_raw).strip()
    if pc == "" and uz == "": return True, "", "", "Both blank"
    return (_norm_text(pc) == _norm_text(uz)), pc, uz, "String match"


def run_withholding_audit(paycom_df, uzio_long_df, mapping_df, labels_by_state, filing_map,
                         paycom_emp_id_col, paycom_status_col, paycom_state_col, paycom_fn_col, paycom_ln_col):

    federal_wide, state_wide = _pivot_uzio_long_to_wide(uzio_long_df)

    # Normalize IDs
    pay = paycom_df.copy()
    pay[paycom_emp_id_col] = pay[paycom_emp_id_col].astype(str).fillna("").str.strip()
    federal_wide["employee_id"] = federal_wide["employee_id"].astype(str).fillna("").str.strip()
    if not state_wide.empty:
        state_wide["employee_id"] = state_wide["employee_id"].astype(str).fillna("").str.strip()
        state_wide["state_code"] = state_wide["state_code"].astype(str).fillna("").str.strip().str.upper()

    # Index state_wide for fast (emp_id, state) lookup.
    state_wide_idx = state_wide.set_index(["employee_id", "state_code"]) if not state_wide.empty else pd.DataFrame()
    states_by_emp = (
        state_wide.groupby("employee_id")["state_code"].apply(lambda s: sorted(set(s))).to_dict()
        if not state_wide.empty else {}
    )

    pay_ids = set(pay[paycom_emp_id_col].replace("nan", "").replace("", pd.NA).dropna().tolist())
    uz_ids = set(federal_wide["employee_id"].replace("nan", "").replace("", pd.NA).dropna().tolist())
    if not state_wide.empty:
        uz_ids |= set(state_wide["employee_id"].replace("nan", "").replace("", pd.NA).dropna().tolist())
    all_ids = sorted(list(pay_ids | uz_ids))

    pay_idx = {str(x): i for i, x in enumerate(pay[paycom_emp_id_col].astype(str))}
    fed_idx = {str(x): i for i, x in enumerate(federal_wide["employee_id"].astype(str))}

    mismatches = []
    fields_used = []

    for eid in all_ids:
        p_i = pay_idx.get(eid)
        u_i = fed_idx.get(eid)

        p_missing_row = p_i is None
        u_missing_fed = u_i is None
        u_missing_states = eid not in states_by_emp

        p_status = str(pay.loc[p_i, paycom_status_col]) if not p_missing_row and paycom_status_col in pay.columns else ""
        # p_state = work-location state (priority 1). Used both for SIT lookups
        # and for displaying which state the comparison was run against.
        p_state = ""
        if not p_missing_row and paycom_state_col and paycom_state_col in pay.columns:
            p_state = str(pay.loc[p_i, paycom_state_col]).strip().upper()
            if p_state.lower() == "nan":
                p_state = ""
        p_first = str(pay.loc[p_i, paycom_fn_col]) if not p_missing_row and paycom_fn_col and paycom_fn_col in pay.columns else ""
        p_last = str(pay.loc[p_i, paycom_ln_col]) if not p_missing_row and paycom_ln_col and paycom_ln_col in pay.columns else ""

        u_first = str(federal_wide.loc[u_i, "employee_first_name"]) if not u_missing_fed and "employee_first_name" in federal_wide.columns else ""
        u_last = str(federal_wide.loc[u_i, "employee_last_name"]) if not u_missing_fed and "employee_last_name" in federal_wide.columns else ""

        for _, mr in mapping_df.iterrows():
            uz_key_raw = mr["Uzio Field Key"]
            pc_col = mr["PayCom Column"]

            # Resolve the UZIO field key. Most rows are a single key, but a few
            # composite keys encode a state-dependent business rule from the
            # Mapping.xlsx Comments column. Handled explicitly per-row rather
            # than via blind slash-split, because the rule isn't generic.
            uz_key = _resolve_uz_key_for_row(uz_key_raw, p_state)
            is_sit = uz_key.startswith("SIT_")

            # Route the UZIO lookup to the right table.
            uz_val = ""
            u_missing_col = True
            if is_sit:
                # SIT fields must match by (emp_id, work-location state).
                if p_state and not u_missing_states and (eid, p_state) in state_wide_idx.index:
                    state_row = state_wide_idx.loc[(eid, p_state)]
                    if uz_key in state_row.index:
                        uz_val = state_row.get(uz_key, "")
                        u_missing_col = False
            else:
                if not u_missing_fed and uz_key in federal_wide.columns:
                    uz_val = federal_wide.loc[u_i, uz_key]
                    u_missing_col = False

            p_missing_col = (pc_col not in pay.columns)
            pc_val = pay.loc[p_i, pc_col] if not p_missing_row and not p_missing_col else ""

            dtype = _infer_type(uz_key, pc_col)
            label = _field_label_for(uz_key, p_state, labels_by_state)

            if eid == all_ids[0]:
                fields_used.append({
                    "Uzio Field Key": uz_key,
                    "Field Label": label,
                    "Paycom Column": pc_col,
                    "Data Type": dtype,
                    "Logic": mr["Comments"],
                })

            # Skip when we genuinely have nothing to compare on one side.
            if p_missing_row or p_missing_col or u_missing_col:
                continue
            if is_sit and not p_state:
                continue  # no work-location state known
            if not is_sit and u_missing_fed:
                continue  # employee not in UZIO at federal level

            is_match = False
            if dtype == "filing_status":
                is_match, pn, un, r = _compare_filing_status(pc_val, uz_val, filing_map)
            elif dtype == "amount":
                is_match, pn, un, r = _compare_amount(pc_val, uz_val)
            elif dtype == "integer":
                is_match, pn, un, r = _compare_integer(pc_val, uz_val)
            elif dtype == "boolean":
                is_match, pn, un, r = _compare_boolean(pc_val, uz_val)
            else:
                is_match, pn, un, r = _compare_string(pc_val, uz_val)

            if not is_match:
                mismatches.append({
                    "Employee ID": eid,
                    "Paycom Status": p_status,
                    "Paycom State": p_state,
                    "Paycom First Name": p_first,
                    "Paycom Last Name": p_last,
                    "UZIO First Name": u_first,
                    "UZIO Last Name": u_last,
                    "Field Label": label,
                    "Paycom Column": pc_col,
                    "Paycom Value": str(pc_val),
                    "UZIO Field Key": uz_key,
                    "UZIO Stored Value": str(uz_val),
                    "Paycom Normalized": pn,
                    "UZIO Normalized / UI": un,
                    "Rule Applied": (
                        f"{r}  [matched against work-location state {p_state}]"
                        if is_sit else r
                    ),
                })

    all_miss_df = pd.DataFrame(mismatches, columns=DETAIL_COLUMNS)
    
    act_miss_df = pd.DataFrame(columns=DETAIL_COLUMNS)
    if not all_miss_df.empty:
        act_miss_df = all_miss_df[all_miss_df["Paycom Status"].str.lower().isin(ACTIVE_STATUSES)].copy()

    # Summary
    sum_data = {
        "Metric": [
            "Total Paycom employees",
            "Total UZIO employees",
            "Employees missing in UZIO (Paycom-only)",
            "# mapped fields compared",
            "Total mismatches (mapped only)",
            "Active mismatches (mapped only)"
        ],
        "Value": [
            len(pay_ids),
            len(uz_ids),
            len(pay_ids - uz_ids),
            len(mapping_df),
            len(all_miss_df),
            len(act_miss_df)
        ]
    }
    summary_df = pd.DataFrame(sum_data)

    # Missing in UZIO
    missing_in_uzio_df = pd.DataFrame(columns=["Employee ID", "Status", "First Name", "Last Name", "State", "Position", "Work Location"])
    missing_ids = list(pay_ids - uz_ids)
    if missing_ids:
        m_list = []
        for eid in missing_ids:
            p_i = pay_idx.get(eid)
            m_list.append({
                "Employee ID": eid,
                "Status": str(pay.loc[p_i, paycom_status_col]) if paycom_status_col in pay.columns else "",
                "First Name": str(pay.loc[p_i, paycom_fn_col]) if paycom_fn_col and paycom_fn_col in pay.columns else "",
                "Last Name": str(pay.loc[p_i, paycom_ln_col]) if paycom_ln_col and paycom_ln_col in pay.columns else "",
                "State": str(pay.loc[p_i, paycom_state_col]) if paycom_state_col and paycom_state_col in pay.columns else "",
                "Position": "",
                "Work Location": ""
            })
        missing_in_uzio_df = pd.DataFrame(m_list)

    # Maps
    field_map_df = pd.DataFrame(fields_used)
    filing_ui_df = pd.DataFrame(list(filing_map.items()), columns=["DB Key", "UI Label"])
    norm_rules_df = pd.DataFrame({"Normalization Rules Applied": NORMALIZATION_NOTES})

    return summary_df, act_miss_df, all_miss_df, missing_in_uzio_df, field_map_df, filing_ui_df, norm_rules_df

def build_report_bytes(sum_df, act_df, all_df, miss_df, f_map_df, ui_map_df, rules_df):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        sum_df.to_excel(writer, sheet_name="Summary", index=False)
        act_df.to_excel(writer, sheet_name="Active_Mismatches", index=False)
        all_df.to_excel(writer, sheet_name="All_Mismatches", index=False)
        miss_df.to_excel(writer, sheet_name="Missing_in_UZIO", index=False)
        f_map_df.to_excel(writer, sheet_name="Field_Mapping_Used", index=False)
        ui_map_df.to_excel(writer, sheet_name="FilingStatus_UI_Map", index=False)
        rules_df.to_excel(writer, sheet_name="Normalization_Rules", index=False)

        # Auto format columns
        for sheet in writer.sheets.values():
            for col in sheet.columns:
                max_length = 0
                c = col[0].column_letter
                for cell in col:
                    try: max_length = max(max_length, len(str(cell.value)))
                    except: pass
                sheet.column_dimensions[c].width = min(max_length + 2, 60)
    return out.getvalue()

def render_ui():
    st.title(APP_TITLE)
    client_name = st.text_input("Client Name", value="Client", key="paycom_withholding_client")

    c1, c2 = st.columns(2)
    with c1: paycom_file = st.file_uploader("Paycom export (CSV/XLSX)", type=["csv", "xlsx"])
    with c2: uzio_file = st.file_uploader("UZIO export (CSV/XLSX - long format)", type=["csv", "xlsx"])

    if st.button("Run Audit", type="primary", disabled=not (paycom_file and uzio_file)):
        with st.spinner("Running audit..."):
            try:
                def read_file(f):
                    f.seek(0)
                    if f.name.lower().endswith(".csv"):
                        return pd.read_csv(io.BytesIO(f.getvalue()), dtype=str, keep_default_na=False)
                    return pd.read_excel(io.BytesIO(f.getvalue()), engine="openpyxl", dtype=str, keep_default_na=False)

                paycom_df = read_file(paycom_file)
                uzio_long_df = read_file(uzio_file)

                # Config is fully inlined — no external file lookup.
                mapping_df = _load_mapping_df()
                labels_by_state = _load_labels_by_state()
                filing_map = _load_filing_status_code()

                emp_id_col, status_col, state_col, fn_col, ln_col = _autodetect_paycom_cols(paycom_df)

                s_df, act_df, all_df, miss_df, f_map_df, ui_map_df, rules_df = run_withholding_audit(
                    paycom_df=paycom_df, uzio_long_df=uzio_long_df, mapping_df=mapping_df,
                    labels_by_state=labels_by_state, filing_map=filing_map,
                    paycom_emp_id_col=emp_id_col, paycom_status_col=status_col,
                    paycom_state_col=state_col, paycom_fn_col=fn_col, paycom_ln_col=ln_col
                )

                rep_bytes = build_report_bytes(s_df, act_df, all_df, miss_df, f_map_df, ui_map_df, rules_df)

                st.success("Report generated successfully.")
                timestamp = pd.Timestamp.now().strftime('%d_%m_%Y_%H%M')
                st.download_button(
                    label="Download Audit Report",
                    data=rep_bytes,
                    file_name=f"{client_name}_Uzio_Paycom_Withholding_Audit_Report_{timestamp}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

                st.subheader("Summary (preview)")
                st.dataframe(s_df, use_container_width=True)

            except Exception as e:
                st.error(f"Error: {e}")
                st.exception(e)

if __name__ == "__main__":
    st.set_page_config(layout="wide")
    render_ui()
