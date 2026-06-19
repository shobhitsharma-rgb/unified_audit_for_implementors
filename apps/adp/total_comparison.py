import streamlit as st
import pandas as pd
import io
import re
from utils.audit_utils import clean_money_val, norm_colname

def load_mapping(file, cat_name, adp_col, uzio_col):
    """Load a mapping file and return a list of mappings (ADP_Name, UZIO_Name)."""
    try:
        file.seek(0)
        if str(file.name).lower().endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        # Normalize headers to find columns
        df.columns = [norm_colname(c) for c in df.columns]
        
        # Finding the actual column names in the sheet
        actual_adp_col = next((c for c in df.columns if adp_col.lower() in c.lower()), None)
        actual_uzio_col = next((c for c in df.columns if uzio_col.lower() in c.lower()), None)
        
        if not actual_adp_col or not actual_uzio_col:
            st.warning(f"Could not find exact columns in {cat_name} mapping. Looking for '{adp_col}' and '{uzio_col}'. Available: {list(df.columns)}")
            return []
            
        mappings = []
        for _, row in df.iterrows():
            a_val = str(row[actual_adp_col]).strip()
            u_val = str(row[actual_uzio_col]).strip()
            if a_val and u_val and a_val.lower() != 'nan' and u_val.lower() != 'nan':
                mappings.append({
                    "Category": cat_name,
                    "ADP_Name": a_val,
                    "UZIO_Name": u_val
                })
        return mappings
    except Exception as e:
        st.error(f"Error loading {cat_name} mapping: {e}")
        return []

def format_pay_date(date_val):
    if pd.isna(date_val) or str(date_val).strip() in ["", "nan", "NaT"]:
        return "Unknown"
    try:
        dt = pd.to_datetime(date_val)
        return dt.strftime('%Y-%m-%d')
    except:
        return str(date_val).strip()

def normalize_id(id_val):
    """Strip hyphens, leading zeros, and whitespace from ID for matching."""
    if pd.isna(id_val):
        return "Unknown"
    s = str(id_val).replace('-', '').replace(' ', '').strip()
    return s.lstrip('0') if s else "Unknown"

def find_header_and_data(file):
    """Find the correct header row and read the data, skipping metadata sheets."""
    file.seek(0)
    if str(file.name).lower().endswith('.csv'):
        file.seek(0)
        df_peek = pd.read_csv(file, header=None, nrows=50)
        header_idx = 0
        for i, row in df_peek.iterrows():
            row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
            if "employee id" in row_str or "employee name" in row_str:
                header_idx = i
                break
                
        file.seek(0)
        df = pd.read_csv(file, header=header_idx)
        
        header_top = None
        if header_idx > 0:
            header_top = df_peek.iloc[header_idx - 1].tolist()
            
        target_sheet = "Sheet1"
    else:
        xls = pd.ExcelFile(file)
        target_sheet = xls.sheet_names[0]
        
        if len(xls.sheet_names) > 1 and "criteria" in xls.sheet_names[0].lower():
            target_sheet = xls.sheet_names[1]
        
        df_peek = pd.read_excel(xls, sheet_name=target_sheet, header=None, nrows=50)
        header_idx = 0
        for i, row in df_peek.iterrows():
            row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
            if "employee id" in row_str or "employee name" in row_str:
                header_idx = i
                break
                
        df = pd.read_excel(xls, sheet_name=target_sheet, header=header_idx)
        
        header_top = None
        if header_idx > 0:
            header_top = df_peek.iloc[header_idx - 1].tolist()

    # --- GRAND TOTAL ROW DETECTION ---
    # Sometimes ADP exports include a grand total at the very bottom but fail to clear
    # the last employee's ID from that row, messing up totals for that employee.
    if len(df) > 1:
        last_row = df.iloc[-1]
        prev_row = df.iloc[-2]
        
        shared_cols = 0
        for c in df.columns[:5]:
            v_last = str(last_row[c]).strip()
            v_prev = str(prev_row[c]).strip()
            if v_last and v_last == v_prev and v_last.lower() != 'nan':
                shared_cols += 1
                
        if shared_cols >= 1:
            for c in df.columns:
                try:
                    val_last = clean_money_val(last_row[c])
                    if val_last > 100:
                        sum_rest = sum(clean_money_val(x) for x in df[c].iloc[:-1])
                        if sum_rest > 0 and abs(val_last - sum_rest) < sum_rest * 0.05:
                            df = df.iloc[:-1]
                            break
                except:
                    continue
                    
    return df, header_top, target_sheet

def calculate_totals(df, header_top, column_names):
    """Sum up values for columns that match any of the provided names, handling multi-row headers."""
    found_cols = []
    emp_tots = {}
    emp_row_counts = {}
    
    # --- STRICT ROW FILTERING ---
    id_col = next((c for c in df.columns if any(x in str(c).lower() for x in ["associate id", "employee id", "file #"])), None)
    # Prioritize 'pay date' / 'check date' before 'period end' to avoid using quarterly
    # period end dates instead of actual pay dates when both columns exist in the file.
    date_col = next((c for c in df.columns if any(x == str(c).lower().strip() for x in ["pay date", "check date"])), None)
    if date_col is None:
        date_col = next((c for c in df.columns if any(x in str(c).lower() for x in ["pay date", "period end", "check date"])), None)
    
    if id_col:
        df_clean = df[df[id_col].notna()].copy()
        df_clean[id_col] = df_clean[id_col].apply(normalize_id)
        df_clean = df_clean[
            (df_clean[id_col] != "Unknown") & 
            (~df_clean[id_col].str.lower().str.contains("total|grand", na=False))
        ]
    else:
        mask = df.iloc[:, 0].astype(str).str.lower().str.contains("total|grand", na=False)
        df_clean = df[~mask].copy()
    
    norm_cols_main = {norm_colname(c).lower(): i for i, c in enumerate(df.columns)}
    norm_cols_top = {}
    if header_top:
        for i, c in enumerate(header_top):
            if pd.notna(c) and str(c).strip() != "":
                norm_cols_top[norm_colname(c).lower()] = i

    cols_to_sum = []
    for name in column_names:
        n_name = norm_colname(name).lower()
        if n_name in norm_cols_main:
            idx = norm_cols_main[n_name]
            cols_to_sum.append(df.columns[idx])
            found_cols.append(df.columns[idx])
        elif n_name in norm_cols_top:
            start_idx = norm_cols_top[n_name]
            end_idx = len(df.columns)
            if header_top:
                for k in range(start_idx + 1, len(header_top)):
                    if pd.notna(header_top[k]) and str(header_top[k]).strip() != "":
                        end_idx = k
                        break
            for k in range(start_idx, end_idx):
                main_h = str(df.columns[k]).lower()
                if any(x in main_h for x in ['amount', 'total', 'current', 'ee', 'er', 'tax']):
                    if not any(x in main_h for x in ['wages', 'hours', 'rate', 'basis', 'taxable']):
                        cols_to_sum.append(df.columns[k])
                        found_cols.append(f"{df.columns[k]}")
                        
    for _, row in df_clean.iterrows():
        eid = row[id_col] if id_col else "Unknown"
        pay_date = format_pay_date(row[date_col]) if date_col else "Unknown"
        
        row_tot = sum(clean_money_val(row[c]) for c in set(cols_to_sum))
        
        key = (eid, pay_date)
        if key not in emp_tots:
            emp_tots[key] = 0.0
            emp_row_counts[key] = 0
        emp_tots[key] += row_tot
        emp_row_counts[key] += 1
            
    return sum(emp_tots.values()), found_cols, emp_tots, emp_row_counts

def detect_duplicate_pay_periods(df):
    """Find employees with more than one row for the same Start Date / End Date / Pay Date.
    Returns a DataFrame of every row that participates in a duplicate group, sorted so
    duplicates appear together. Empty DataFrame when no duplicates found.
    """
    def find_col(candidates):
        for cand in candidates:
            for c in df.columns:
                if str(c).strip().lower() == cand.lower():
                    return c
        return None

    eid_col   = find_col(["Employee ID", "Associate ID", "File #"])
    first_col = find_col(["First Name"])
    last_col  = find_col(["Last Name"])
    start_col = find_col(["Start Date", "Period Start"])
    end_col   = find_col(["End Date", "Period End"])
    pay_col   = find_col(["Pay Date", "Check Date"])
    gross_col = find_col(["Gross Pay", "Gross"])
    net_col   = find_col(["Net Pay", "Net"])

    if not all([eid_col, start_col, end_col, pay_col]):
        return pd.DataFrame()

    work = df[df[eid_col].notna()].copy()
    work[eid_col] = work[eid_col].astype(str).str.strip()
    work = work[(work[eid_col] != "") & (~work[eid_col].str.lower().str.contains("total|grand", na=False))]

    group_keys = [eid_col, start_col, end_col, pay_col]
    counts = work.groupby(group_keys).size().reset_index(name="_n")
    dup_keys = counts[counts["_n"] > 1]
    if dup_keys.empty:
        return pd.DataFrame()

    dup_rows = work.merge(dup_keys, on=group_keys, how="inner")

    def classify_row(row):
        if gross_col and clean_money_val(row.get(gross_col)) != 0:
            return "Detail (real values)"
        dash_count = sum(1 for v in row.values if str(v).strip() == "-")
        if dash_count >= 5:
            return "Skeleton (dashes / zeros)"
        return "Zero detail"

    dup_rows["Row Type"] = dup_rows.apply(classify_row, axis=1)

    name_parts = []
    if first_col:
        name_parts.append(dup_rows[first_col].fillna("").astype(str).str.strip())
    if last_col:
        name_parts.append(dup_rows[last_col].fillna("").astype(str).str.strip())
    if name_parts:
        dup_rows["Employee Name"] = name_parts[0]
        for p in name_parts[1:]:
            dup_rows["Employee Name"] = (dup_rows["Employee Name"] + " " + p).str.strip()
    else:
        dup_rows["Employee Name"] = ""

    out_cols = {
        eid_col: "Employee ID",
        "Employee Name": "Employee Name",
        start_col: "Start Date",
        end_col: "End Date",
        pay_col: "Pay Date",
        "Row Type": "Row Type",
    }
    if gross_col: out_cols[gross_col] = "Gross Pay"
    if net_col:   out_cols[net_col]   = "Net Pay"
    out_cols["_n"] = "Rows in Group"

    out = dup_rows[list(out_cols.keys())].rename(columns=out_cols)
    out = out.sort_values(["Employee ID", "Pay Date", "Row Type"]).reset_index(drop=True)
    return out


def compute_pay_stub_count_diff(adp_data_list, df_uzio):
    """Per employee, compare distinct Pay Date count between combined ADP files and UZIO file."""

    def find_col(df, candidates):
        for cand in candidates:
            for c in df.columns:
                if str(c).strip().lower() == cand.lower():
                    return c
        for cand in candidates:
            for c in df.columns:
                if cand.lower() in str(c).strip().lower():
                    return c
        return None

    adp_stubs, adp_names = {}, {}
    for df_adp, _, _ in adp_data_list:
        eid_col   = find_col(df_adp, ["Associate ID", "Employee ID", "File #"])
        pay_col   = find_col(df_adp, ["Check Date", "Pay Date", "Pay Period End", "Period End Date"])
        first_col = find_col(df_adp, ["First Name", "Employee First Name"])
        last_col  = find_col(df_adp, ["Last Name", "Employee Last Name"])
        full_col  = find_col(df_adp, ["Employee Name", "Name"])
        if not eid_col or not pay_col:
            continue
        for _, row in df_adp.iterrows():
            raw_eid = str(row[eid_col]).strip()
            if not raw_eid or raw_eid.lower() in ("nan", "total", "grand"):
                continue
            pay = format_pay_date(row[pay_col])
            if pay == "Unknown":
                continue
            key = normalize_id(raw_eid)
            if key == "Unknown":
                continue
            adp_stubs.setdefault(key, set()).add(pay)
            if key not in adp_names:
                if first_col or last_col:
                    fn = str(row[first_col]).strip() if first_col and pd.notna(row[first_col]) else ""
                    ln = str(row[last_col]).strip() if last_col and pd.notna(row[last_col]) else ""
                    nm = (fn + " " + ln).strip()
                elif full_col and pd.notna(row[full_col]):
                    nm = str(row[full_col]).strip()
                else:
                    nm = ""
                if nm:
                    adp_names[key] = nm

    eid_col   = find_col(df_uzio, ["Employee ID"])
    pay_col   = find_col(df_uzio, ["Pay Date"])
    first_col = find_col(df_uzio, ["First Name"])
    last_col  = find_col(df_uzio, ["Last Name"])

    uzio_stubs, uzio_names = {}, {}
    if eid_col and pay_col:
        for _, row in df_uzio.iterrows():
            raw_eid = str(row[eid_col]).strip()
            if not raw_eid or raw_eid.lower() in ("nan", "total", "grand"):
                continue
            pay = format_pay_date(row[pay_col])
            if pay == "Unknown":
                continue
            key = normalize_id(raw_eid)
            if key == "Unknown":
                continue
            uzio_stubs.setdefault(key, set()).add(pay)
            if key not in uzio_names:
                fn = str(row[first_col]).strip() if first_col and pd.notna(row[first_col]) else ""
                ln = str(row[last_col]).strip() if last_col and pd.notna(row[last_col]) else ""
                nm = (fn + " " + ln).strip()
                if nm:
                    uzio_names[key] = nm

    rows = []
    all_keys = set(adp_stubs.keys()) | set(uzio_stubs.keys())
    for k in all_keys:
        a_dates = adp_stubs.get(k, set())
        u_dates = uzio_stubs.get(k, set())
        a_n, u_n = len(a_dates), len(u_dates)
        diff = u_n - a_n
        if diff == 0:
            status = "Match"
        elif diff > 0:
            status = "Extra in UZIO"
        else:
            status = "Missing in UZIO"
        missing_in_uzio = sorted(a_dates - u_dates)
        missing_in_adp  = sorted(u_dates - a_dates)
        rows.append({
            "Employee ID": k,
            "Employee Name": uzio_names.get(k) or adp_names.get(k, ""),
            "ADP Pay Stubs": a_n,
            "UZIO Pay Stubs": u_n,
            "Difference": diff,
            "Status": status,
            "Pay Dates Missing in UZIO": ", ".join(missing_in_uzio) if missing_in_uzio else "",
            "Pay Dates Missing in ADP":  ", ".join(missing_in_adp)  if missing_in_adp  else "",
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["_sort"] = df["Status"].apply(lambda s: 0 if s != "Match" else 1)
    df = df.sort_values(["_sort", "Employee ID"]).drop(columns=["_sort"]).reset_index(drop=True)
    return df


# Standard federal tax rates used for verification (in percent)
STANDARD_TAX_RATES = {
    "Social Security EE": 6.20,
    "Social Security ER": 6.20,
    "Medicare EE":        1.45,
    "Medicare ER":        1.45,
    "FUTA ER":            0.60,
}
RATE_TOLERANCE_PCT = 0.05


def _filter_data_rows(df, eid_col):
    """Drop blank/total rows so sums aren't doubled by grand-total rows."""
    if not eid_col:
        return df
    work = df[df[eid_col].notna()].copy()
    work[eid_col] = work[eid_col].astype(str).str.strip()
    return work[(work[eid_col] != "") & (~work[eid_col].str.lower().str.contains("total|grand", na=False))]


def _sum_uzio_section(df, header_top, section_name, side):
    """Sum Taxable Wages and EE/ER Amount within a UZIO section header. side is 'EE' or 'ER'."""
    if not header_top:
        return 0.0, 0.0
    eid_col = next((c for c in df.columns if any(x in str(c).lower() for x in ["employee id", "associate id"])), None)
    work = _filter_data_rows(df, eid_col)
    target = norm_colname(section_name).lower()
    wages = amount = 0.0
    for i, h in enumerate(header_top):
        if pd.notna(h) and norm_colname(str(h)).lower() == target:
            end_i = len(df.columns)
            for j in range(i + 1, len(header_top)):
                if pd.notna(header_top[j]) and str(header_top[j]).strip() != "":
                    end_i = j
                    break
            for k in range(i, end_i):
                col = str(df.columns[k]).strip().lower()
                if "taxable wages" in col:
                    wages += work.iloc[:, k].apply(clean_money_val).sum()
                elif side == "EE" and (col == "ee amount" or col.startswith("ee amount.")):
                    amount += work.iloc[:, k].apply(clean_money_val).sum()
                elif side == "ER" and (col == "er amount" or col.startswith("er amount.")):
                    amount += work.iloc[:, k].apply(clean_money_val).sum()
            break
    return wages, amount


def _sum_adp_for_uzio_name(adp_data_list, adp_names, side):
    """Best-effort sum of (taxable wages, amount) across ADP files for the given source codes."""
    if not adp_names:
        return 0.0, 0.0
    total_w = total_a = 0.0
    for df_adp, adp_top, _ in adp_data_list:
        eid_col = next((c for c in df_adp.columns if any(x in str(c).lower() for x in ["associate id", "employee id", "file #"])), None)
        work = _filter_data_rows(df_adp, eid_col)
        norm_main = {norm_colname(c).lower(): i for i, c in enumerate(df_adp.columns)}
        norm_top  = {norm_colname(str(c)).lower(): i for i, c in enumerate(adp_top or []) if pd.notna(c) and str(c).strip() != ""}
        for name in adp_names:
            n = norm_colname(name).lower()
            if n in norm_main:
                idx = norm_main[n]
                total_a += work.iloc[:, idx].apply(clean_money_val).sum()
                tax_col = str(df_adp.columns[idx])
                cand_names = []
                if re.search(r"\btax\b", tax_col, re.I):
                    cand_names.append(re.sub(r"\btax\b", "Wages", tax_col, flags=re.I))
                cand_names.extend([tax_col + " Wages", tax_col + " Taxable Wages"])
                found_wages = False
                for cn in cand_names:
                    nn = norm_colname(cn).lower()
                    if nn in norm_main:
                        total_w += work.iloc[:, norm_main[nn]].apply(clean_money_val).sum()
                        found_wages = True
                        break
                if not found_wages:
                    for off in (-1, 1, -2, 2):
                        j = idx + off
                        if 0 <= j < len(df_adp.columns):
                            ch = str(df_adp.columns[j]).lower()
                            if "wages" in ch and "tax" not in ch:
                                total_w += work.iloc[:, j].apply(clean_money_val).sum()
                                break
            elif n in norm_top:
                start_idx = norm_top[n]
                end_i = len(df_adp.columns)
                for j in range(start_idx + 1, len(adp_top)):
                    if pd.notna(adp_top[j]) and str(adp_top[j]).strip() != "":
                        end_i = j
                        break
                for k in range(start_idx, end_i):
                    ch = str(df_adp.columns[k]).strip().lower()
                    if "taxable wages" in ch:
                        total_w += work.iloc[:, k].apply(clean_money_val).sum()
                    elif side == "EE" and (ch == "ee amount" or ch.startswith("ee amount.")):
                        total_a += work.iloc[:, k].apply(clean_money_val).sum()
                    elif side == "ER" and (ch == "er amount" or ch.startswith("er amount.")):
                        total_a += work.iloc[:, k].apply(clean_money_val).sum()
    return total_w, total_a


def compute_tax_rate_verification(df_uzio, uzio_top, adp_data_list, mappings):
    """Build the tax-rate verification table (SS, Medicare, FUTA, SUTA per state)."""
    uzio_to_adp = {}
    for m in mappings:
        if m.get("Category") == "Taxes":
            uzio_to_adp.setdefault(m["UZIO_Name"], []).append(m["ADP_Name"])

    targets = [
        ("Social Security", "EE", "Social Security Tax",          STANDARD_TAX_RATES["Social Security EE"]),
        ("Social Security", "ER", "Employer Social Security Tax", STANDARD_TAX_RATES["Social Security ER"]),
        ("Medicare",        "EE", "Medicare Tax",                 STANDARD_TAX_RATES["Medicare EE"]),
        ("Medicare",        "ER", "Employer Medicare Tax",        STANDARD_TAX_RATES["Medicare ER"]),
        ("FUTA",            "ER", "Federal Unemployment Tax",     STANDARD_TAX_RATES["FUTA ER"]),
    ]
    if uzio_top:
        suta_re = re.compile(r"^\s*([A-Z]{2})\s+STATE\s+UNEMPLOYMENT\s+TAX\s*$", re.I)
        for h in uzio_top:
            if pd.notna(h):
                m = suta_re.match(str(h))
                if m:
                    targets.append((f"SUTA - {m.group(1).upper()}", "ER", str(h).strip(), None))

    rows = []
    for tax, side, uzio_name, std in targets:
        u_w, u_a = _sum_uzio_section(df_uzio, uzio_top, uzio_name, side)
        a_w, a_a = _sum_adp_for_uzio_name(adp_data_list, uzio_to_adp.get(uzio_name, []), side)
        u_rate = (u_a / u_w * 100) if u_w > 0 else None
        a_rate = (a_a / a_w * 100) if a_w > 0 else None

        if std is None:
            status = "Info (Employer-set)"
            std_disp = "Employer-set"
        else:
            off_u = (u_rate is not None) and abs(u_rate - std) > RATE_TOLERANCE_PCT
            off_a = (a_rate is not None) and abs(a_rate - std) > RATE_TOLERANCE_PCT
            status = "Mismatch" if (off_u or off_a) else "Match"
            std_disp = f"{std:.2f}%"

        rows.append({
            "Tax": tax,
            "Side": side,
            "ADP Taxable Wages":  round(a_w, 2),
            "ADP Amount":         round(a_a, 2),
            "ADP Effective Rate": (f"{a_rate:.4f}%" if a_rate is not None else "-"),
            "UZIO Taxable Wages": round(u_w, 2),
            "UZIO Amount":        round(u_a, 2),
            "UZIO Effective Rate":(f"{u_rate:.4f}%" if u_rate is not None else "-"),
            "Standard Rate":      std_disp,
            "Status":             status,
        })
    return pd.DataFrame(rows)


def run_comparison(adp_files, uzio_file, mappings):
    """Main logic to compare totals based on mappings."""
    try:
        df_uzio, uzio_top, uzio_sheet = find_header_and_data(uzio_file)
        adp_data_list = []
        for adp_file in adp_files:
            df_adp, adp_top, adp_sheet = find_header_and_data(adp_file)
            adp_data_list.append((df_adp, adp_top, adp_sheet))
    except Exception as e:
        return None, f"Error reading payroll files: {e}", None, None, None

    results = []
    employee_mismatches = []
    
    unique_uzio_items = {}
    for m in mappings:
        u_name = m["UZIO_Name"]
        if u_name not in unique_uzio_items:
            unique_uzio_items[u_name] = {"Category": m["Category"], "ADP_Names": []}
        unique_uzio_items[u_name]["ADP_Names"].append(m["ADP_Name"])

    for u_name, data in unique_uzio_items.items():
        cat = data["Category"]
        adp_names = data["ADP_Names"]
        
        adp_total = 0.0
        adp_cols = []
        adp_emp_detail = {}
        adp_emp_counts = {}
        for df_a, adp_t, _ in adp_data_list:
            tot, cols, emp_m, emp_c = calculate_totals(df_a, adp_t, adp_names)
            adp_total += tot
            for c in cols:
                if c not in adp_cols:
                    adp_cols.append(c)
            for (eid, p_date), v in emp_m.items():
                if eid not in adp_emp_detail: adp_emp_detail[eid] = {}
                adp_emp_detail[eid][p_date] = adp_emp_detail[eid].get(p_date, 0.0) + v
            for (eid, p_date), c_val in emp_c.items():
                if eid not in adp_emp_counts: adp_emp_counts[eid] = {}
                adp_emp_counts[eid][p_date] = adp_emp_counts[eid].get(p_date, 0) + c_val
        
        uzio_total, uzio_cols, uzio_emp_m, _ = calculate_totals(df_uzio, uzio_top, [u_name])
        uzio_emp_detail = {}
        for (eid, p_date), v in uzio_emp_m.items():
            if eid not in uzio_emp_detail: uzio_emp_detail[eid] = {}
            uzio_emp_detail[eid][p_date] = uzio_emp_detail[eid].get(p_date, 0.0) + v
        
        diff = uzio_total - adp_total
        status = "Match" if abs(diff) <= 0.02 else "Mismatch"
        
        results.append({
            "Category": cat,
            "UZIO Item": u_name,
            "ADP Total": round(adp_total, 2),
            "UZIO Total": round(uzio_total, 2),
            "Difference": round(diff, 2),
            "Status": status,
            "ADP Columns Found": ", ".join(adp_cols) if adp_cols else "None",
            "UZIO Columns Found": ", ".join(uzio_cols) if uzio_cols else "None"
        })
        
        if status == "Mismatch":
            all_emp_ids = set(adp_emp_detail.keys()).union(set(uzio_emp_detail.keys()))
            for eid in all_emp_ids:
                if eid == "Unknown": continue
                
                emp_adp_total = sum(adp_emp_detail.get(eid, {}).values())
                emp_uzio_total = sum(uzio_emp_detail.get(eid, {}).values())
                
                if abs(emp_uzio_total - emp_adp_total) > 0.02:
                    adp_dates = adp_emp_detail.get(eid, {})
                    uzio_dates = uzio_emp_detail.get(eid, {})
                    all_dates = set(adp_dates.keys()).union(set(uzio_dates.keys()))
                    
                    for p_date in all_dates:
                        val_adp = adp_dates.get(p_date, 0.0)
                        val_uzio = uzio_dates.get(p_date, 0.0)
                        date_diff = val_uzio - val_adp
                        
                        if abs(date_diff) > 0.02:
                            multiple_entries = "Yes" if adp_emp_counts.get(eid, {}).get(p_date, 0) > 1 else "No"
                            employee_mismatches.append({
                                "Associate ID": eid,
                                "Pay Date": p_date,
                                "Category": cat,
                                "UZIO Item": u_name,
                                "ADP Amount": round(val_adp, 2),
                                "UZIO Amount": round(val_uzio, 2),
                                "Difference": round(date_diff, 2),
                                "Multiple ADP Entries on Same Date": multiple_entries
                            })

    df_results = pd.DataFrame(results)
    df_emp_mismatches = pd.DataFrame(employee_mismatches)

    # Three additional analyses on the loaded data
    df_dups        = detect_duplicate_pay_periods(df_uzio)
    df_stub_counts = compute_pay_stub_count_diff(adp_data_list, df_uzio)
    df_tax_rates   = compute_tax_rate_verification(df_uzio, uzio_top, adp_data_list, mappings)

    out_buffer = io.BytesIO()
    with pd.ExcelWriter(out_buffer, engine='xlsxwriter') as writer:
        wb = writer.book
        red_fill   = wb.add_format({"bg_color": "#FFE5E5", "font_color": "#9C0006"})
        green_fill = wb.add_format({"bg_color": "#E5F5E5", "font_color": "#006100"})

        df_results.to_excel(writer, sheet_name="Full Comparison", index=False)
        df_mismatches = df_results[df_results["Status"] == "Mismatch"][["Category", "UZIO Item", "ADP Columns Found", "UZIO Columns Found", "ADP Total", "UZIO Total", "Difference"]]
        df_mismatches.to_excel(writer, sheet_name="Mismatches Only", index=False)

        sheet_names = ["Full Comparison", "Mismatches Only"]
        dfs_to_format = [df_results, df_mismatches]
        if not df_emp_mismatches.empty:
            df_emp_mismatches.to_excel(writer, sheet_name="Employee Mismatches", index=False)
            sheet_names.append("Employee Mismatches")
            dfs_to_format.append(df_emp_mismatches)

        # Duplicate Pay Periods (UZIO file)
        if df_dups.empty:
            placeholder = pd.DataFrame({"Result": ["No duplicate pay-period entries detected in the UZIO file."]})
            placeholder.to_excel(writer, sheet_name="Duplicate Pay Periods", index=False)
            sheet_names.append("Duplicate Pay Periods"); dfs_to_format.append(placeholder)
        else:
            df_dups.to_excel(writer, sheet_name="Duplicate Pay Periods", index=False)
            sheet_names.append("Duplicate Pay Periods"); dfs_to_format.append(df_dups)

        # Pay Stub Counts (ADP combined vs UZIO)
        if df_stub_counts.empty:
            placeholder = pd.DataFrame({"Result": ["Could not compute pay-stub counts (missing ID or Pay Date column)."]})
            placeholder.to_excel(writer, sheet_name="Pay Stub Counts", index=False)
            sheet_names.append("Pay Stub Counts"); dfs_to_format.append(placeholder)
        else:
            df_stub_counts.to_excel(writer, sheet_name="Pay Stub Counts", index=False)
            sheet_names.append("Pay Stub Counts"); dfs_to_format.append(df_stub_counts)
            sh = writer.sheets["Pay Stub Counts"]
            n_rows = len(df_stub_counts)
            status_col = list(df_stub_counts.columns).index("Status")
            sh.conditional_format(1, 0, n_rows, len(df_stub_counts.columns) - 1, {
                "type": "formula",
                "criteria": f'=INDIRECT(ADDRESS(ROW(),{status_col + 1}))<>"Match"',
                "format": red_fill,
            })

        # Tax Rate Verification (SS / Medicare / FUTA / per-state SUTA)
        if df_tax_rates.empty:
            placeholder = pd.DataFrame({"Result": ["Could not build tax rate verification (no tax sections detected)."]})
            placeholder.to_excel(writer, sheet_name="Tax Rate Verification", index=False)
            sheet_names.append("Tax Rate Verification"); dfs_to_format.append(placeholder)
        else:
            df_tax_rates.to_excel(writer, sheet_name="Tax Rate Verification", index=False)
            sheet_names.append("Tax Rate Verification"); dfs_to_format.append(df_tax_rates)
            sh = writer.sheets["Tax Rate Verification"]
            n_rows = len(df_tax_rates)
            status_col = list(df_tax_rates.columns).index("Status")
            sh.conditional_format(1, 0, n_rows, len(df_tax_rates.columns) - 1, {
                "type": "formula",
                "criteria": f'=INDIRECT(ADDRESS(ROW(),{status_col + 1}))="Mismatch"',
                "format": red_fill,
            })
            sh.conditional_format(1, 0, n_rows, len(df_tax_rates.columns) - 1, {
                "type": "formula",
                "criteria": f'=INDIRECT(ADDRESS(ROW(),{status_col + 1}))="Match"',
                "format": green_fill,
            })

        for sheet_name, curr_df in zip(sheet_names, dfs_to_format):
            if curr_df.empty:
                continue
            sheet = writer.sheets[sheet_name]
            for i, col in enumerate(curr_df.columns):
                column_len = max(curr_df[col].astype(str).map(len).max(), len(col)) + 2
                sheet.set_column(i, i, min(column_len, 50))

    return df_results, out_buffer.getvalue(), df_dups, df_stub_counts, df_tax_rates

# ---------------------------------------------------------------------------
# Auto-detect helper for bulk upload
# ---------------------------------------------------------------------------
def auto_detect_files(uploaded_files):
    """
    Given a list of uploaded files, auto-detect each file's role by inspecting
    column headers. Returns a dict with keys:
      'adp', 'uzio', 'earn', 'ded', 'cont', 'tax', 'unknown'
    """
    result = {
        'adp': [], 'uzio': None,
        'earn': None, 'ded': None, 'cont': None, 'tax': None,
        'unknown': []
    }

    for f in uploaded_files:
        f.seek(0)
        try:
            name = f.name.lower()
            if name.endswith('.csv'):
                df_peek = pd.read_csv(f, nrows=3, dtype=str)
                cols = [str(c).lower().strip() for c in df_peek.columns]
            else:
                # Excel: check all sheets, skip 'criteria' metadata sheets
                xls = pd.ExcelFile(f)
                cols = []
                for sheet in xls.sheet_names:
                    if 'criteria' in sheet.lower():
                        continue
                    df_peek = pd.read_excel(xls, sheet_name=sheet, header=None, nrows=5, dtype=str)
                    # Try to find the actual header row
                    for i, row in df_peek.iterrows():
                        row_vals = [str(v).lower().strip() for v in row if pd.notna(v) and str(v).strip()]
                        if any(x in ' '.join(row_vals) for x in ['employee id', 'associate id', 'source earning', 'source deduction', 'source tax']):
                            cols = row_vals
                            break
                    if cols:
                        break
                if not cols:
                    # Fallback: just use column headers of first non-criteria sheet
                    for sheet in xls.sheet_names:
                        if 'criteria' not in sheet.lower():
                            df_peek = pd.read_excel(xls, sheet_name=sheet, nrows=3, dtype=str)
                            cols = [str(c).lower().strip() for c in df_peek.columns]
                            break
            f.seek(0)
        except Exception:
            f.seek(0)
            result['unknown'].append(f)
            continue

        col_str = " | ".join(cols)

        # --- Mapping files (most specific, check first) ---
        if 'source tax code name' in col_str:
            result['tax'] = f
        elif 'source earning code name' in col_str:
            result['earn'] = f
        elif 'source deduction code name' in col_str:
            result['ded'] = f
        elif 'source contribution code name' in col_str:
            result['cont'] = f

        # --- Uzio Prior Payroll Register ---
        # Has 'employee id' + payroll data columns like 'regular wage', 'gross pay', 'pay date'
        elif 'employee id' in col_str and any(x in col_str for x in [
            'regular wage', 'gross pay', 'overtime', 'pay date', 'first name'
        ]) and 'associate id' not in col_str:
            result['uzio'] = f

        # --- ADP Prior Payroll file ---
        # Must have ASSOCIATE ID *and* a payroll-specific column (not just a lookup/deduction file)
        elif 'associate id' in col_str and any(x in col_str for x in [
            'regular earnings', 'gross pay', 'regular hours', 'total earnings', 'net pay'
        ]):
            result['adp'].append(f)

        else:
            result['unknown'].append(f)

    return result


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
def render_ui():
    st.title("ADP - Prior Payroll Audit Tool")
    st.markdown("""
    Compares the totals of payroll elements (Earnings, Deductions, Contributions, Taxes)
    between ADP and UZIO reports based on provided mapping files.
    """)

    st.info(
        "📋 **Reminder — Check the Uzio Payroll Allocation Report for negative taxable wages**\n\n"
        "Always download the **Uzio Payroll Allocation Report** and verify that **no employee "
        "has negative taxable wages**. If an employee shows negative taxable wages, that "
        "paystub must be **refreshed** before you rely on these totals. These negative entries "
        "are easy to identify — they appear as the **red-marked entries** in the Payroll "
        "Allocation Report."
    )

    # ── Upload mode toggle ──────────────────────────────────────────────────
    upload_mode = st.radio(
        "Upload Mode",
        ["📦 Bulk Upload (select all files at once)", "🗂️ Manual Upload (file by file)"],
        horizontal=True,
        key="tc_upload_mode"
    )

    adp_files = []
    uzio_file = earn_file = cont_file = ded_file = tax_file = None

    # ── BULK MODE ────────────────────────────────────────────────────────────
    if upload_mode.startswith("📦"):
        st.info(
            "Select **all** your files at once — ADP payroll file(s), UZIO register, "
            "and all 4 mapping files. The tool will automatically classify each file.",
            icon="💡"
        )
        bulk_files = st.file_uploader(
            "Drop all files here",
            type=["xlsx", "xls", "csv"],
            accept_multiple_files=True,
            key="tc_bulk"
        )
        if bulk_files:
            detected = auto_detect_files(bulk_files)
            adp_files = detected['adp']
            uzio_file = detected['uzio']
            earn_file = detected['earn']
            ded_file  = detected['ded']
            cont_file = detected['cont']
            tax_file  = detected['tax']

            st.markdown("#### Auto-detected file roles")
            summary_rows = []
            for f in adp_files:
                summary_rows.append({"File": f.name, "Detected As": "ADP Payroll"})
            if uzio_file:
                summary_rows.append({"File": uzio_file.name, "Detected As": "UZIO Register"})
            if earn_file:
                summary_rows.append({"File": earn_file.name, "Detected As": "Earnings Mapping"})
            if ded_file:
                summary_rows.append({"File": ded_file.name, "Detected As": "Deductions Mapping"})
            if cont_file:
                summary_rows.append({"File": cont_file.name, "Detected As": "Contributions Mapping"})
            if tax_file:
                summary_rows.append({"File": tax_file.name, "Detected As": "Taxes Mapping"})
            for f in detected['unknown']:
                summary_rows.append({"File": f.name, "Detected As": "Unknown - not used"})

            if summary_rows:
                role_df = pd.DataFrame(summary_rows)
                # Colour the Detected As column
                def _colour_role(val):
                    if "ADP" in val or "UZIO" in val or "Mapping" in val:
                        return "color: green"
                    return "color: orange"
                st.dataframe(
                    role_df.style.map(_colour_role, subset=["Detected As"]),
                    use_container_width=True, hide_index=True
                )

            missing = []
            if not adp_files: missing.append("ADP Payroll file")
            if not uzio_file: missing.append("UZIO Register")
            if not earn_file: missing.append("Earnings Mapping")
            if not ded_file:  missing.append("Deductions Mapping")
            if not cont_file: missing.append("Contributions Mapping")
            if not tax_file:  missing.append("Taxes Mapping")
            if missing:
                st.warning(f"Still missing: **{', '.join(missing)}**. Please add these files too.")

    # ── MANUAL MODE ──────────────────────────────────────────────────────────
    else:
        with st.expander("📁 Upload Payroll Reports", expanded=True):
            col1, col2 = st.columns(2)
            with col1:
                adp_files = st.file_uploader(
                    "Upload ADP Prior Payroll File(s)",
                    type=["xlsx", "xls", "csv"],
                    accept_multiple_files=True,
                    key="tc_adp"
                )
            with col2:
                uzio_file = st.file_uploader(
                    "Upload UZIO Prior Payroll Register",
                    type=["xlsx", "xls", "csv"],
                    key="tc_uzio"
                )

        with st.expander("🗺️ Upload Mapping Files", expanded=True):
            m_col1, m_col2 = st.columns(2)
            with m_col1:
                earn_file = st.file_uploader("Earnings Mapping File",      type=["xlsx", "xls", "csv"], key="tc_m_earn")
                cont_file = st.file_uploader("Contributions Mapping File", type=["xlsx", "xls", "csv"], key="tc_m_cont")
            with m_col2:
                ded_file  = st.file_uploader("Deductions Mapping File",    type=["xlsx", "xls", "csv"], key="tc_m_ded")
                tax_file  = st.file_uploader("Taxes Mapping File",         type=["xlsx", "xls", "csv"], key="tc_m_tax")

    # ── RUN AUDIT ────────────────────────────────────────────────────────────
    if "audit_results" not in st.session_state:
        st.session_state.audit_results = None
    if "audit_report" not in st.session_state:
        st.session_state.audit_report = None
    if "audit_dups" not in st.session_state:
        st.session_state.audit_dups = None
    if "audit_stub_counts" not in st.session_state:
        st.session_state.audit_stub_counts = None
    if "audit_tax_rates" not in st.session_state:
        st.session_state.audit_tax_rates = None

    all_ready = (
        adp_files and len(adp_files) > 0 and
        all([uzio_file, earn_file, cont_file, ded_file, tax_file])
    )

    if all_ready:
        if st.button("Run Total Comparison", type="primary", use_container_width=True):
            with st.spinner("Processing files and calculating totals..."):
                all_mappings = []
                all_mappings.extend(load_mapping(earn_file, "Earnings",      "Source Earning Code Name",      "Uzio Earning Code Name"))
                all_mappings.extend(load_mapping(ded_file,  "Deductions",    "Source Deduction Code Name",    "Uzio Deduction Code Name"))
                all_mappings.extend(load_mapping(cont_file, "Contributions", "Source Contribution Code Name", "Uzio Contribution Code Name"))
                all_mappings.extend(load_mapping(tax_file,  "Taxes",         "Source Tax Code Name",          "Uzio Tax Code Description"))

                if not all_mappings:
                    st.error("No mappings could be loaded. Please check the mapping file column headers.")
                    return

                res_df, report_data, dup_df, stub_df, tax_df = run_comparison(adp_files, uzio_file, all_mappings)
                if res_df is not None:
                    st.session_state.audit_results     = res_df
                    st.session_state.audit_report      = report_data
                    st.session_state.audit_dups        = dup_df
                    st.session_state.audit_stub_counts = stub_df
                    st.session_state.audit_tax_rates   = tax_df
                else:
                    st.error(f"Failed to generate results. Error: {report_data}")

    if st.session_state.audit_results is not None:
        results_df  = st.session_state.audit_results
        report_data = st.session_state.audit_report

        st.success("Comparison completed!")

        matches    = len(results_df[results_df["Status"] == "Match"])
        mismatches = len(results_df[results_df["Status"] == "Mismatch"])

        m1, m2, m3 = st.columns(3)
        m1.metric("Total Items", len(results_df))
        m2.metric("Matches",    matches)
        m3.metric("Mismatches", mismatches,
                  delta=mismatches if mismatches > 0 else None,
                  delta_color="inverse")

        st.subheader("Comparison Results")

        def color_status(val):
            return 'color: green' if val == 'Match' else 'color: red'

        st.dataframe(
            results_df.style.map(color_status, subset=['Status']),
            use_container_width=True
        )

        # Duplicate pay-period entries (UZIO file)
        dup_df = st.session_state.audit_dups
        if dup_df is not None and not dup_df.empty:
            affected_emps   = dup_df["Employee ID"].nunique()
            affected_groups = dup_df.groupby(["Employee ID", "Start Date", "End Date", "Pay Date"]).ngroups
            st.warning(
                f"Detected {affected_groups} duplicated pay-period group(s) across {affected_emps} employee(s) "
                "in the UZIO Prior Payroll Register. See the **'Duplicate Pay Periods'** tab in the report."
            )
            with st.expander(f"View duplicate pay-period entries ({len(dup_df)} rows)", expanded=False):
                st.dataframe(dup_df, use_container_width=True)

        # Pay-stub count comparison
        stub_df = st.session_state.audit_stub_counts
        if stub_df is not None and not stub_df.empty:
            mismatched = stub_df[stub_df["Status"] != "Match"]
            if not mismatched.empty:
                st.warning(
                    f"Found {len(mismatched)} employee(s) with a pay-stub-count mismatch between ADP and UZIO. "
                    "See the **'Pay Stub Counts'** tab in the report."
                )
            with st.expander(f"View pay-stub count comparison ({len(stub_df)} employees)", expanded=False):
                def color_stub(val):
                    return 'background-color: #FFE5E5' if val != "Match" else ''
                st.dataframe(
                    stub_df.style.map(color_stub, subset=["Status"]),
                    use_container_width=True,
                )

        # Tax-rate verification
        tax_df = st.session_state.audit_tax_rates
        if tax_df is not None and not tax_df.empty:
            mismatched = tax_df[tax_df["Status"] == "Mismatch"]
            if not mismatched.empty:
                st.warning(
                    f"Found {len(mismatched)} tax line(s) where the effective rate differs from the standard rate "
                    "by more than 0.05%. See the **'Tax Rate Verification'** tab in the report."
                )
            with st.expander("View tax rate verification (SS / Medicare / FUTA / SUTA)", expanded=False):
                def color_tax(val):
                    if val == "Mismatch": return 'background-color: #FFE5E5'
                    if val == "Match":    return 'background-color: #E5F5E5'
                    return ''
                st.dataframe(
                    tax_df.style.map(color_tax, subset=["Status"]),
                    use_container_width=True,
                )

        st.download_button(
            label="Download Full Comparison Report",
            data=report_data,
            file_name="Prior payroll audit report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="tc_download_v2",
            use_container_width=True
        )

if __name__ == "__main__":
    st.set_page_config(page_title="Total Comparison Tool", layout="wide")
    render_ui()
