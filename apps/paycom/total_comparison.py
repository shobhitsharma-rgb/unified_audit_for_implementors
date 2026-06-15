import streamlit as st
import pandas as pd
import io
import re
import os
from utils.audit_utils import clean_money_val, norm_colname

def load_mapping(file, cat_name, source_col, uzio_col):
    """Load a mapping file and return a list of mappings (Source_Name, UZIO_Name)."""
    try:
        file.seek(0)
        if str(file.name).lower().endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)
        # Normalize headers to find columns
        df.columns = [norm_colname(c) for c in df.columns]
        
        # Finding the actual column names in the sheet
        actual_source_col = next((c for c in df.columns if source_col.lower() in c.lower()), None)
        actual_uzio_col = next((c for c in df.columns if uzio_col.lower() in c.lower()), None)
        
        if not actual_source_col or not actual_uzio_col:
            st.warning(f"Could not find exact columns in {cat_name} mapping. Looking for '{source_col}' and '{uzio_col}'. Available: {list(df.columns)}")
            return []
            
        mappings = []
        for _, row in df.iterrows():
            s_val = str(row[actual_source_col]).strip()
            u_val = str(row[actual_uzio_col]).strip()
            if s_val and u_val and s_val.lower() != 'nan' and u_val.lower() != 'nan':
                mappings.append({
                    "Category": cat_name,
                    "Source_Name": s_val,
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

def parse_paycom_filename_date(filename):
    """Extract the third date from Paycom filename: Priorpayroll_MMDDYYYY_MMDDYYYY_MMDDYYYY.xlsx"""
    # Look for any sequence of 8 digits that might be a date
    match = re.findall(r'(\d{8})', filename)
    if len(match) >= 3:
        d_str = match[2] # Third date
        try:
            return f"{d_str[4:]}-{d_str[:2]}-{d_str[2:4]}" # YYYY-MM-DD
        except:
            return "Unknown"
    # Fallback for 2-date pattern
    if len(match) >= 2:
        d_str = match[1]
        try:
            return f"{d_str[4:]}-{d_str[:2]}-{d_str[2:4]}"
        except:
            return "Unknown"
    # Fallback for 1-date pattern
    if len(match) >= 1:
        d_str = match[0]
        try:
            return f"{d_str[4:]}-{d_str[:2]}-{d_str[2:4]}"
        except:
            return "Unknown"
    return "Unknown"

def find_header_and_data_uzio(file):
    """Specific logic for Uzio reports (often multi-row headers)."""
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
            
        return df, header_top, "Sheet1"

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
        
    return df, header_top, target_sheet

def find_header_and_data_paycom(file):
    """Specific logic for Paycom reports."""
    file.seek(0)
    if str(file.name).lower().endswith('.csv'):
        file.seek(0)
        df_peek = pd.read_csv(file, header=None, nrows=20)
        header_idx = 0
        for i, row in df_peek.iterrows():
            row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
            if any(kw in row_str for kw in ["ee code", "description", "earning", "amount", "row labels"]):
                header_idx = i
                break
        file.seek(0)
        df = pd.read_csv(file, header=header_idx)
        return df, None, "Sheet1"

    # Read first sheet
    xls = pd.ExcelFile(file)
    df_peek = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=None, nrows=10)
    
    # Try to find header row dynamically
    header_idx = 0
    for i, row in df_peek.iterrows():
        row_str = " ".join([str(x).lower() for x in row if pd.notna(x)])
        if any(kw in row_str for kw in ["ee code", "description", "earning", "amount", "row labels"]):
            header_idx = i
            break
            
    df = pd.read_excel(xls, sheet_name=xls.sheet_names[0], header=header_idx)
    return df, None, xls.sheet_names[0]

def calculate_totals_uzio(df, header_top, column_names):
    """Sum up values for Uzio columns (Wide format)."""
    found_cols = []
    emp_tots = {}
    
    # Header aliases for ID and Date
    id_aliases = ["employee id", "file #", "associate id", "ee code"]
    date_aliases = ["pay date", "check date", "period end"]
    
    id_col = next((c for c in df.columns if any(x in str(c).lower() for x in id_aliases)), None)
    date_col = next((c for c in df.columns if any(x in str(c).lower() for x in date_aliases)), None)
    
    if id_col:
        df_clean = df[df[id_col].notna()].copy()
        df_clean[id_col] = df_clean[id_col].apply(normalize_id)
        df_clean = df_clean[~df_clean[id_col].str.lower().str.contains("total|grand", na=False)]
    else:
        df_clean = df.copy()

    # Exact (case-insensitive) header lookup wins over the normalized lookup so
    # that 'Bonus' and 'Bonus (Hours)' don't collapse onto the same key — norm_colname
    # strips parens, which would otherwise hand 'Bonus' the 'Bonus (Hours)' column.
    exact_cols_main = {str(c).strip().lower(): i for i, c in enumerate(df.columns)}
    norm_cols_main = {}
    for i, c in enumerate(df.columns):
        norm_cols_main.setdefault(norm_colname(c).lower(), i)
    norm_cols_top = {}
    if header_top:
        for i, c in enumerate(header_top):
            if pd.notna(c) and str(c).strip() != "":
                norm_cols_top.setdefault(norm_colname(c).lower(), i)

    cols_to_sum = []
    for name in column_names:
        raw_name = str(name).strip().lower()
        n_name = norm_colname(name).lower()
        if raw_name in exact_cols_main:
            idx = exact_cols_main[raw_name]
            cols_to_sum.append(df.columns[idx])
            found_cols.append(df.columns[idx])
        elif n_name in norm_cols_main:
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
        eid = row[id_col] if id_col else "Summary"
        pay_date = format_pay_date(row[date_col]) if date_col else "Unknown"
        row_tot = sum(clean_money_val(row[c]) for c in set(cols_to_sum))
        key = (eid, pay_date)
        emp_tots[key] = emp_tots.get(key, 0.0) + row_tot
            
    return sum(emp_tots.values()), found_cols, emp_tots

def calculate_totals_paycom(df, mapping_source_names, filename, uzio_item_name=""):
    """Sum up values for Paycom (Long format)."""
    found_items = set()
    emp_tots = {}
    
    # Flexible column detection
    id_aliases = ["ee code", "employee code", "file #", "clock #", "associate id"]
    desc_aliases = ["type description", "description", "earning/deduction/tax", "code description", "row labels"]
    amt_aliases = ["current amount", "amount", "total amount", "value", "sum of amount"]
    
    id_col = next((c for c in df.columns if any(x in str(c).lower() for x in id_aliases)), None)
    desc_col = next((c for c in df.columns if any(x in str(c).lower() for x in desc_aliases)), None)
    code_desc_col = next((c for c in df.columns if "code description" in str(c).lower()), None)
    
    # Prefer 'Current Amount' over 'Total Amount' if both are in aliases
    amt_col = next((c for c in df.columns if "current amount" in str(c).lower()), None)
    if not amt_col:
        amt_col = next((c for c in df.columns if any(x in str(c).lower() for x in amt_aliases)), None)
    
    if not all([desc_col, amt_col]):
        return 0.0, [], {}

    pay_date = parse_paycom_filename_date(filename)
    norm_mappings = [n.lower().strip() for n in mapping_source_names]
    
    for _, row in df.iterrows():
        raw_desc = str(row[desc_col]).strip()
        val_desc = raw_desc.lower()
        
        # Exact match based on mapping file
        if val_desc in norm_mappings:
            if code_desc_col and pd.notna(row[code_desc_col]):
                if str(row[code_desc_col]).strip().lower() == "employee benefits":
                    continue
            # Differentiate Employee vs Employer for Social Security and Medicare
            if "medicare" in val_desc or "social security" in val_desc or "ssc" in val_desc:
                if code_desc_col and pd.notna(row[code_desc_col]):
                    code_desc_val = str(row[code_desc_col]).strip().lower()
                    is_employer_tax = "employer" in uzio_item_name.lower() or "er " in uzio_item_name.lower()
                    
                    if is_employer_tax and "client side" not in code_desc_val:
                        continue
                    if not is_employer_tax and "w/h" not in code_desc_val:
                        continue

            eid = normalize_id(row[id_col]) if id_col else "Summary"
            amount = clean_money_val(row[amt_col])
            
            key = (eid, pay_date)
            emp_tots[key] = emp_tots.get(key, 0.0) + amount
            found_items.add(raw_desc)
            
    return sum(emp_tots.values()), list(found_items), emp_tots

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
    if not eid_col:
        return df
    work = df[df[eid_col].notna()].copy()
    work[eid_col] = work[eid_col].astype(str).str.strip()
    return work[(work[eid_col] != "") & (~work[eid_col].str.lower().str.contains("total|grand", na=False))]


def _sum_uzio_section(df, header_top, section_name, side):
    """Sum Taxable Wages and EE/ER Amount within a UZIO section header."""
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


def _sum_paycom_for_uzio_name(paycom_data_list, source_names):
    """Best-effort sum of (taxable wages, amount) on Paycom long-format rows.
    Wages are inferred from rows whose Description matches the tax description with 'tax'->'wages',
    or contains 'taxable wages' / 'wages' for the same tax."""
    if not source_names:
        return 0.0, 0.0
    desc_aliases = ["type description", "description", "earning/deduction/tax", "code description", "row labels"]
    norm_targets = [n.lower().strip() for n in source_names]
    # Build wage-side description candidates
    wage_targets = set()
    for n in norm_targets:
        wage_targets.add(re.sub(r"\btax\b", "wages", n, flags=re.I))
        wage_targets.add(n + " wages")
        wage_targets.add(n + " taxable wages")

    total_w = total_a = 0.0
    for df_p, _ in paycom_data_list:
        desc_col = next((c for c in df_p.columns if any(x in str(c).lower() for x in desc_aliases)), None)
        amt_col  = next((c for c in df_p.columns if "current amount" in str(c).lower()), None)
        if not amt_col:
            amt_col = next((c for c in df_p.columns if any(x in str(c).lower() for x in ["amount", "total amount", "value", "sum of amount"])), None)
        if not desc_col or not amt_col:
            continue
        for _, row in df_p.iterrows():
            d = str(row[desc_col]).strip().lower()
            if d in norm_targets:
                total_a += clean_money_val(row[amt_col])
            elif d in wage_targets:
                total_w += clean_money_val(row[amt_col])
    return total_w, total_a


def compute_tax_rate_verification(df_uzio, uzio_top, paycom_data_list, mappings):
    """Build the tax-rate verification table (SS, Medicare, FUTA, SUTA per state) for Paycom."""
    uzio_to_source = {}
    for m in mappings:
        if m.get("Category") == "Taxes":
            uzio_to_source.setdefault(m["UZIO_Name"], []).append(m["Source_Name"])

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
        p_w, p_a = _sum_paycom_for_uzio_name(paycom_data_list, uzio_to_source.get(uzio_name, []))
        u_rate = (u_a / u_w * 100) if u_w > 0 else None
        p_rate = (p_a / p_w * 100) if p_w > 0 else None

        if std is None:
            status = "Info (Employer-set)"
            std_disp = "Employer-set"
        else:
            off_u = (u_rate is not None) and abs(u_rate - std) > RATE_TOLERANCE_PCT
            off_p = (p_rate is not None) and abs(p_rate - std) > RATE_TOLERANCE_PCT
            status = "Mismatch" if (off_u or off_p) else "Match"
            std_disp = f"{std:.2f}%"

        rows.append({
            "Tax": tax,
            "Side": side,
            "Paycom Taxable Wages":  round(p_w, 2),
            "Paycom Amount":         round(p_a, 2),
            "Paycom Effective Rate": (f"{p_rate:.4f}%" if p_rate is not None else "-"),
            "UZIO Taxable Wages":    round(u_w, 2),
            "UZIO Amount":           round(u_a, 2),
            "UZIO Effective Rate":   (f"{u_rate:.4f}%" if u_rate is not None else "-"),
            "Standard Rate":         std_disp,
            "Status":                status,
        })
    return pd.DataFrame(rows)


def run_comparison(paycom_files, uzio_file, mappings):
    """Main logic to compare totals based on mappings."""
    try:
        df_uzio, uzio_top, _ = find_header_and_data_uzio(uzio_file)
        paycom_data_list = []
        for p_file in paycom_files:
            df_p, _, _ = find_header_and_data_paycom(p_file)
            paycom_data_list.append((df_p, p_file.name))
    except Exception as e:
        return None, None, f"Error reading payroll files: {e}", None

    results = []
    employee_mismatches = []
    
    unique_uzio_items = {}
    for m in mappings:
        u_name = m["UZIO_Name"]
        if u_name not in unique_uzio_items:
            unique_uzio_items[u_name] = {"Category": m["Category"], "Source_Names": []}
        unique_uzio_items[u_name]["Source_Names"].append(m["Source_Name"])

    for u_name, data in unique_uzio_items.items():
        cat = data["Category"]
        source_names = data["Source_Names"]
        
        paycom_total = 0.0
        paycom_items_found = []
        paycom_emp_detail = {} # (eid) -> {date: amount}
        
        for df_p, fname in paycom_data_list:
            tot, found, emp_m = calculate_totals_paycom(df_p, source_names, fname, u_name)
            paycom_total += tot
            for f in found:
                if f not in paycom_items_found: paycom_items_found.append(f)
            for (eid, p_date), v in emp_m.items():
                if eid not in paycom_emp_detail: paycom_emp_detail[eid] = {}
                paycom_emp_detail[eid][p_date] = paycom_emp_detail[eid].get(p_date, 0.0) + v
        
        uzio_total, uzio_cols, uzio_emp_m = calculate_totals_uzio(df_uzio, uzio_top, [u_name])
        uzio_emp_detail = {}
        for (eid, p_date), v in uzio_emp_m.items():
            if eid not in uzio_emp_detail: uzio_emp_detail[eid] = {}
            uzio_emp_detail[eid][p_date] = uzio_emp_detail[eid].get(p_date, 0.0) + v
        
        diff = uzio_total - paycom_total
        status = "Match" if abs(diff) <= 0.02 else "Mismatch"
        
        results.append({
            "Category": cat,
            "UZIO Item": u_name,
            "Paycom Total": round(paycom_total, 2),
            "UZIO Total": round(uzio_total, 2),
            "Difference": round(diff, 2),
            "Status": status,
            "Paycom Items Found": ", ".join(paycom_items_found) if paycom_items_found else "None",
            "UZIO Columns Found": ", ".join(uzio_cols) if uzio_cols else "None"
        })
        
        if status == "Mismatch":
            all_emp_ids = set(paycom_emp_detail.keys()).union(set(uzio_emp_detail.keys()))
            for eid in all_emp_ids:
                if eid == "Unknown": continue
                
                emp_p_total = sum(paycom_emp_detail.get(eid, {}).values())
                emp_u_total = sum(uzio_emp_detail.get(eid, {}).values())
                
                if abs(emp_u_total - emp_p_total) > 0.02:
                    p_dates = paycom_emp_detail.get(eid, {})
                    u_dates = uzio_emp_detail.get(eid, {})
                    all_dates = set(p_dates.keys()).union(set(u_dates.keys()))
                    
                    for p_date in all_dates:
                        val_p = p_dates.get(p_date, 0.0)
                        val_u = u_dates.get(p_date, 0.0)
                        date_diff = val_u - val_p
                        
                        if abs(date_diff) > 0.02:
                            employee_mismatches.append({
                                "Associate ID": eid,
                                "Pay Date": p_date,
                                "Category": cat,
                                "UZIO Item": u_name,
                                "Paycom Amount": round(val_p, 2),
                                "UZIO Amount": round(val_u, 2),
                                "Difference": round(date_diff, 2)
                            })

    df_results = pd.DataFrame(results)
    df_emp_mismatches = pd.DataFrame(employee_mismatches)

    # Tax-rate verification (SS / Medicare / FUTA / SUTA per state)
    df_tax_rates = compute_tax_rate_verification(df_uzio, uzio_top, paycom_data_list, mappings)

    out_buffer = io.BytesIO()
    with pd.ExcelWriter(out_buffer, engine='xlsxwriter') as writer:
        wb = writer.book
        red_fill   = wb.add_format({"bg_color": "#FFE5E5", "font_color": "#9C0006"})
        green_fill = wb.add_format({"bg_color": "#E5F5E5", "font_color": "#006100"})

        df_results.to_excel(writer, sheet_name="Full Comparison", index=False)
        df_mismatches = df_results[df_results["Status"] == "Mismatch"][["Category", "UZIO Item", "Paycom Items Found", "UZIO Columns Found", "Paycom Total", "UZIO Total", "Difference"]]
        df_mismatches.to_excel(writer, sheet_name="Mismatches Only", index=False)

        sheet_names = ["Full Comparison", "Mismatches Only"]
        dfs_to_format = [df_results, df_mismatches]
        if not df_emp_mismatches.empty:
            df_emp_mismatches.to_excel(writer, sheet_name="Employee Mismatches", index=False)
            sheet_names.append("Employee Mismatches")
            dfs_to_format.append(df_emp_mismatches)

        if df_tax_rates.empty:
            pd.DataFrame({"Result": ["Could not build tax rate verification (no tax sections detected)."]}).to_excel(
                writer, sheet_name="Tax Rate Verification", index=False
            )
        else:
            df_tax_rates.to_excel(writer, sheet_name="Tax Rate Verification", index=False)
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
        sheet_names.append("Tax Rate Verification")
        dfs_to_format.append(df_tax_rates if not df_tax_rates.empty else pd.DataFrame({"Result": ["Could not build tax rate verification."]}))

        for sheet_name, curr_df in zip(sheet_names, dfs_to_format):
            sheet = writer.sheets[sheet_name]
            for i, col in enumerate(curr_df.columns):
                column_len = max(curr_df[col].astype(str).map(len).max() if not curr_df.empty else 10, len(col)) + 2
                sheet.set_column(i, i, min(column_len, 50))

    return df_results, df_emp_mismatches, out_buffer.getvalue(), df_tax_rates

def render_ui():
    st.title("Paycom - Prior Payroll Audit Tool")
    st.markdown("""
    This tool compares the totals of payroll elements (Earnings, Deductions, Contributions, Taxes) 
    between Paycom and UZIO reports based on provided mapping files.
    
    **Required Files**:
    1.  **Paycom Prior Payroll File(s)** (Excel/CSV - format `Priorpayroll_MMDDYYYY_MMDDYYYY_MMDDYYYY.xlsx`)
    2.  **UZIO Prior Payroll Register File** (Excel/CSV)
    3.  **4 Mapping Files** (Earnings, Deductions, Contributions, Taxes)
    """)
    
    with st.expander("📁 Upload Payroll Reports", expanded=True):
        col1, col2 = st.columns(2)
        with col1:
            paycom_files = st.file_uploader(
                "Upload Paycom Prior Payroll File(s)",
                type=["xlsx", "xls", "csv"],
                accept_multiple_files=True,
                key="pc_tc_paycom",
                help="Select one or more Paycom files (e.g. Priorpayroll_MMDDYYYY_MMDDYYYY_MMDDYYYY.xlsx)"
            )
            if paycom_files:
                st.caption(f"✅ {len(paycom_files)} Paycom file(s) uploaded: {', '.join(f.name for f in paycom_files)}")
        with col2:
            uzio_file = st.file_uploader("Upload UZIO Prior Payroll Register", type=["xlsx", "xls", "csv"], key="pc_tc_uzio")

    with st.expander("🗺️ Upload Mapping Files", expanded=True):
        m_col1, m_col2 = st.columns(2)
        with m_col1:
            earn_file = st.file_uploader("Earnings Mapping File", type=["xlsx", "xls", "csv"], key="pc_tc_m_earn")
            cont_file = st.file_uploader("Contributions Mapping File", type=["xlsx", "xls", "csv"], key="pc_tc_m_cont")
        with m_col2:
            ded_file = st.file_uploader("Deductions Mapping File", type=["xlsx", "xls", "csv"], key="pc_tc_m_ded")
            tax_file = st.file_uploader("Taxes Mapping File", type=["xlsx", "xls", "csv"], key="pc_tc_m_tax")

    if "pc_audit_results" not in st.session_state:
        st.session_state.pc_audit_results = None
    if "pc_audit_emp_mismatches" not in st.session_state:
        st.session_state.pc_audit_emp_mismatches = None
    if "pc_audit_report" not in st.session_state:
        st.session_state.pc_audit_report = None
    if "pc_audit_tax_rates" not in st.session_state:
        st.session_state.pc_audit_tax_rates = None

    if paycom_files and len(paycom_files) > 0 and all([uzio_file, earn_file, cont_file, ded_file, tax_file]):
        if st.button("Run Total Comparison", type="primary"):
            with st.spinner("Processing files and calculating totals..."):
                all_mappings = []
                all_mappings.extend(load_mapping(earn_file, "Earnings", "Source Earning Code Name", "Uzio Earning Code Name"))
                all_mappings.extend(load_mapping(ded_file, "Deductions", "Source Deduction Code Name", "Uzio Deduction Code Name"))
                all_mappings.extend(load_mapping(cont_file, "Contributions", "Source Contribution Code Name", "Uzio Contribution Code Name"))
                all_mappings.extend(load_mapping(tax_file, "Taxes", "Source Tax Code Name", "Uzio Tax Code Description"))

                if not all_mappings:
                    st.error("No mappings could be loaded from the mapping files. Please check the column headers.")
                    return

                result = run_comparison(paycom_files, uzio_file, all_mappings)
                if result[0] is not None:
                    res_df, emp_mismatch_df, report_data, tax_df = result
                    st.session_state.pc_audit_results = res_df
                    st.session_state.pc_audit_emp_mismatches = emp_mismatch_df
                    st.session_state.pc_audit_report = report_data
                    st.session_state.pc_audit_tax_rates = tax_df
                else:
                    st.error(f"Failed to generate results. Error: {result[2] if len(result) > 2 else result[1]}")

        if st.session_state.pc_audit_results is not None:
            results_df = st.session_state.pc_audit_results
            emp_mismatch_df = st.session_state.pc_audit_emp_mismatches
            report_data = st.session_state.pc_audit_report
            
            st.success("Comparison completed!")
            matches = len(results_df[results_df["Status"] == "Match"])
            mismatches = len(results_df[results_df["Status"] == "Mismatch"])
            emp_mismatch_count = len(emp_mismatch_df["Associate ID"].unique()) if emp_mismatch_df is not None and not emp_mismatch_df.empty else 0
            
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Total Items", len(results_df))
            m2.metric("Matches", matches)
            m3.metric("Mismatches", mismatches, delta=mismatches if mismatches > 0 else None, delta_color="inverse")
            m4.metric("Employees with Discrepancies", emp_mismatch_count, delta=emp_mismatch_count if emp_mismatch_count > 0 else None, delta_color="inverse")
            
            st.subheader("Comparison Results")
            def color_status(val):
                return 'color: green' if val == 'Match' else 'color: red'
            st.dataframe(results_df.style.map(color_status, subset=['Status']), use_container_width=True)

            # --- Employee-level mismatch drill-down ---
            if emp_mismatch_df is not None and not emp_mismatch_df.empty:
                with st.expander(f"🔍 Employee-Level Mismatches ({emp_mismatch_count} employee(s) affected)", expanded=True):
                    st.markdown("""
                    The table below shows employees where Paycom and UZIO totals do **not** match,  
                    broken down by **Associate ID → UZIO Item → Pay Date** for easy identification.
                    """)
                    def color_diff(val):
                        try:
                            return 'color: red' if float(val) != 0 else 'color: green'
                        except:
                            return ''
                    st.dataframe(
                        emp_mismatch_df.sort_values(["Associate ID", "UZIO Item", "Pay Date"])
                                       .style.map(color_diff, subset=["Difference"]),
                        use_container_width=True
                    )
            else:
                st.info("✅ No employee-level discrepancies found across all pay periods.")

            # Tax-rate verification
            tax_df = st.session_state.pc_audit_tax_rates
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
                file_name=f"Paycom_prior_payroll_audit_report.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="pc_tc_download"
            )

if __name__ == "__main__":
    st.set_page_config(page_title="Paycom Total Comparison Tool", layout="wide")
    render_ui()
