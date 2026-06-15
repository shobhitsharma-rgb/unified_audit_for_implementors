
import io
import re
from datetime import datetime
from typing import Dict, List, Optional, Tuple

import pandas as pd
import yaml

def _norm_col(c) -> str:
    return str(c).strip().replace("\n", " ").strip() if c is not None else ""

def _read_any(uploaded_bytes: bytes, filename: str) -> pd.DataFrame:
    name = (filename or "").lower()
    bio = io.BytesIO(uploaded_bytes)
    if name.endswith(".csv"):
        return pd.read_csv(bio, dtype=str)
    if name.endswith(".txt"):
        try:
            return pd.read_csv(bio, dtype=str, sep="\t")
        except Exception:
            bio.seek(0)
            return pd.read_csv(bio, dtype=str)
    return pd.read_excel(bio, engine="openpyxl", dtype=str)

def _clean_blank(x):
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    s = str(x).strip()
    if s.lower() in {"nan", "none", "null"}:
        return ""
    return s

_BOOL_TRUE = {"true","t","yes","y","1","on"}
_BOOL_FALSE = {"false","f","no","n","0","off"}

def _normalize_bool(x: str) -> Optional[str]:
    s = _clean_blank(x).lower()
    if s == "":
        return None
    if s in _BOOL_TRUE:
        return "Yes"
    if s in _BOOL_FALSE:
        return "No"
    return None

def _to_float(s: str) -> Optional[float]:
    ss = _clean_blank(s)
    if ss == "":
        return None
    ss = ss.replace("$","").replace(",","")
    ss = ss.replace("(", "-").replace(")", "")
    try:
        return float(ss)
    except Exception:
        return None

def _maybe_cents_to_dollars(field_key: str, value: str) -> str:
    v = _clean_blank(value)
    if v == "":
        return ""
    k = (field_key or "").upper()
    hint = any(t in k for t in ["AMOUNT","WITHHOLD","WITHHOLDING","ADD","ADDL","OVERRIDE","VALUE"])
    if not hint:
        return v
    vv = v.replace(",", "")
    if re.fullmatch(r"-?\d+", vv):
        try:
            n = int(vv)
            if abs(n) >= 100:
                return f"{n/100:.2f}"
        except Exception:
            return v
    return v

def _strip_punct_lower(s: str) -> str:
    s = _clean_blank(s).lower()
    s = re.sub(r"[\W_]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def _normalize_for_compare(field_key: str, value: str) -> str:
    v = _clean_blank(value)

    b = _normalize_bool(v)
    if b is not None:
        return b

    v2 = _maybe_cents_to_dollars(field_key, v)

    f = _to_float(v2)
    if f is not None:
        if any(t in (field_key or "").upper() for t in ["AMOUNT","WITHHOLD","WITHHOLDING","ADD","ADDL","OVERRIDE","VALUE","PERCENT"]):
            return f"{f:.2f}"
        if abs(f - int(f)) < 1e-9:
            return str(int(f))
        return str(f)

    return v2.strip()

def _filing_status_match(a: str, b: str) -> bool:
    aa = _strip_punct_lower(a)
    bb = _strip_punct_lower(b)
    if aa == bb:
        return True
    if aa and bb and (aa in bb or bb in aa):
        return True
    return False

def load_key_mapping_yml(path_or_bytes) -> Dict[str, str]:
    if isinstance(path_or_bytes, (bytes, bytearray)):
        data = yaml.safe_load(io.BytesIO(path_or_bytes))
    else:
        with open(path_or_bytes, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)

    labels: Dict[str, str] = {}
    withholding_es = (data or {}).get("withholding_es", {})
    for _, items in (withholding_es or {}).items():
        if not isinstance(items, dict):
            continue
        for k, meta in items.items():
            if isinstance(meta, dict) and "label" in meta and k not in labels:
                labels[k] = str(meta["label"])
    return labels

def load_filing_status_map_from_txt(path_or_bytes) -> Dict[str, str]:
    if isinstance(path_or_bytes, (bytes, bytearray)):
        text = path_or_bytes.decode("utf-8", errors="ignore")
    else:
        with open(path_or_bytes, "r", encoding="utf-8") as f:
            text = f.read()

    m: Dict[str, str] = {}
    for code, label in re.findall(r'([A-Z0-9_]+)\("([^"]+)"\)', text):
        m[code] = label
    return m

def load_mapping(mapping_bytes: bytes, mapping_filename: str, uzio_key_col: str, payroll_col_col: str) -> pd.DataFrame:
    df = _read_any(mapping_bytes, mapping_filename)
    df.columns = [_norm_col(c) for c in df.columns]

    if uzio_key_col not in df.columns:
        raise ValueError(f"Mapping file is missing required column: {uzio_key_col}")
    if payroll_col_col not in df.columns:
        raise ValueError(f"Mapping file is missing required column: {payroll_col_col}")

    out = df[[uzio_key_col, payroll_col_col]].copy()
    out = out.rename(columns={uzio_key_col: "uzio_key", payroll_col_col: "payroll_col"})
    out["uzio_key"] = out["uzio_key"].astype(str).map(_clean_blank)
    out["payroll_col"] = out["payroll_col"].astype(str).map(_clean_blank)
    out = out[(out["uzio_key"] != "") & (out["payroll_col"] != "")]
    out = out.drop_duplicates()
    return out

def pivot_uzio_long(df_long: pd.DataFrame,
                    employee_id_col: str,
                    key_col: str,
                    value_col: str) -> pd.DataFrame:
    df = df_long.copy()
    df.columns = [_norm_col(c) for c in df.columns]
    for c in [employee_id_col, key_col, value_col]:
        if c not in df.columns:
            raise ValueError(f"UZIO file is missing required column: {c}")
    df[employee_id_col] = df[employee_id_col].astype(str).map(_clean_blank)
    df[key_col] = df[key_col].astype(str).map(_clean_blank)
    df[value_col] = df[value_col].astype(str).map(_clean_blank)

    df = df[df[employee_id_col] != ""]
    wide = df.pivot_table(
        index=employee_id_col,
        columns=key_col,
        values=value_col,
        aggfunc=lambda x: next((v for v in reversed(list(x)) if _clean_blank(v) != ""), _clean_blank(list(x)[-1]) if len(x) else "")
    )
    wide = wide.reset_index()
    wide.columns = [_norm_col(c) for c in wide.columns]
    return wide

def run_withholding_audit(payroll_bytes: bytes,
                          payroll_filename: str,
                          uzio_bytes: bytes,
                          uzio_filename: str,
                          mapping_bytes: bytes,
                          mapping_filename: str,
                          mapping_payroll_col_name: str,
                          payroll_employee_id_col: str,
                          uzio_employee_id_col: str,
                          uzio_key_col: str,
                          uzio_value_col: str,
                          active_flag_col: Optional[str],
                          active_values: Optional[List[str]],
                          key_label_map: Dict[str, str],
                          filing_status_map: Dict[str, str]) -> Tuple[bytes, Dict[str, int]]:

    payroll = _read_any(payroll_bytes, payroll_filename)
    uzio_long = _read_any(uzio_bytes, uzio_filename)

    payroll.columns = [_norm_col(c) for c in payroll.columns]
    uzio_long.columns = [_norm_col(c) for c in uzio_long.columns]

    if payroll_employee_id_col not in payroll.columns:
        raise ValueError(f"Payroll file missing employee id column: {payroll_employee_id_col}")

    uzio_wide = pivot_uzio_long(uzio_long, uzio_employee_id_col, uzio_key_col, uzio_value_col)

    mapping = load_mapping(mapping_bytes, mapping_filename, "Uzio Columns", mapping_payroll_col_name)

    payroll["employee_id"] = payroll[payroll_employee_id_col].astype(str).map(_clean_blank)
    uzio_wide["employee_id"] = uzio_wide[uzio_employee_id_col].astype(str).map(_clean_blank)

    payroll = payroll[payroll["employee_id"] != ""].copy()
    uzio_wide = uzio_wide[uzio_wide["employee_id"] != ""].copy()

    if active_flag_col and active_flag_col in payroll.columns and active_values:
        av = {str(x).strip().lower() for x in active_values if str(x).strip() != ""}
        payroll["_is_active"] = payroll[active_flag_col].astype(str).map(lambda x: _clean_blank(x).lower() in av)
    else:
        payroll["_is_active"] = True

    merged = payroll.merge(uzio_wide, on="employee_id", how="left", suffixes=("_payroll", "_uzio"))

    missing_in_uzio = merged[merged[uzio_employee_id_col].isna()]["employee_id"].dropna().astype(str).unique().tolist()

    rows = []
    for _, r in mapping.iterrows():
        uz_key = str(r["uzio_key"]).strip()
        pay_col = str(r["payroll_col"]).strip()

        if pay_col not in payroll.columns:
            continue

        if uz_key not in merged.columns:
            merged[uz_key] = ""

        def _uzio_val_convert(v):
            vv = _clean_blank(v)
            if vv == "":
                return ""
            if ("FILING_STATUS" in uz_key.upper()) and (vv in filing_status_map):
                return filing_status_map.get(vv, vv)
            return vv

        pay_vals = merged[pay_col].astype(str).map(_clean_blank)
        uz_vals_raw = merged[uz_key].astype(str).map(_uzio_val_convert)

        for i in range(len(merged)):
            emp_id = merged.iloc[i]["employee_id"]
            is_active = bool(merged.iloc[i]["_is_active"])

            pv_raw = pay_vals.iat[i]
            uv_raw = uz_vals_raw.iat[i]

            pv = _normalize_for_compare(uz_key, pv_raw)
            uv = _normalize_for_compare(uz_key, uv_raw)

            if uv == "" and pv in {"0", "0.00", "No"}:
                uv = pv
            if pv == "" and uv in {"0", "0.00", "No"}:
                pv = uv

            if "FILING_STATUS" in uz_key.upper():
                is_match = _filing_status_match(pv, uv)
            else:
                is_match = (pv == uv)

            if not is_match:
                rows.append({
                    "employee_id": emp_id,
                    "is_active": is_active,
                    "withholding_field_key": uz_key,
                    "withholding_field_label": key_label_map.get(uz_key, uz_key),
                    "payroll_column": pay_col,
                    "payroll_value_raw": pv_raw,
                    "uzio_value_raw": uv_raw,
                    "payroll_value_norm": pv,
                    "uzio_value_norm": uv
                })

    mismatches = pd.DataFrame(rows)

    all_mismatch = mismatches.copy()
    active_mismatch = mismatches[mismatches["is_active"] == True].copy() if not mismatches.empty else mismatches.copy()

    metrics = {
        "Payroll Employees": int(payroll["employee_id"].nunique()),
        "UZIO Employees": int(uzio_wide["employee_id"].nunique()),
        "Missing in UZIO": int(len(missing_in_uzio)),
        "All Mismatch Rows": int(len(all_mismatch)),
        "Active Mismatch Rows": int(len(active_mismatch)),
        "Employees w/ Any Mismatch": int(all_mismatch["employee_id"].nunique()) if not all_mismatch.empty else 0,
        "Active Employees w/ Any Mismatch": int(active_mismatch["employee_id"].nunique()) if not active_mismatch.empty else 0,
    }

    summary_df = pd.DataFrame(
        [["Report Generated", datetime.now().strftime("%Y-%m-%d %H:%M:%S")], ["", ""]] +
        [[k, v] for k, v in metrics.items()] +
        [["", ""], ["Missing in UZIO (employee_id)", ""]],
        columns=["Metric", "Value"]
    )
    if missing_in_uzio:
        miss_df = pd.DataFrame({"Metric": [""] * len(missing_in_uzio), "Value": missing_in_uzio})
        summary_df = pd.concat([summary_df, miss_df], ignore_index=True)
    else:
        summary_df = pd.concat([summary_df, pd.DataFrame([["", "(none)"]], columns=["Metric","Value"])], ignore_index=True)

    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary_df.to_excel(writer, sheet_name="Summary", index=False)
        active_mismatch.to_excel(writer, sheet_name="Active_Mismatches", index=False)
        all_mismatch.to_excel(writer, sheet_name="All_Mismatches", index=False)

        for sheet in ["Summary","Active_Mismatches","All_Mismatches"]:
            ws = writer.book[sheet]
            for col in ws.columns:
                max_len = 0
                col_letter = col[0].column_letter
                for cell in col:
                    try:
                        max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
                    except Exception:
                        pass
                ws.column_dimensions[col_letter].width = min(max_len + 2, 60)

    return out.getvalue(), metrics
