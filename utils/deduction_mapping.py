"""Load an Employee Deduction Mapping CSV into the dict the deduction audits use.

The ADP Prior Payroll Setup Helper already emits this file as
`<Client>_EE_Deductions_mapping.csv` with four columns:

    Source Deduction Code, Source Deduction Code Name,
    Uzio Deduction Code,   Uzio Deduction Code Name

which is exactly the mapping the deduction audits used to ask for by hand, one
dropdown per deduction. `_run_deduction_audit` looks a row up by DEDUCTION
DESCRIPTION and falls back to DEDUCTION CODE:

    deduction_name = mapping.get(raw_desc) or mapping.get(raw_code)

so both the source NAME and the source CODE are registered as keys pointing at
the Uzio deduction name. The audit also lower-cases its lookups, so casing here
does not matter (unlike the onboarding API, which is why the setup helper copies
source names verbatim).

Rows the file cannot answer are reported rather than dropped. The audit skips an
unmapped deduction silently (`continue`), so anything this loader cannot resolve
has to be shown on screen or it vanishes from the comparison without a trace.
"""

import io

import pandas as pd

SOURCE_CODE_COL = "Source Deduction Code"
SOURCE_NAME_COL = "Source Deduction Code Name"
UZIO_CODE_COL = "Uzio Deduction Code"
UZIO_NAME_COL = "Uzio Deduction Code Name"
REQUIRED_COLUMNS = [SOURCE_CODE_COL, SOURCE_NAME_COL, UZIO_CODE_COL, UZIO_NAME_COL]


def _norm(c):
    return str(c).replace("\u00a0", " ").strip().lower()


def _blank(v):
    return v is None or pd.isna(v) or str(v).strip() == ""


def load_deduction_mapping(file):
    """Parse the mapping file.

    Returns (mapping, report).
      mapping : {source name or code -> Uzio deduction name}
      report  : {'pairs': [(source_name, source_code, uzio_name)],
                 'blank_target': [source_name],   # row present, Uzio side empty
                 'conflicts': [(key, kept, dropped)]}

    Raises ValueError with a readable message when the file is not this mapping.
    """
    name = (getattr(file, "name", "") or "").lower()
    file.seek(0)
    raw = file.read()
    file.seek(0)
    if name.endswith(".csv"):
        # utf-8-sig so a BOM-prefixed header still matches.
        df = pd.read_csv(io.BytesIO(raw), dtype=str, encoding="utf-8-sig")
    else:
        df = pd.read_excel(io.BytesIO(raw), dtype=str)

    by_norm = {_norm(c): c for c in df.columns}
    missing = [c for c in REQUIRED_COLUMNS if _norm(c) not in by_norm]
    if missing:
        raise ValueError(
            "This does not look like an Employee Deduction Mapping file — "
            "missing column(s): " + ", ".join(missing) + ". Expected: "
            + ", ".join(REQUIRED_COLUMNS) + "."
        )

    src_name_c = by_norm[_norm(SOURCE_NAME_COL)]
    src_code_c = by_norm[_norm(SOURCE_CODE_COL)]
    uzio_name_c = by_norm[_norm(UZIO_NAME_COL)]

    mapping = {}
    pairs, blank_target, conflicts = [], [], []
    for _, row in df.iterrows():
        s_name = "" if _blank(row.get(src_name_c)) else str(row[src_name_c]).strip()
        s_code = "" if _blank(row.get(src_code_c)) else str(row[src_code_c]).strip()
        u_name = "" if _blank(row.get(uzio_name_c)) else str(row[uzio_name_c]).strip()
        if not s_name and not s_code:
            continue
        if not u_name:
            # The setup helper leaves this empty for deductions it deliberately
            # does not assign to employees (garnishments, child support, tax
            # liens) and for anything it could not resolve. Either way the audit
            # cannot compare it, so say so instead of silently skipping.
            blank_target.append(s_name or s_code)
            continue
        pairs.append((s_name, s_code, u_name))
        for key in (s_name, s_code):
            if not key:
                continue
            prev = mapping.get(key)
            if prev is not None and prev != u_name:
                conflicts.append((key, prev, u_name))
                continue
            mapping[key] = u_name

    return mapping, {"pairs": pairs, "blank_target": blank_target, "conflicts": conflicts}


def unmapped_source_deductions(source_deductions, mapping):
    """Which of the vendor file's deductions the mapping cannot answer.

    Matched case-insensitively, the same way the audit looks them up.
    """
    keys = {k.strip().lower() for k in mapping}
    return sorted(d for d in source_deductions if str(d).strip().lower() not in keys)
