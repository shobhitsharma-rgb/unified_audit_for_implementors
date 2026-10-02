"""Read an uploaded file whether the user exported it as Excel or as CSV.

Every tool takes files the client exports from ADP / Paycom / Uzio, and the same
report can come out as .xlsx one day and .csv the next. These helpers let a tool
accept both without each module growing its own sniffing logic.

What belongs here: reading. Nothing here writes a file, and nothing here knows
about a particular vendor's column names.

Note the one case these helpers deliberately do NOT cover: a file that is loaded
with openpyxl, filled in and saved back out (the Uzio census .xlsm, the time-off
and qualified-overtime templates). Those must stay Excel — a CSV cannot carry the
sheets and formatting the tool writes into them.
"""
import io

import pandas as pd

CSV_EXTENSIONS = (".csv", ".txt", ".tsv")


def is_csv_upload(file) -> bool:
    """True when the upload is a text/CSV file rather than a workbook.

    Decided on the file name first (Streamlit gives us `.name`), then on the
    bytes: a real .xlsx/.xlsm is a zip ("PK"), an old .xls starts with 0xD0CF.
    """
    name = (getattr(file, "name", "") or "").lower()
    if name.endswith(CSV_EXTENSIONS):
        return True
    if name.endswith((".xlsx", ".xlsm", ".xls")):
        return False
    try:
        head = file.getvalue()[:8]
    except Exception:
        return False
    return not (head.startswith(b"PK") or head.startswith(b"\xd0\xcf\x11\xe0"))


def read_table(file, sheet_name=0, **kwargs):
    """`pd.read_excel` for a workbook, `pd.read_csv` for a CSV, same call shape.

    Excel-only arguments are dropped for CSV, and `sheet_name=None` (every sheet)
    returns a one-entry dict so callers that loop over sheets keep working.
    A CSV that is not UTF-8 is retried as latin-1, which is what ADP and Paycom
    exports occasionally are.
    """
    if hasattr(file, "seek"):
        file.seek(0)
    if not is_csv_upload(file):
        return pd.read_excel(file, sheet_name=sheet_name, **kwargs)

    data = file.getvalue()
    csv_kwargs = {k: v for k, v in kwargs.items() if k not in ("engine", "keep_vba")}
    try:
        df = pd.read_csv(io.BytesIO(data), **csv_kwargs)
    except UnicodeDecodeError:
        df = pd.read_csv(io.BytesIO(data), encoding="latin1", **csv_kwargs)
    return {"Sheet1": df} if sheet_name is None else df


def as_excel_stream(file):
    """The upload as something openpyxl can open.

    A workbook is handed back untouched. A CSV is re-packed into an in-memory
    .xlsx with one sheet, so readers built on openpyxl (ADP writes `=ROUND()`
    formulas, which pandas cannot see) work on CSV input without a second code
    path. Values from a CSV are plain, so formula handling simply passes through.
    """
    if hasattr(file, "seek"):
        file.seek(0)
    if not is_csv_upload(file):
        return file

    import openpyxl

    data = file.getvalue()
    try:
        df = pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False)
    except UnicodeDecodeError:
        df = pd.read_csv(io.BytesIO(data), dtype=str, keep_default_na=False, encoding="latin1")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append([str(c) for c in df.columns])
    for row in df.itertuples(index=False):
        ws.append([None if v == "" else v for v in row])
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    out.name = (getattr(file, "name", "upload") or "upload") + ".xlsx"
    return out
