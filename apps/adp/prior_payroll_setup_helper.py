"""ADP - Prior Payroll Setup Helper Tool.

Discovers what to configure in Uzio for an ADP Prior Payroll migration.
Given a sanitized ADP Prior Payroll file plus the State Tax Code master CSV,
emits an Excel workbook with:

  - Earnings_Codes      - REGULAR / OVERTIME + every ADDITIONAL EARNINGS code
                          with $ total, employee count, paired hours, avg rate.
  - Contributions       - 401k / 403b / 457 / Roth / HSA / FSA voluntary codes.
  - Deductions          - all other voluntary deductions, with pre-tax vs
                          post-tax verdict per code.
  - Taxes_Discovered    - every '* - EMPLOYEE TAX' / '* - EMPLOYER TAX' column.
  - Tax_Mapping         - one row per (tax_type, state) in the
                          'Payroll_Mappings_Tax_Mapping_CORRECTED' format.
  - Bonus_Classification- FLSA test (discretionary vs non-discretionary).

Pre/post-tax algorithm:
  gap_FIT = TOTAL EARNINGS - FEDERAL INCOME - EMPLOYEE TAXABLE.
  Find any subset of a row's non-zero deductions summing to gap_FIT (within
  $0.02). Every member of any passing subset is pre-tax for FIT. ONE positive
  proof anywhere in the file = pre-tax for everyone (the rule never varies
  per employee). Same logic for FICA / MEDI / SIT taxables to derive the
  flavor: section_125 (pre-FIT/FICA/MEDI/SIT) vs 401k_traditional
  (pre-FIT/SIT only, NOT pre-FICA/MEDI).
"""

import base64
import io
import json
import os
import re
from itertools import combinations

import pandas as pd
import streamlit as st

from apps.adp.prior_payroll_sanity import read_input_file, _find_col
from utils.audit_utils import clean_money_val


# ---------- helpers ----------

def _num(v):
    try:
        return clean_money_val(v)
    except Exception:
        return 0.0


def _strip_prefix(col, prefixes):
    s = str(col).strip()
    for p in prefixes:
        if s.upper().startswith(p.upper()):
            rest = s[len(p):].lstrip(" :").strip()
            return rest
    return s


# ---------- column categorization ----------

EARN_PREFIXES = ["ADDITIONAL EARNINGS"]
HOUR_PREFIXES = ["ADDITIONAL HOURS"]
DED_PREFIX = "VOLUNTARY DEDUCTION"
MEMO_PREFIX = "MEMO"

CONTRIB_PATTERN = re.compile(
    r"\b(401[Kk]?|403[Bb]?|457|ROTH|HSA|FSA|RETIRE|RETIREMENT)\b"
)


def categorize_columns(df):
    earn_cols, hour_cols, tax_cols, taxable_cols = [], [], [], []
    ded_cols, memo_cols = [], []
    for c in df.columns:
        s = str(c).strip(); u = s.upper()
        if u in ("REGULAR EARNINGS", "OVERTIME EARNINGS"):
            earn_cols.append(c)
        elif u.startswith("ADDITIONAL EARNINGS"):
            earn_cols.append(c)
        elif u in ("REGULAR HOURS", "OVERTIME HOURS"):
            hour_cols.append(c)
        elif u.startswith("ADDITIONAL HOURS"):
            hour_cols.append(c)
        elif u.startswith(DED_PREFIX):
            ded_cols.append(c)
        elif u.startswith(MEMO_PREFIX):
            memo_cols.append(c)
        elif u.endswith("TAXABLE"):
            taxable_cols.append(c)
        elif u.endswith("EMPLOYEE TAX") or u.endswith("EMPLOYER TAX"):
            if u.startswith("TOTAL "):
                continue  # aggregate, not a real tax row
            tax_cols.append(c)

    # Positional earnings window — authoritative for ADP Payroll History exports:
    # the earnings block ALWAYS sits between TOTAL HOURS and TOTAL EARNINGS
    # (REGULAR EARNINGS, OVERTIME EARNINGS, then the additional-earnings columns,
    # which may carry the "ADDITIONAL EARNINGS :" prefix or be bare labels).
    # The prefix scan above stays as the fallback for files without the markers.
    upper_cols = [str(c).strip().upper() for c in df.columns]
    if "TOTAL HOURS" in upper_cols and "TOTAL EARNINGS" in upper_cols:
        i0 = upper_cols.index("TOTAL HOURS")
        i1 = upper_cols.index("TOTAL EARNINGS")
        if i1 > i0 + 1:
            earn_cols = list(df.columns)[i0 + 1:i1]

    return {
        "earnings": earn_cols, "hours": hour_cols,
        "taxes": tax_cols, "taxables": taxable_cols,
        "deductions": ded_cols, "memos": memo_cols,
    }


# ---------- catalog builders ----------

def build_earnings_catalog(df, earn_cols, hour_cols):
    hour_lookup = {}
    for h in hour_cols:
        u = str(h).strip().upper()
        if u == "REGULAR HOURS":
            hour_lookup["REGULAR EARNINGS"] = h
        elif u == "OVERTIME HOURS":
            hour_lookup["OVERTIME EARNINGS"] = h
        else:
            code = _strip_prefix(h, HOUR_PREFIXES)
            hour_lookup[f"ADDITIONAL EARNINGS  : {code}"] = h
            hour_lookup[code] = h

    rows = []
    for c in earn_cols:
        amounts = df[c].apply(_num)
        total = float(amounts.sum())
        emp_count = int((amounts != 0).sum())
        u = str(c).strip().upper()
        if u == "REGULAR EARNINGS":
            code = "REGULAR"; kind = "Regular Wage"
        elif u == "OVERTIME EARNINGS":
            code = "OVERTIME"; kind = "Overtime"
        else:
            code = _strip_prefix(c, EARN_PREFIXES); kind = "Additional Earning"

        h_col = hour_lookup.get(str(c).strip()) or hour_lookup.get(code)
        if h_col is not None and h_col in df.columns:
            hours_total = float(df[h_col].apply(_num).sum())
            avg_rate = total / hours_total if hours_total > 0 else None
        else:
            hours_total = None; avg_rate = None

        rows.append({
            "Source Column": str(c).strip(), "Code": code, "Kind": kind,
            "Total $": round(total, 2), "Employees": emp_count,
            "Total Hours": round(hours_total, 2) if hours_total is not None else None,
            "Avg Rate ($/hr)": round(avg_rate, 4) if avg_rate is not None else None,
        })
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# UZIO "Add Earning" form mapping config + enrichment
#
# Ported from the mature Paycom Prior Payroll Setup Helper. The enrichment logic
# is platform-agnostic — it keys off a generic (Type Code, Type Description) pair
# — so only the EXTRACTION differs. ADP earnings come from columns (REGULAR /
# OVERTIME EARNINGS + every "ADDITIONAL EARNINGS : <name>" column), which
# build_earnings_catalog() already turns into {Code, Kind, ...} rows;
# adp_earnings_to_setup_rows() adapts those into the (Type Code, Type Description)
# shape this layer consumes.
#
# The UZIO Earning form fields:
#   Earning Type (dropdown, driver)  Earning Name (text)  Display Order (text)
#   Paid Earning (Yes/No)  Hourly Based Earning (Yes/No)
#   Rate Determination Factor / Rate
#   Subject to garnishment disposable income? (Yes/No)
#   Subject to Workers' Compensation (Yes/No)  Taxability Type (dropdown)
#   Include Bonus in Overtime (bonus only)  Time Off Policy  W-2 Box
# Earning Type is the driver: for a mapped type UZIO auto-fills + locks Hourly /
# Taxability; "Other" leaves them editable.
# ─────────────────────────────────────────────────────────────────────────────

EARNING_TYPE_OTHER = "Other"
EARNING_TAXABILITY_TAXABLE = "Taxable"
EARNING_TAXABILITY_NONTAX = "Non-Taxable"

# Full UZIO "Earning Type" dropdown options (verbatim from the live UI). Used to
# populate the manual-mapping dropdown for earnings that fall to "Other".
UZIO_EARNING_TYPES = [
    "Bonus", "Commission", "Vacation", "Reimbursements",
    "Group Term Life Insurance", "Cash Tip", "Pay Check Tips",
    "Expense reimbursement", "Mileage reimbursement", "Stock Options",
    "Severance", "3rd Party Sick Pay - Taxable", "3rd Party Sick Pay - Nontaxable",
    "Dividend", "Moving Expenses", "Clothing Allowance", "Tool Allowance",
    "Tuition Assistance", "Non Tax Tuition Assistance", "Allocated Tips",
    "COVID 100 Sick", "COVID 2/3 Sick", "COVID Family Leave", "Sick", "Other",
    "Owner's Draw", "Unpaid Time Off", "OT Adjustment", "Station Closure",
    "DA Recognition - TWA",
]


def _edesc(type_description):
    return " ".join((type_description or "").lower().split())


# ── (1) UZIO auto-created / default earnings — SKIP, never re-create ──────────
# UZIO seeds these on every company; the setup must not re-create them. Detected
# from the (code, description) via DEFAULT_EARNING_RULES; first match wins.
DEFAULT_EARNING_RULES = [
    (lambda d: "look back" in d or "lookback" in d,                 "Lookback bonus"),
    (lambda d: "realtime" in d or "real time" in d,                 "Realtime bonus"),
    (lambda d: "ot adjustment" in d or d == "otadj",                "OT Adjustment"),
    (lambda d: "double" in d and "overtime" in d,                   "Double Overtime"),
    (lambda d: "overtime" in d and ("weighted" in d or "(weighted)" in d), "Overtime"),
    (lambda d: d in ("overtime", "overtime hours", "ot", "overtime earnings"), "Overtime"),
    (lambda d: "holiday" in d and "premium" in d,                   "Holiday Premium"),
    (lambda d: d == "holiday",                                      "Holiday"),
    (lambda d: "meal break" in d,                                   "Meal Break Premium"),
    (lambda d: "rest break" in d,                                   "Rest Break Premium"),
    (lambda d: "retro" in d and "overtime" in d and "pay" in d,     "OT Adjustment"),
    (lambda d: "pto" in d and ("balance" in d or "payout" in d),    "PTO Balance Payout"),
    (lambda d: d in ("reimbursements", "reimbursement"),            "Reimbursements"),
    (lambda d: (d == "regular" or d == "regular wage" or d == "regular pay"
                or d == "regular earnings" or d.startswith("regular "))
               and "retro" not in d,                               "Regular Wage"),
]


def default_earning_name(type_code, type_description):
    """Return the UZIO default earning name this earning corresponds to (so it is
    SKIPPED), or "" if it's a real earning to create."""
    d = _edesc(type_description)
    for pred, uzio_name in DEFAULT_EARNING_RULES:
        try:
            if pred(d):
                return uzio_name
        except Exception:
            pass
    return ""


# ── (2) Earning Type mapping (driver dropdown) ───────────────────────────────
# Keyword -> UZIO Earning Type dropdown label. Ordered; first match wins.
EARNING_TYPE_KEYWORD_MAP = [
    ("unpaid time off",  "Unpaid Time Off"),
    ("unpaid leave",     "Unpaid Time Off"),
    ("vto",              "Unpaid Time Off"),
    ("paid time off",    "Vacation"),
    ("vacation",         "Vacation"),
    ("pto",              "Vacation"),
    ("tuition reimbursement",     "Non Tax Tuition Assistance"),
    ("non tax tuition",           "Non Tax Tuition Assistance"),
    ("tuition",                   "Tuition Assistance"),
    ("expense reimburs", "Expense reimbursement"),
    ("mileage",          "Mileage reimbursement"),
    ("reimburs",         "Reimbursements"),
    ("commission",       "Commission"),
    ("severance",        "Severance"),
    ("group term life",  "Group Term Life Insurance"),
    ("paycheck tip",     "Pay Check Tips"),
    ("pay check tip",    "Pay Check Tips"),
    ("tip",              "Cash Tip"),
    ("sick",             "Sick"),
    ("bonus",            "Bonus"),
    ("station closure",  "Station Closure"),
]
EARNING_TYPE_EXACT_MAP = {}

# ── (3) Hourly + Taxability are driven by Earning Type. For a mapped type UZIO
# auto-fills + locks these; for "Other" they're editable defaults. ──
EARNING_TYPE_HOURLY_NO = {
    "Bonus", "OT Adjustment", "Reimbursements", "Expense reimbursement",
    "Mileage reimbursement", "Tuition Assistance", "Non Tax Tuition Assistance",
    "Severance", "Commission", "Cash Tip", "Pay Check Tips",
}
EARNING_TYPE_NONTAX = {
    "Reimbursements", "Expense reimbursement", "Mileage reimbursement",
    "Non Tax Tuition Assistance",
}

# Defaults for "Other" earnings (per implementor: hourly Yes, Taxable).
EARNING_DEFAULT_PAID = "Yes"
EARNING_OTHER_HOURLY = "Yes"
EARNING_DEFAULT_DISPOSABLE = "Yes"
EARNING_DEFAULT_WORKERS_COMP = "Yes"
EARNING_DEFAULT_W2_BOX = "Not Required"

# Per-Earning-Type overrides for the "disposable income" / "Workers' Comp" radios.
EARNING_TYPE_FIELD_DEFAULTS = {
    "Reimbursements": {"disposable": "No", "workersComp": "No"},
}

# Rate Determination Factor only appears for "Other" earnings with Hourly=Yes.
EARNING_RATE_FACTOR_MULTIPLES = "Multiples of Regular Wage Rate"
EARNING_RATE_DEFAULT_VALUE = "1"
EARNING_NA = "NA"

EARNING_INCLUDE_OT_COL = "Include Bonus in Overtime Calculation"
EARNING_INCLUDE_OT_DEFAULT = "No"

EARNING_TIMEOFF_COL = "Time Off Policy"
EARNING_TIMEOFF_VACATION = "All"
EARNING_TIMEOFF_ALL = "All"

# Canonical column order for the Earnings tab + UI dataframe.
EARNING_OUTPUT_COLUMNS = [
    "Type Code", "Type Description", "Earning Type", "Earning Name",
    "Display Order", "Paid Earning", "Hourly Based Earning",
    "Rate Determination Factor", "Rate",
    "Subject to garnishment disposable income", "Subject to Workers Compensation",
    "Taxability Type", EARNING_INCLUDE_OT_COL, EARNING_TIMEOFF_COL, "W-2 Box",
]


def autosync_row_key(type_code, type_description):
    """Stable, unique key for an earning row, shared by the UI toggle/override
    maps and the enrichment lookup. The (code, description) pair is unique because
    the upstream extractor dedupes on exactly that pair."""
    return f"{(type_code or '').strip()}||{(type_description or '').strip()}"


def is_bonus_earning(type_description):
    """True if this earning is a bonus that gets the 'Include in Overtime' /
    discretionary question — contains 'bonus' but is NOT a Lookback / Realtime
    bonus (those are system defaults, skipped)."""
    d = _edesc(type_description)
    if "bonus" not in d:
        return False
    if "look back" in d or "lookback" in d or "realtime" in d or "real time" in d:
        return False
    return True


def map_to_earning_type(type_code, type_description):
    """Return the UZIO Earning Type for an earning. Forced-Other rules first,
    then exact map, then keyword map, then "Other"."""
    d = _edesc(type_description)
    # FORCE Other for hourly bonuses: UZIO's Bonus type locks Hourly = No, but an
    # hourly bonus must stay hourly, so create it as "Other" (Hourly=Yes).
    if "bonus" in d and "hour" in d:
        return EARNING_TYPE_OTHER
    if d in EARNING_TYPE_EXACT_MAP:
        return EARNING_TYPE_EXACT_MAP[d]
    blob = f"{type_code or ''} {type_description or ''}".lower()
    for kw, etype in EARNING_TYPE_KEYWORD_MAP:
        if kw in blob:
            return etype
    return EARNING_TYPE_OTHER


# ── ADP earning label parsing + code catalog ─────────────────────────────────
# ADP mashes code + description into one label ("BNH-BONUS HOURS"), and some
# columns carry ONLY the code ("LK2"). We split into Type Code / Type Description
# (like Paycom has natively), and resolve code-only labels through a persistent
# catalog: seeds below + every code→description pair learned from files we've
# already seen (saved to adp_earning_code_catalog.json next to this module), so
# a future code-only file can still be mapped.

# Authoritative ADP earning-code map (per implementor, 2026-06-11). These codes
# ALWAYS mean this earning — the seed name wins even when the file carries its
# own description (e.g. "PVT-PREVIOUS OT" is still OT Adjustment). Rohit extends
# this list over time; codes NOT listed here are resolved from the file label or
# the learned catalog JSON.
ADP_EARNING_CODE_SEEDS = {
    "LK2": "Lookback Bonus",
    "NA2": "Realtime Bonus",
    "PVT": "OT Adjustment",
    "SLE": "Station Closure",
    "TRA": "Training",
    "ITR": "Tuition Reimbursement",
    "BND": "Bonus Discretionary",
    "BNS": "Bonus",
}

ADP_EARNING_CODE_CATALOG_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "adp_earning_code_catalog.json")

# "BNH-BONUS HOURS" → ("BNH", "BONUS HOURS"). Codes run 1-4 chars — deductions
# use single letters too ("K-401K % RETIRE", "R-ROTH"). Code-ONLY labels still
# need 2+ chars (a bare 1-char column label is not a code).
_CODE_DESC_RE = re.compile(r"^([A-Za-z0-9$]{1,4})\s*-\s*(.+)$")
_CODE_ONLY_RE = re.compile(r"^[A-Za-z0-9$]{2,4}$")

# Tokens kept ALL-CAPS when prettifying an ALL-CAPS ADP description.
_TITLE_KEEP_UPPER = {
    "OT", "PTO", "GTL", "NYC", "NY", "FLSA", "HSA", "FSA", "SUI", "SDI",
    "FLI", "ESSTA", "ADP", "II", "III", "MTA", "FUTA", "FICA", "MEDI",
}


def _smart_title(s):
    """ADP descriptions arrive ALL-CAPS ("BONUS HOURS"); UZIO names should read
    "Bonus Hours". Title-case each word but keep known acronyms (OT, PTO, ...)
    and digit-bearing tokens uppercase. Sub-tokens around & and / are handled
    separately so "ACCIDENTAL D&D" → "Accidental D&D" (not "D&d")."""
    def _part(p):
        if p.upper() in _TITLE_KEEP_UPPER or any(ch.isdigit() for ch in p):
            return p.upper()
        return p.capitalize()

    out = []
    for w in str(s).split():
        pieces = re.split(r"([&/])", w)
        out.append("".join(p if p in ("&", "/") else _part(p) for p in pieces))
    return " ".join(out)


def _split_code_desc(label):
    """Split an ADP earning label into (code, description). Returns ("", label)
    when there's no code prefix, (code, "") when the label is code-only."""
    s = str(label).strip()
    m = _CODE_DESC_RE.match(s)
    if m:
        return m.group(1).upper(), m.group(2).strip()
    if _CODE_ONLY_RE.match(s):
        return s.upper(), ""
    return "", s


def load_earning_code_catalog():
    """Everything learned from previously-analyzed files, overlaid with the
    seeds — the implementor's seed list is authoritative, so it wins over any
    learned (file-derived) value for the same code."""
    catalog = {}
    try:
        with open(ADP_EARNING_CODE_CATALOG_PATH, encoding="utf-8") as f:
            saved = json.load(f)
        if isinstance(saved, dict):
            catalog.update({str(k).upper(): str(v) for k, v in saved.items()})
    except (OSError, ValueError):
        pass
    catalog.update(ADP_EARNING_CODE_SEEDS)
    return catalog


def save_learned_earning_codes(new_codes):
    """Merge code→description pairs learned from the current file(s) into the
    catalog JSON. Best-effort: a read-only filesystem must not break analysis."""
    if not new_codes:
        return
    try:
        try:
            with open(ADP_EARNING_CODE_CATALOG_PATH, encoding="utf-8") as f:
                existing = json.load(f)
            if not isinstance(existing, dict):
                existing = {}
        except (OSError, ValueError):
            existing = {}
        merged = {**existing, **{str(k).upper(): str(v) for k, v in new_codes.items()}}
        if merged != existing:
            with open(ADP_EARNING_CODE_CATALOG_PATH, "w", encoding="utf-8") as f:
                json.dump(merged, f, indent=2, sort_keys=True)
    except OSError:
        pass


def adp_earnings_to_setup_rows(earn_catalog_rows):
    """Adapt build_earnings_catalog() output → (Type Code, Type Description) rows
    for the UZIO enrichment layer.

    Splits each ADP label into code + description ("BNH-BONUS HOURS" → BNH /
    "Bonus Hours") so the UZIO Earning Name never carries the code prefix.
    Code-only labels ("LK2") resolve through the earning-code catalog; labels
    with both parts teach the catalog as a side effect. Unresolvable code-only
    labels keep the code as the description and are flagged `_Unknown Code` so
    the UI can call them out. Discovery stats are carried along (prefixed `_`);
    enrich_earnings_for_uzio preserves unknown keys via {**r, ...} and the output
    dataframe selects only the canonical columns, so they never leak."""
    catalog = load_earning_code_catalog()
    learned = {}
    rows = []
    for r in earn_catalog_rows or []:
        label = str(r.get("Code", "")).strip()
        if not label:
            continue
        code, desc = _split_code_desc(label)
        unknown = False
        if code in ADP_EARNING_CODE_SEEDS:
            # Authoritative: the seed name wins even over the file's own
            # description (PVT-PREVIOUS OT is still OT Adjustment).
            desc = ADP_EARNING_CODE_SEEDS[code]
        elif code and desc:
            desc = _smart_title(desc)
            if catalog.get(code) != desc:
                learned[code] = desc
        elif code and not desc:
            desc = catalog.get(code, "")
            if not desc:
                desc = code
                unknown = True
        else:
            desc = _smart_title(desc)
        rows.append({
            "Type Code": code,
            "Type Description": desc,
            "_Source Label": label,
            "_Source Column": str(r.get("Source Column", "") or label).strip(),
            "_Unknown Code": unknown,
            "_Kind": r.get("Kind"),
            "_Total $": r.get("Total $"),
            "_Employees": r.get("Employees"),
            "_Total Hours": r.get("Total Hours"),
            "_Avg Rate ($/hr)": r.get("Avg Rate ($/hr)"),
        })
    save_learned_earning_codes(learned)
    return _dedupe_setup_rows(rows)


def _dedupe_setup_rows(rows):
    """Merge rows sharing the same (Type Code, Type Description) — e.g. the same
    deduction labeled with slightly different column text across two uploaded
    files, or two labels resolving to one seeded name. Duplicate keys would
    otherwise crash Streamlit (every widget key embeds the (code, desc) pair).
    Numeric `_` stats are summed; first row wins for everything else."""
    merged = {}
    order = []
    for r in rows:
        k = (r.get("Type Code"), r.get("Type Description"))
        if k not in merged:
            merged[k] = dict(r)
            order.append(k)
            continue
        tgt = merged[k]
        for f in ("_Total $", "_Employees", "_Total Hours"):
            a, b = tgt.get(f), r.get(f)
            if isinstance(a, (int, float)) and isinstance(b, (int, float)):
                tgt[f] = round(a + b, 2)
            elif a is None:
                tgt[f] = b
    return [merged[k] for k in order]


def filter_default_uzio_earnings(rows):
    """Split extracted earnings into (kept, skipped). UZIO auto-creates a set of
    earnings on every company (Regular Wage, Overtime, Holiday, etc.); the setup
    must not re-create them. Each skipped row is annotated with the matched UZIO
    default name. Returns (kept_rows, skipped_rows)."""
    if not rows:
        return [], []
    kept, skipped = [], []
    for r in rows:
        uzio_name = default_earning_name(r.get("Type Code"), r.get("Type Description"))
        if uzio_name:
            skipped.append({**r, "UZIO Default Earning": uzio_name})
        else:
            kept.append(r)
    return kept, skipped


def enrich_earnings_for_uzio(rows, start_display_order=20, include_in_ot_map=None,
                             earning_type_override_map=None):
    """Add the UZIO Add-Earning form fields to each (already default-filtered)
    earning row as columns.

    `start_display_order`: Display Order for the first created earning; each
    subsequent one increments by 1.
    `include_in_ot_map`: optional {autosync_row_key -> "Yes"/"No"} from the
    "Is the earning Non Discretionary?" UI section; only consulted for bonuses.
    `earning_type_override_map`: optional {autosync_row_key -> Earning Type} from
    the "Map earnings to an earning type" UI accordion.
    """
    include_in_ot_map = include_in_ot_map or {}
    earning_type_override_map = earning_type_override_map or {}
    out = []
    order = start_display_order
    for r in rows:
        etype = map_to_earning_type(r.get("Type Code"), r.get("Type Description"))
        ovr = earning_type_override_map.get(
            autosync_row_key(r.get("Type Code"), r.get("Type Description")))
        if ovr:
            etype = ovr
        if etype == EARNING_TYPE_OTHER:
            hourly = EARNING_OTHER_HOURLY
            taxability = EARNING_TAXABILITY_TAXABLE
        else:
            hourly = "No" if etype in EARNING_TYPE_HOURLY_NO else "Yes"
            taxability = EARNING_TAXABILITY_NONTAX if etype in EARNING_TYPE_NONTAX else EARNING_TAXABILITY_TAXABLE
        if etype == EARNING_TYPE_OTHER and hourly == "Yes":
            rate_factor = EARNING_RATE_FACTOR_MULTIPLES
            rate_value = EARNING_RATE_DEFAULT_VALUE
        elif etype == "Unpaid Time Off":
            rate_factor = EARNING_RATE_FACTOR_MULTIPLES
            rate_value = "0"
        else:
            rate_factor = EARNING_NA
            rate_value = EARNING_NA
        if etype == "Bonus":
            include_ot = include_in_ot_map.get(
                autosync_row_key(r.get("Type Code"), r.get("Type Description")),
                EARNING_INCLUDE_OT_DEFAULT,
            )
        else:
            include_ot = EARNING_NA
        disposable = EARNING_DEFAULT_DISPOSABLE
        workers_comp = EARNING_DEFAULT_WORKERS_COMP
        _fov = EARNING_TYPE_FIELD_DEFAULTS.get(etype)
        if _fov:
            disposable = _fov.get("disposable", disposable)
            workers_comp = _fov.get("workersComp", workers_comp)
        if etype == "Vacation":
            timeoff = EARNING_TIMEOFF_VACATION
        elif etype == "Unpaid Time Off":
            timeoff = EARNING_TIMEOFF_ALL
        else:
            timeoff = EARNING_NA
        out.append({
            **r,
            "Earning Type": etype,
            "Earning Name": r.get("Type Description", ""),
            "Display Order": str(order),
            "Paid Earning": EARNING_DEFAULT_PAID,
            "Hourly Based Earning": hourly,
            "Rate Determination Factor": rate_factor,
            "Rate": rate_value,
            "Subject to garnishment disposable income": disposable,
            "Subject to Workers Compensation": workers_comp,
            "Taxability Type": taxability,
            EARNING_INCLUDE_OT_COL: include_ot,
            EARNING_TIMEOFF_COL: timeoff,
            "W-2 Box": EARNING_DEFAULT_W2_BOX,
        })
        order += 1
    return out


# ─────────────────────────────────────────────────────────────────────────────
# UZIO "Add Deduction" form mapping config
#
# Ported from the Paycom helper. Master Deductions List drives the whole form;
# Method is Fixed $ / % of Gross Pay / % of Disposable Net Pay; Auto-Sync only
# exists for benefit-type masters. Unmappable rows surface as <NEEDS REVIEW>
# (with a manual dropdown in the UI) so a wrong guess is never shipped.
#
# ADP difference vs Paycom: deduction codes are numeric and PER-CLIENT ("73",
# "75"), so there is no explicit code→master table — mapping is keyword
# inference on the description, made tax-aware by the subset-sum Pre/Post Tax
# verdict the ADP classifier proves empirically.
# ─────────────────────────────────────────────────────────────────────────────

NEEDS_REVIEW = "<NEEDS REVIEW>"
MASTER_OTHER = "Other"

# Full UZIO "Master Deductions List" dropdown options (verbatim from the live
# UI). Keep in exact UZIO spelling — the Tampermonkey script matches them.
UZIO_MASTER_DEDUCTIONS = [
    "401(k) Loan", "401k", "Accident Insurance After-tax", "Advance",
    "Basic Life and AD&D", "Cancer Insurance After-tax", "Cancer Insurance Pre-tax",
    "Child Support", "Child Support 2", "Creditor Garnishment",
    "Critical Illness After-tax", "Critical Illness Pre-tax", "Dental After-tax",
    "Dental Pre-tax", "Earned Wage Access", "Federal Tax Lien",
    "Gap Medical Pre-tax", "Group Term Life", "Health Cues Claim",
    "Health Cues Premium", "Health Reimbursement Arrangement (HRA) Pre-tax",
    "Health Savings Account(HSA) Pre-tax", "Hearing Insurance Pre-tax",
    "Hospital Indemnity After-tax", "Hospital Indemnity Pre-tax", "Loan",
    "Med Claim Reimbursement", "Med Plus Premium", "Medical After-tax",
    "Medical Pre-tax", "Overpayment", "Pet Insurance After-tax", "Reverse / Reissue",
    "Roth 401k", "Roth IRA", "Spousal Support Order", "State Tax Lien",
    "Student Loan", "Supplemental Life", "Supplemental Medical After-tax",
    "Supplemental Medical Pre-tax", "Vision After-tax", "Vision Pre-tax",
    "Voluntary AD&D After-tax", "Voluntary AD&D Pre-tax",
    "Voluntary Life Child After-tax", "Voluntary Life Child Pre-tax",
    "Voluntary Life Employee After-tax", "Voluntary Life Employee Pre-tax",
    "Voluntary Life Spouse After-tax", "Voluntary Life Spouse Pre-tax",
    "Voluntary LTD After-tax", "Voluntary STD After-tax", "Wellness Pre-tax",
    "Whole Life Insurance After-tax", "Other",
]

# Method options EXACTLY as they appear in UZIO's Method dropdown.
METHOD_FIXED = "Fixed $"
METHOD_PCT_GROSS = "% of Gross Pay"
METHOD_PCT_DISPOSABLE = "% of Disposable Net Pay"

# Masters that always use % of Disposable Net Pay (CCPA-limited garnishments).
DISPOSABLE_INCOME_MASTERS = {
    "creditor garnishment",
    "federal tax lien",
    "state tax lien",
}

# Masters that always use Fixed $ (court-ordered fixed amounts, loans, etc.).
FIXED_DOLLAR_MASTERS = {
    "child support",
    "child support 2",
    "spousal support order",
    "401(k) loan",
    "loan",
    "advance",
    "overpayment",
    "reverse / reissue",
}

# Masters that always use % of Gross Pay: 401k / Roth 401k deferrals are
# percent-of-pay elections in UZIO (the 401(k) Loan master stays Fixed $).
PCT_GROSS_MASTERS = {
    "401k",
    "roth 401k",
}

# Benefit-type masters show the "Auto-Sync from Uzio Benefits" radio (and, being
# benefit types, also Track arrears = Yes / Arrears Processing = Total Amount).
BENEFIT_TYPE_KEYWORDS = (
    "dental", "medical", "vision", "voluntary life", "critical illness",
    "accident insurance", "cancer insurance", "hospital indemnity",
    "voluntary std", "ad&d", "whole life",
)

# Masters whose "Assign to all employees" is FORCED Yes + disabled by UZIO.
ASSIGN_ALL_LOCKED_MASTERS = {
    "med claim reimbursement",
    "med plus premium",
    "health cues",
    "health cues premium",
}

DEFAULT_DEDUCTION_SCHEDULE = "Every Paycheck"
DEFAULT_ASSIGN_TO_ALL = "No"
ARREARS_PROCESSING_TOTAL = "Total Amount"
AUTOSYNC_NA = "N/A"
# Real masters auto-fill + LOCK the W-2 Box (emit a marker so the automation
# skips it); Other/unmapped leave it editable and get "Not Required".
DEFAULT_W2_BOX = "Not Required"
W2_BOX_LOCKED = "(Auto-filled by UZIO - do not set)"

# UZIO defaults auto-created on every client — never re-create.
DEFAULT_UZIO_DEDUCTIONS_TO_SKIP = {
    "earned wage access",
}

# Authoritative ADP deduction-code map (per implementor, 2026-06-11). A tuple
# means a Pre-tax/After-tax PAIRED family — the empirical verdict picks the
# variant; a plain string is a fixed master. Rohit extends this list over time.
ADP_DEDUCTION_CODE_SEEDS = {
    "ACC": "Voluntary AD&D After-tax",
    "CIL": ("Critical Illness Pre-tax", "Critical Illness After-tax"),
    "HOS": ("Hospital Indemnity Pre-tax", "Hospital Indemnity After-tax"),
    "PAC": "Earned Wage Access",
    "STD": "Voluntary STD After-tax",
    "SPT": "Voluntary STD After-tax",
    "VEE": "Voluntary Life Employee After-tax",
    "WEE": "Whole Life Insurance After-tax",
    "VCH": ("Voluntary Life Child Pre-tax", "Voluntary Life Child After-tax"),
    "VSP": ("Voluntary Life Spouse Pre-tax", "Voluntary Life Spouse After-tax"),
    "DEN": ("Dental Pre-tax", "Dental After-tax"),
    "MED": ("Medical Pre-tax", "Medical After-tax"),
    "VIS": ("Vision Pre-tax", "Vision After-tax"),
    "REV": "Reverse / Reissue",
    "ITR": "Other",
}

# Keyword families with Pre-tax/After-tax PAIRED masters — the empirical
# Pre/Post Tax verdict picks the variant. Keywords are matched on WORD
# BOUNDARIES (so "dental" can never match inside "acciDENTAL"). ADP truncates
# descriptions, hence the short variants ("critical ill", "hos indemnity").
TAX_PAIRED_KEYWORD_MASTERS = [
    ("dental",            "Dental Pre-tax",            "Dental After-tax"),
    ("vision",            "Vision Pre-tax",            "Vision After-tax"),
    ("medical",           "Medical Pre-tax",           "Medical After-tax"),
    ("critical illness",  "Critical Illness Pre-tax",  "Critical Illness After-tax"),
    ("critical ill",      "Critical Illness Pre-tax",  "Critical Illness After-tax"),
    ("hospital indemnity", "Hospital Indemnity Pre-tax", "Hospital Indemnity After-tax"),
    ("hos indemnity",     "Hospital Indemnity Pre-tax", "Hospital Indemnity After-tax"),
    ("indemnity",         "Hospital Indemnity Pre-tax", "Hospital Indemnity After-tax"),
    ("cancer",            "Cancer Insurance Pre-tax",  "Cancer Insurance After-tax"),
]

# Keyword families with a SINGLE master regardless of tax treatment. ORDER
# MATTERS: "spousal" before the generic support rules; "401k loan" before
# "401k". ADP files often label child support as just "SUPPORT" (75-SUPPORT),
# so the bare "support" keyword maps there too.
SINGLE_KEYWORD_MASTERS = [
    ("spousal",       "Spousal Support Order"),
    ("child support", "Child Support"),
    ("support order", "Child Support"),
    ("support",       "Child Support"),
    ("garnish",       "Creditor Garnishment"),
    ("garnishment",   "Creditor Garnishment"),
    ("levy",          "Creditor Garnishment"),
    ("payactiv",      "Earned Wage Access"),
    ("earned wage",   "Earned Wage Access"),
    ("roth",          "Roth 401k"),
    ("401k loan",     "401(k) Loan"),
    ("401(k) loan",   "401(k) Loan"),
    ("401k",          "401k"),
    ("hsa",           "Health Savings Account(HSA) Pre-tax"),
]


def _kw_in(desc, kw):
    """Word-boundary keyword test: `kw` must not be embedded inside a longer
    alphabetic run ("dental" ⊄ "accidental"; "401k" ✓ in "401k % retire").
    A trailing DIGIT is allowed — ADP appends sequence numbers to descriptions
    ("401K LOAN1", "SUPPORT2"), which must still match their keyword."""
    return re.search(
        rf"(?<![a-z0-9]){re.escape(kw)}(?![a-z])", desc) is not None

# Canonical column order for the Deductions tab + the UI dataframe.
DEDUCTION_OUTPUT_COLUMNS = [
    "Type Code", "Type Description", "Pre/Post Tax",
    "UZIO Master Deductions List", "UZIO Deduction Type", "UZIO Deduction Name",
    "UZIO Method", "Amount per pay", "Auto-Sync from Uzio Benefits",
    "Assign to all employees", "Deduction Schedule", "Track arrears",
    "Arrears Processing Method", "W-2 Box",
]


def map_adp_to_uzio_master(type_code, type_description, tax_treatment=""):
    """Return the UZIO Master Deductions List value for an ADP deduction.

    Order: authoritative code seeds (ADP_DEDUCTION_CODE_SEEDS; a tuple picks the
    Pre/After-tax variant by the empirical verdict) → keyword inference on the
    description (word-boundary, tax-aware) → NEEDS_REVIEW sentinel."""
    is_post = (tax_treatment or "").strip().lower() == "post tax"

    code = (type_code or "").strip().upper()
    seeded = ADP_DEDUCTION_CODE_SEEDS.get(code)
    if seeded is not None:
        if isinstance(seeded, tuple):
            return seeded[1] if is_post else seeded[0]
        return seeded

    desc = (type_description or "").strip().lower()

    # Reverse / Reissue needs BOTH words; "issu" tolerates ADP's truncation
    # ("REVERSE/REISSU") and still matches "issue"/"reissue"/"reissued".
    if "reverse" in desc and "issu" in desc:
        return "Reverse / Reissue"

    for kw, pre_master, post_master in TAX_PAIRED_KEYWORD_MASTERS:
        if _kw_in(desc, kw):
            return post_master if is_post else pre_master

    for kw, master in SINGLE_KEYWORD_MASTERS:
        if _kw_in(desc, kw):
            return master
    return NEEDS_REVIEW


def determine_method(uzio_master, type_description):
    """Garnishment masters → % of Disposable Net Pay; fixed-dollar masters →
    Fixed $; 401k / Roth 401k → % of Gross Pay; a "%" in the description →
    % of Gross Pay; else Fixed $."""
    master_l = (uzio_master or "").strip().lower()
    if master_l in DISPOSABLE_INCOME_MASTERS:
        return METHOD_PCT_DISPOSABLE
    if master_l in FIXED_DOLLAR_MASTERS:
        return METHOD_FIXED
    if master_l in PCT_GROSS_MASTERS:
        return METHOD_PCT_GROSS
    if "%" in (type_description or ""):
        return METHOD_PCT_GROSS
    return METHOD_FIXED


def is_benefit_type(uzio_master):
    """True if this master shows the Auto-Sync radio."""
    m = (uzio_master or "").strip().lower()
    return any(kw in m for kw in BENEFIT_TYPE_KEYWORDS)


def is_assign_all_locked(uzio_master):
    """True if UZIO forces 'Assign to all employees' = Yes for this master."""
    return (uzio_master or "").strip().lower() in ASSIGN_ALL_LOCKED_MASTERS


def adp_deductions_to_setup_rows(ded_rows):
    """Adapt classify_deductions_pretax() output → (Type Code, Type Description,
    Pre/Post Tax) rows for the UZIO enrichment layer.

    Splits the ADP label ("73-GARNISHMENT" → 73 / "Garnishment") and converts the
    empirical verdict (pre_tax / post_tax) to the UZIO wording. Discovery stats
    ride along under `_` keys for the UI."""
    rows = []
    for r in ded_rows or []:
        label = str(r.get("Code", "")).strip()
        if not label:
            continue
        code, desc = _split_code_desc(label)
        desc = _smart_title(desc) if desc else code
        verdict = str(r.get("Verdict", "")).strip().lower()
        if verdict == "pre_tax":
            tax = "Pre-tax"
        elif verdict == "post_tax":
            tax = "Post Tax"
        else:
            tax = "Unknown"
        rows.append({
            "Type Code": code,
            "Type Description": desc,
            "Pre/Post Tax": tax,
            "_Source Label": label,
            "_Source Column": str(r.get("Source Column", "") or label).strip(),
            "_Total $": r.get("Total $"),
            "_Employees": r.get("Employees"),
            "_Confidence": r.get("Confidence"),
            "_Flavor": r.get("Pre-Tax Flavor"),
        })
    return _dedupe_setup_rows(rows)


def filter_default_uzio_deductions(rows):
    """Split deduction rows into (kept, skipped): UZIO auto-creates a small set
    on every client (currently Earned Wage Access); never re-create them. A row
    is skipped when its description matches directly OR when its code/keywords
    resolve to the Earned Wage Access master (e.g. PAC-PAYACTIV)."""
    if not rows:
        return [], []
    kept, skipped = [], []
    for r in rows:
        td = (r.get("Type Description") or "").strip().lower()
        is_default = td in DEFAULT_UZIO_DEDUCTIONS_TO_SKIP
        if not is_default:
            master = map_adp_to_uzio_master(
                r.get("Type Code"), r.get("Type Description"), r.get("Pre/Post Tax"))
            is_default = master == "Earned Wage Access"
        (skipped if is_default else kept).append(r)
    return kept, skipped


def enrich_deductions_for_uzio(rows, auto_sync_map=None, master_override_map=None):
    """Add every UZIO "Add Deduction" form field to the adapted deduction rows
    (see DEDUCTION_OUTPUT_COLUMNS). `auto_sync_map` / `master_override_map` are
    optional {autosync_row_key -> value} maps from the UI controls. The internal
    "Is Benefit Type" bool drives the Auto-Sync toggle bar and is not exported."""
    auto_sync_map = auto_sync_map or {}
    master_override_map = master_override_map or {}
    out = []
    for r in rows:
        master = map_adp_to_uzio_master(
            r.get("Type Code"), r.get("Type Description"), r.get("Pre/Post Tax")
        )
        ovr = master_override_map.get(
            autosync_row_key(r.get("Type Code"), r.get("Type Description")))
        if ovr:
            master = ovr
        method = determine_method(master, r.get("Type Description"))
        benefit = is_benefit_type(master) and master != NEEDS_REVIEW

        if benefit:
            auto_sync = auto_sync_map.get(
                autosync_row_key(r.get("Type Code"), r.get("Type Description")), "No"
            )
            track_arrears = "Yes"
            arrears_method = ARREARS_PROCESSING_TOTAL
        else:
            auto_sync = AUTOSYNC_NA
            track_arrears = "No"
            arrears_method = ""

        assign_all = "Yes" if is_assign_all_locked(master) else DEFAULT_ASSIGN_TO_ALL

        # "Other"/unmapped: master is generic, real name lives in Deduction Name;
        # UZIO locks "Other" to Post Tax.
        if master in (MASTER_OTHER, NEEDS_REVIEW):
            ded_name = r.get("Type Description", "")
        else:
            ded_name = master
        ded_type = "Post Tax" if master == MASTER_OTHER else r.get("Pre/Post Tax", "")

        w2_box = DEFAULT_W2_BOX if master in (MASTER_OTHER, NEEDS_REVIEW) else W2_BOX_LOCKED

        out.append({
            **r,
            "UZIO Master Deductions List": master,
            "UZIO Deduction Type": ded_type,
            "UZIO Deduction Name": ded_name,
            "UZIO Method": method,
            "Amount per pay": "",
            "Auto-Sync from Uzio Benefits": auto_sync,
            "Assign to all employees": assign_all,
            "Deduction Schedule": DEFAULT_DEDUCTION_SCHEDULE,
            "Track arrears": track_arrears,
            "Arrears Processing Method": arrears_method,
            "W-2 Box": w2_box,
            "Is Benefit Type": benefit,
        })
    return out


# ─────────────────────────────────────────────────────────────────────────────
# UZIO "Add Contribution" form mapping config
#
# Ported from the Paycom helper. The Contribution form: free-text name, optional
# link to a company deduction, Method, optional limits, W-2 Box, assign-to-all.
#
# ADP source: employer match money lives in MEMO columns whose label contains
# the word MATCH ("MEMO : K-401K MATCH", "MEMO : K-ROTH 401K MATCH"). Other memo
# columns (PR BEGIN DATE, PTO balances, 401K MAX EL, ZONE 1 TAX...) are
# informational and are NOT contributions.
# ─────────────────────────────────────────────────────────────────────────────
CONTRIB_METHOD_FORMULA = "Formula"
CONTRIB_METHOD_FIXED = "Fixed $"
# Tiered employer-match formula for 401k / Roth 401k matches (safe harbor):
# 100% of the first 1%, then 50% of the next 4%.
CONTRIB_FORMULA_TIERS = [
    (100, 1),
    (50, 4),
]
CONTRIB_DEFAULT_ASSIGN_TO_ALL = "No"
CONTRIB_DEFAULT_W2_BOX = "Not Required"
CONTRIB_LINK_NONE = "(none - do not link)"

CONTRIBUTION_OUTPUT_COLUMNS = [
    "Type Code", "Type Description", "Contribution Name",
    "Link to Company Deduction", "Linked Deduction", "Method",
    "Formula", "Monthly Limit", "Annual Limit", "W-2 Box",
    "Assign to all employees",
]

_MEMO_PREFIX_RE = re.compile(r"^MEMO\s*:?\s*", re.IGNORECASE)
_MATCH_WORD_RE = re.compile(r"\bmatch\b", re.IGNORECASE)
# "Roth:MEMO : N" — the Roth-match column the Sanity-Check 401k/Roth split carves
# out of a combined memo column. Recognized as the Roth 401k match.
_ROTH_PREFIX_RE = re.compile(r"^\s*roth\s*:", re.IGNORECASE)

# Value-based employer-match test thresholds. A real match is a SMALL slice of
# each person's gross (never the whole paycheck), and it only appears where the
# employee has a 401k/Roth deferral.
_CONTRIB_RATIO_MIN = 0.003     # 0.3% of gross
_CONTRIB_RATIO_MAX = 0.10      # 10% of gross  (gross-equal memos like "J" ~ 1.0)
_CONTRIB_COOCCUR_MIN = 0.60    # >=60% of the column's rows also have a deferral

# Labels marking a memo as INFORMATIONAL (hour balances, dates, wage trackers
# like Federal Qualified Overtime). The value-based detector must never flag
# these as employer match money — hour counts and OT-premium dollars routinely
# land in the small-%-of-gross band and co-occur with deferrals at any client
# with high 401k participation. Name-based detection ('MATCH' / 'Roth:') is
# deliberately NOT filtered: an explicit MATCH label always wins.
_INFO_MEMO_KEYWORD_RE = re.compile(
    r"(?<![A-Z0-9])("
    r"PTO|SICK|VAC|VACATION|HOLIDAY|HOURS|HRS|BAL|BALANCE|DATE|ZONE|TAX|MAX"
    r"|BONUS|OT|QOT|FDQOT|OVERTIME"
    r")(?![A-Z0-9])"
)


def _format_formula(tiers):
    """Human-readable formula string, e.g. '100% of first 1%; 50% of next 4%'."""
    parts = []
    for i, (match, upto) in enumerate(tiers):
        word = "first" if i == 0 else "next"
        parts.append(f"{match}% of {word} {upto}%")
    return "; ".join(parts)


def _memo_candidate_cols(df):
    """Every MEMO column plus any 'Roth:MEMO ...' split column (excluding the
    TOTAL MEMOS aggregate). These are the options the contribution picker shows."""
    out = []
    for c in df.columns:
        s = str(c).strip()
        u = s.upper()
        if u.startswith("TOTAL"):
            continue
        if u.startswith("MEMO") or _ROTH_PREFIX_RE.match(s):
            out.append(c)
    return out


def _retirement_deferral_cols(df):
    """Employee 401k / Roth deferral DEDUCTION columns (memos excluded) — used to
    confirm a memo's match money lines up with people who actually defer."""
    out = []
    for c in df.columns:
        u = str(c).strip().upper()
        if u.startswith("VOLUNTARY DEDUCTION") and ("401" in u or "ROTH" in u):
            out.append(c)
    return out


def _looks_like_employer_match(df, col, gross_col, deferral_cols):
    """Value-based test (name-blind): the column's dollars are a small % of each
    person's gross AND most of its rows co-occur with a 401k/Roth deferral.
    Rejects gross-equal memos (ratio ~1.0) and flat trackers (no deferral link)."""
    if gross_col is None:
        return False
    ratios, with_deferral, n = [], 0, 0
    for _, row in df.iterrows():
        v = _num(row.get(col))
        if v <= 0:
            continue
        g = _num(row.get(gross_col))
        if g <= 0:
            continue
        n += 1
        ratios.append(v / g)
        if any(_num(row.get(d)) > 0 for d in deferral_cols):
            with_deferral += 1
    if n < 3 or not ratios:
        return False
    ratios.sort()
    typical = ratios[len(ratios) // 2]   # median ratio to gross
    if not (_CONTRIB_RATIO_MIN <= typical <= _CONTRIB_RATIO_MAX):
        return False
    return (with_deferral / n) >= _CONTRIB_COOCCUR_MIN


def _memo_contribution_name(raw_col):
    """(kind, display_name) for a memo column. kind = 'roth' / '401k'. Uses a
    descriptive label when present (e.g. 'K-401K MATCH' -> '401K Match'); opaque
    codes (e.g. 'N') fall back to a generic name by kind."""
    raw = str(raw_col).strip()
    low = raw.lower()
    kind = "roth" if (_ROTH_PREFIX_RE.match(raw) or "roth" in low) else "401k"
    default_name = "Roth 401K Match" if kind == "roth" else "401K Match"
    label = _MEMO_PREFIX_RE.sub("", _ROTH_PREFIX_RE.sub("", raw)).strip()
    code, desc = _split_code_desc(label)
    desc = _smart_title(desc) if desc else ""
    if desc and any(k in desc.lower() for k in ("401", "roth", "match")):
        name = desc
    else:
        name = default_name
    return kind, (code or label or name), name


def build_memo_candidate_rows(df):
    """One setup-row per memo candidate column, each flagged `_Detected` if the
    three-tier auto-detection thinks it's an employer contribution:
      1. label contains 'MATCH'         (fast path, e.g. Happy Delivery)
      2. 'Roth:<col>' split column      (Roth 401k match)
      3. value-based match              (small % of gross + deferral co-occurrence)
    Empty (all-zero) columns are never auto-detected. The UI shows every
    candidate so the user can override the auto-pick."""
    gross_col = _find_col(df, ["GROSS PAY"])
    deferral_cols = _retirement_deferral_cols(df)
    cands = []
    for c in _memo_candidate_cols(df):
        raw = str(c).strip()
        label = _MEMO_PREFIX_RE.sub("", _ROTH_PREFIX_RE.sub("", raw)).strip()
        amounts = df[c].apply(_num)
        total = round(float(amounts.sum()), 2)
        emp = int((amounts != 0).sum())
        has_money = float(amounts.abs().sum()) > 0
        is_match_kw = bool(_MATCH_WORD_RE.search(label))
        is_roth_split = bool(_ROTH_PREFIX_RE.match(raw))
        is_info = bool(_INFO_MEMO_KEYWORD_RE.search(label.upper()))
        is_value = (not is_info) and _looks_like_employer_match(
            df, c, gross_col, deferral_cols)
        detected = has_money and (is_match_kw or is_roth_split or is_value)
        kind, code, name = _memo_contribution_name(raw)
        cands.append({
            "Type Code": code,
            "Type Description": name,
            "_Source Column": raw,
            "_Source Label": raw,
            "_Total $": total,
            "_Employees": emp,
            "_Detected": detected,
            "_Kind": kind,
        })

    # Opaque memo codes all fall back to the same generic name ("401K Match"),
    # which would create identically-named contributions in UZIO. Suffix the
    # source code onto duplicates so each contribution name is unique — and a
    # false positive that survives detection stays visible at a glance.
    name_counts = {}
    for c in cands:
        name_counts[c["Type Description"]] = name_counts.get(c["Type Description"], 0) + 1
    for c in cands:
        if name_counts[c["Type Description"]] > 1 and c["Type Code"]:
            c["Type Description"] = f'{c["Type Description"]} ({c["Type Code"]})'
    return cands


def build_contributions_from_memo_cols(candidates, selected_cols):
    """Pick the candidate rows for the user-selected memo columns (the manual
    override). Order follows `selected_cols`."""
    by_col = {c["_Source Column"]: c for c in (candidates or [])}
    return [by_col[col] for col in (selected_cols or []) if col in by_col]


def map_contribution_to_deduction(type_code, type_description, available_deduction_masters):
    """Default UZIO deduction to link a contribution to (only when that deduction
    is actually being created for this client): Roth match → Roth 401k,
    401k match → 401k, Medical ER → Medical Pre/After-tax. Else no default."""
    avail = {(_m or "").strip().lower() for _m in (available_deduction_masters or [])}
    desc = (type_description or "").strip().lower()

    def have(name):
        return name.strip().lower() in avail

    is_roth = "roth" in desc
    is_401k = "401k" in desc or "401(k)" in desc
    is_medical = "medical" in desc or "med er" in desc or "med memo" in desc

    if is_roth and have("Roth 401k"):
        return "Roth 401k"
    if is_401k and have("401k"):
        return "401k"
    if is_medical:
        if have("Medical Pre-tax"):
            return "Medical Pre-tax"
        if have("Medical After-tax"):
            return "Medical After-tax"
    return ""


def is_formula_contribution(type_description):
    """401k / Roth 401k matches use the tiered Formula; everything else Fixed $."""
    desc = (type_description or "").strip().lower()
    return "401k" in desc or "401(k)" in desc


def enrich_contributions_for_uzio(rows, link_map=None):
    """Add the UZIO Add-Contribution form fields to each contribution row.
    `link_map`: optional {autosync_row_key -> linked deduction display name}."""
    link_map = link_map or {}
    out = []
    for r in rows:
        key = autosync_row_key(r.get("Type Code"), r.get("Type Description"))
        linked = (link_map.get(key) or "").strip()
        if linked in ("", CONTRIB_LINK_NONE):
            linked, link_yn = "", "No"
        else:
            link_yn = "Yes"
        use_formula = is_formula_contribution(r.get("Type Description"))
        out.append({
            **r,
            "Contribution Name": r.get("Type Description", ""),
            "Link to Company Deduction": link_yn,
            "Linked Deduction": linked,
            "Method": CONTRIB_METHOD_FORMULA if use_formula else CONTRIB_METHOD_FIXED,
            "Formula": _format_formula(CONTRIB_FORMULA_TIERS) if use_formula else "",
            "Monthly Limit": "",
            "Annual Limit": "",
            "W-2 Box": CONTRIB_DEFAULT_W2_BOX,
            "Assign to all employees": CONTRIB_DEFAULT_ASSIGN_TO_ALL,
        })
    return out


# ─────────────────────────────────────────────────────────────────────────────
# Taxes: ADP → UZIO tax-catalog mapping
#
# ADP tax columns are structured "<NAME> - EMPLOYEE TAX" / "<NAME> - EMPLOYER
# TAX", so the side is explicit; the state(s) come from the WORKED IN STATE
# column. Mapping is NOT a creation step — UZIO already owns these taxes; we
# only produce the mapping file (Source tax → UZIO tax_code / unique_tax_id)
# against the bundled UZIO tax catalog (shared with the Paycom helper).
#
# Column → catalog TYPE token (4th dash-segment of unique_tax_id):
#   FEDERAL INCOME EE → FIT        MEDICARE EE/ER → MEDI / ER_MEDI
#   SOCIAL SECURITY EE/ER → FICA / ER_FICA        FUTA ER → ER_FUTA
#   WORKED IN STATE EE → SIT (one row per worked-in state)
#   SUI/SDI EE → SDI or SUI        SUI/SDI ER → ER_SUTA or ER_SDI
#   FAMILY LEAVE INS EE/ER → FLI / ER_FLI
#   MTA ER → ER_POP (NY MCTMT Employer Payroll Tax)
#   LIVED-IN LOCAL EE → local city/school — user confirms via the finder
# ─────────────────────────────────────────────────────────────────────────────

MAPPING_TAX_COLUMNS = [
    "Source Tax Code", "Source Tax Code Name", "Source Tax Code Description",
    "Uzio Tax Code", "Unique Tax ID", "Uzio Tax Code Description",
    "Uzio Sub-Tax Description",
]

# Single source of truth for the catalog — bundled once, in the Paycom app dir.
UZIO_TAX_CATALOG_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)), "..", "paycom", "uzio_tax_catalog.csv")


def load_uzio_tax_catalog(path=UZIO_TAX_CATALOG_PATH):
    """Read the bundled UZIO tax catalog into a list of dicts ([] if missing)."""
    try:
        cat_df = pd.read_csv(path, dtype=str, keep_default_na=False, encoding="utf-8-sig")
    except Exception:
        return []
    out = []
    for _, r in cat_df.iterrows():
        out.append({
            "state_abbreviation": str(r.get("state_abbreviation", "")).strip(),
            "tax_code": str(r.get("tax_code", "")).strip(),
            "unique_tax_id": str(r.get("unique_tax_id", "")).strip(),
            "tax_name": str(r.get("tax_name", "")).strip(),
            "sub_tax_desc": str(r.get("sub_tax_desc", "")).strip(),
        })
    return out


_TAX_CATALOG_CACHE = None


def _cached_tax_catalog():
    global _TAX_CATALOG_CACHE
    if _TAX_CATALOG_CACHE is None:
        _TAX_CATALOG_CACHE = load_uzio_tax_catalog()
    return _TAX_CATALOG_CACHE


def _uti_token(unique_tax_id):
    """The tax-TYPE token = 4th dash-segment of the unique_tax_id."""
    parts = (unique_tax_id or "").split("-")
    return parts[3].upper() if len(parts) >= 4 else ""


def _is_statewide(r):
    """True for the statewide catalog entry (place segment == '0000')."""
    p = (r.get("unique_tax_id") or "").split("-")
    return len(p) >= 3 and p[2] == "0000"


def _worked_in_states(df):
    """Distinct 2-letter states from the WORKED IN STATE column, most-common
    first (the dominant state leads)."""
    col = _find_col(df, ["WORKED IN STATE"])
    if not col:
        return []
    from collections import Counter
    c = Counter()
    for v in df[col].dropna().astype(str):
        s = v.strip().upper()
        if len(s) == 2 and s.isalpha():
            c[s] += 1
    return [s for s, _ in c.most_common()]


def adp_tax_key(row):
    """Stable widget/resolution key for one extracted ADP tax row."""
    return f"{row['Side']}||{row['Tax Column']}||{row.get('State') or ''}"


def extract_adp_taxes(df, tax_cols):
    """Turn the ADP tax columns into mapping rows. State-scoped taxes (SIT,
    SUI/SDI, FLI) expand to one row per worked-in state. Each row:
    Tax Column / Tax Name / Side / State ("FED", 2-letter, or None) /
    Tokens (catalog match keys, priority order) / Tier."""
    states = _worked_in_states(df) or []
    dominant = states[0] if states else None
    rows = []
    for c in tax_cols or []:
        raw = str(c).strip()
        u = raw.upper()
        employer = u.endswith("EMPLOYER TAX")
        side = "Employer" if employer else "Employee"
        name = re.sub(r"\s*-\s*EMPLOY(EE|ER)\s+TAX$", "", u).strip()
        n = name.lower()

        def add(state, tokens, tier, sub_pref=None):
            rows.append({
                "Tax Column": raw, "Tax Name": _smart_title(name), "Side": side,
                "State": state, "Tokens": tokens, "Tier": tier,
                "Sub Pref": sub_pref,
            })

        if "federal income" in n:
            add("FED", ["FIT"], "federal")
        elif "medicare" in n:
            add("FED", ["ER_MEDI"] if employer else ["MEDI"], "federal")
        elif "social security" in n:
            add("FED", ["ER_FICA"] if employer else ["FICA"], "federal")
        elif "futa" in n:
            add("FED", ["ER_FUTA"], "federal")
        elif "worked in state" in n:
            for st_ab in (states or [None]):
                add(st_ab, ["SIT"], "state")
        elif "sui" in n or "sdi" in n:
            toks = ["ER_SUTA", "ER_SDI"] if employer else ["SDI", "SUI"]
            for st_ab in (states or [None]):
                add(st_ab, toks, "state")
        elif "medical leave" in n:
            # PFML states (e.g. MA) split Family vs Medical leave into separate
            # ADP columns that share the FLI catalog tax but differ only by
            # sub_tax_desc — carry a preference so the matcher picks the right one.
            toks = ["ER_FLI"] if employer else ["FLI"]
            for st_ab in (states or [None]):
                add(st_ab, toks, "state", sub_pref="medical")
        elif "family leave" in n or "fli" in n:
            toks = ["ER_FLI"] if employer else ["FLI"]
            for st_ab in (states or [None]):
                add(st_ab, toks, "state", sub_pref="family")
        elif "mta" in n:
            add("NY", ["ER_POP"], "state")
        elif "local" in n:
            add(dominant, [], "local")
        else:
            add(dominant, [], "unknown")
    return rows


def adp_tax_best_match(row, catalog):
    """Best-guess catalog row for an extracted ADP tax (None when the tool
    shouldn't guess — locals/unknowns, or no statewide row for any token)."""
    state, tokens = row.get("State"), row.get("Tokens") or []
    if not state or not tokens:
        return None
    pool = [r for r in catalog if r["state_abbreviation"].upper() == state]
    if state == "NY" and "ER_POP" in tokens:
        # MTA: several NY ER_POP taxes exist — MCTMT is the one ADP calls MTA.
        mctmt = [r for r in pool if _uti_token(r["unique_tax_id"]) == "ER_POP"
                 and "mctmt" in r["tax_name"].lower()]
        if mctmt:
            return mctmt[0]
    sub_pref = row.get("Sub Pref")
    for tok in tokens:
        hits = [r for r in pool if _uti_token(r["unique_tax_id"]) == tok]
        if not hits:
            continue
        # FLI families (e.g. MA Paid FMLY & Medical Leave) carry multiple catalog
        # rows that differ only by sub_tax_desc — narrow to the family/medical
        # variant this column is about. Single-row FLI states (NY, NJ, CT...) have
        # nothing to narrow, so they're unaffected.
        if sub_pref and tok in ("FLI", "ER_FLI"):
            kw = "MEDICAL" if sub_pref == "medical" else "FAMILY"
            pref = [r for r in hits if kw in (r.get("sub_tax_desc") or "").upper()]
            if pref:
                hits = pref
        statewide = next((r for r in hits if _is_statewide(r)), None)
        return statewide or hits[0]
    return None


def build_adp_tax_mapping_rows(taxes, resolved):
    """Emit mapping-file rows (MAPPING_TAX_COLUMNS order). `resolved` maps
    adp_tax_key(row) -> chosen catalog row (or None)."""
    rows = []
    for t in taxes:
        m = resolved.get(adp_tax_key(t))
        rows.append({
            "Source Tax Code": "",
            "Source Tax Code Name": t["Tax Column"],
            # ADP doesn't use this column — keep the header for template parity
            # but leave it blank (per the client tax-mapping sample).
            "Source Tax Code Description": "",
            "Uzio Tax Code": (m or {}).get("tax_code", ""),
            "Unique Tax ID": (m or {}).get("unique_tax_id", ""),
            "Uzio Tax Code Description": (m or {}).get("tax_name", ""),
            "Uzio Sub-Tax Description": (m or {}).get("sub_tax_desc", ""),
        })
    return rows


def _csv_bytes(rows, columns):
    return pd.DataFrame(rows, columns=columns).to_csv(index=False).encode("utf-8-sig")


# ─────────────────────────────────────────────────────────────────────────────
# API "Mapping files" (Source name -> UZIO name)
#
# Same contract as the Paycom helper: 4-column CSVs translating the source name
# into the exact UZIO name; only the two *Name* columns are filled (the API
# matches on name). Header spelling is verbatim from the client templates.
#
# ADP difference: source items are COLUMN HEADERS, not row values — so the
# "Source ... Code Name" carries the original column label VERBATIM
# ("ADDITIONAL EARNINGS  : BNH-BONUS HOURS", "MEMO : K-401K MATCH"), exactly as
# the API will read it from the source file.
#
# Skipped UZIO defaults are INCLUDED (earnings: Regular Wage / Overtime /
# Lookback bonus...; deductions: Earned Wage Access) because the API still
# uploads data against those existing, system-created UZIO items.
# ─────────────────────────────────────────────────────────────────────────────

MAPPING_EARNING_COLUMNS = [
    "Source Earning Code", "Source Earning Code Name",
    "Uzio Earning Code", "Uzio Earning Code Name",
]
MAPPING_DEDUCTION_COLUMNS = [
    "Source Deduction Code", "Source Deduction Code Name",
    "Uzio Deduction Code", "Uzio Deduction Code Name",
]
MAPPING_CONTRIBUTION_COLUMNS = [
    "Source Contribution Code", "Source Contribution Code Name",
    "Uzio Contribution Code", "Uzio Contribution Code Name",
]


def _src_col_name(r):
    return r.get("_Source Column") or r.get("_Source Label") or r.get("Type Description", "")


def build_earnings_mapping_rows(enriched_earnings, skipped_earnings):
    """One row per earning — created ones AND skipped UZIO defaults. Uzio name =
    the created Earning Name (kept) or the UZIO default name (skipped)."""
    rows = []
    for r in enriched_earnings or []:
        rows.append({
            "Source Earning Code": "",
            "Source Earning Code Name": _src_col_name(r),
            "Uzio Earning Code": "",
            "Uzio Earning Code Name": r.get("Earning Name", r.get("Type Description", "")),
        })
    for r in skipped_earnings or []:
        rows.append({
            "Source Earning Code": "",
            "Source Earning Code Name": _src_col_name(r),
            "Uzio Earning Code": "",
            "Uzio Earning Code Name": r.get("UZIO Default Earning", ""),
        })
    return rows


def build_deductions_mapping_rows(enriched_deds, skipped_deds=None):
    """One row per created deduction (Uzio name = UZIO Deduction Name), plus the
    skipped UZIO defaults (PAC-Payactiv → Earned Wage Access) so payroll data can
    still upload against them."""
    rows = []
    for r in enriched_deds or []:
        rows.append({
            "Source Deduction Code": "",
            "Source Deduction Code Name": _src_col_name(r),
            "Uzio Deduction Code": "",
            "Uzio Deduction Code Name": r.get("UZIO Deduction Name", r.get("Type Description", "")),
        })
    for r in skipped_deds or []:
        rows.append({
            "Source Deduction Code": "",
            "Source Deduction Code Name": _src_col_name(r),
            "Uzio Deduction Code": "",
            "Uzio Deduction Code Name": "Earned Wage Access",
        })
    return rows


def build_contributions_mapping_rows(enriched_contribs):
    """One row per created contribution (Uzio name = Contribution Name)."""
    rows = []
    for r in enriched_contribs or []:
        rows.append({
            "Source Contribution Code": "",
            "Source Contribution Code Name": _src_col_name(r),
            "Uzio Contribution Code": "",
            "Uzio Contribution Code Name": r.get("Contribution Name", r.get("Type Description", "")),
        })
    return rows


def _mapping_csv_bytes(rows, columns):
    """Plain UTF-8 CSV bytes (NO BOM — the API matches headers/names verbatim
    and a BOM would corrupt the first header). Values pass through untouched."""
    df = pd.DataFrame(rows, columns=columns) if rows else pd.DataFrame(columns=columns)
    return df.to_csv(index=False).encode("utf-8")


# HTML/JS for a SINGLE button that downloads all four mapping CSVs as separate
# files (st.download_button can only emit one file per click). Rendered via
# st.components.v1.html — its iframe sandbox includes `allow-downloads`, so the
# programmatic anchor clicks below actually download. `__FILES__` is replaced
# with a JSON array of [filename, base64-csv] pairs. The 400ms stagger lets the
# browser accept the multi-file download (Chrome prompts once to allow it).
MAPPING_DOWNLOAD_HTML = """
<style>
  .dlwrap { font-family: "Source Sans Pro", system-ui, -apple-system, sans-serif; }
  .dl-btn {
    background:#ff4b4b; color:#fff; border:1px solid #ff4b4b; border-radius:.5rem;
    padding:.55rem 1rem; font-size:1rem; font-weight:600; cursor:pointer;
  }
  .dl-btn:hover { background:#e53935; border-color:#e53935; }
  .dl-btn:active { transform:translateY(1px); }
  .dl-note { color:#808495; font-size:.8rem; margin-top:.5rem; }
</style>
<div class="dlwrap">
  <button class="dl-btn" id="dlAllMaps">&#128229; Download all mapping CSVs</button>
  <div class="dl-note">Saves the mapping files (earnings, deductions, contributions, taxes)
    as separate files. The first time, Chrome may ask to
    &ldquo;allow downloading multiple files&rdquo; &mdash; choose <b>Allow</b>.</div>
</div>
<script>
  (function () {
    var FILES = __FILES__;
    var btn = document.getElementById('dlAllMaps');
    btn.addEventListener('click', function () {
      FILES.forEach(function (f, i) {
        setTimeout(function () {
          var a = document.createElement('a');
          a.href = 'data:text/csv;charset=utf-8;base64,' + f[1];
          a.download = f[0];
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
        }, i * 400);
      });
    });
  })();
</script>
"""


def build_setup_xlsx(enriched_earnings, enriched_deductions, enriched_contributions):
    """UZIO Setup workbook: Earnings + Deductions + Contributions tabs (canonical
    column orders), frozen header rows, Pre/Post Tax color-coded on the
    Deductions tab (green = Pre-tax, red = Post Tax, grey = Unknown)."""
    import xlsxwriter  # noqa: F401  (local import keeps the module importable)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#1F4E78", "font_color": "white",
            "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        cell_fmt = wb.add_format({"border": 1, "valign": "vcenter"})
        pre_fmt = wb.add_format({"border": 1, "valign": "vcenter",
                                 "bg_color": "#C6EFCE", "font_color": "#006100"})
        post_fmt = wb.add_format({"border": 1, "valign": "vcenter",
                                  "bg_color": "#FFC7CE", "font_color": "#9C0006"})
        unk_fmt = wb.add_format({"border": 1, "valign": "vcenter",
                                 "bg_color": "#D9D9D9", "font_color": "#3F3F3F"})

        def _write_tab(sheet_name, rows, columns, color_col=None):
            df = pd.DataFrame(rows, columns=columns)
            df.to_excel(writer, sheet_name=sheet_name, index=False,
                        startrow=1, header=False)
            ws = writer.sheets[sheet_name]
            for col_idx, col_name in enumerate(columns):
                ws.write(0, col_idx, col_name, header_fmt)
                body = df[col_name].astype(str) if not df.empty else pd.Series([], dtype=str)
                width = max([len(col_name)] + [len(v) for v in body]) + 2
                ws.set_column(col_idx, col_idx, min(max(width, 10), 42), cell_fmt)
            if color_col and color_col in columns and not df.empty:
                ci = columns.index(color_col)
                for ri, val in enumerate(df[color_col].astype(str), start=1):
                    v = val.strip().lower()
                    fmt = pre_fmt if v == "pre-tax" else post_fmt if v == "post tax" else unk_fmt
                    ws.write(ri, ci, val, fmt)
            ws.freeze_panes(1, 0)

        _write_tab("Earnings", enriched_earnings, EARNING_OUTPUT_COLUMNS)
        _write_tab("Deductions", enriched_deductions, DEDUCTION_OUTPUT_COLUMNS,
                   color_col="Pre/Post Tax")
        _write_tab("Contributions", enriched_contributions, CONTRIBUTION_OUTPUT_COLUMNS)
    buf.seek(0)
    return buf.getvalue()


# ---------- pre/post-tax classifier ----------

def _row_gap(row, total_earn_col, taxable_col):
    return _num(row.get(total_earn_col)) - _num(row.get(taxable_col))


def _subset_sum_match(amounts, target, tol=0.02):
    n = len(amounts)
    if n == 0:
        return []
    matches = []
    for r in range(1, n + 1):
        for combo in combinations(range(n), r):
            s = sum(amounts[i] for i in combo)
            if abs(s - target) <= tol:
                matches.append(combo)
    return matches


def _name_heuristic(col):
    u = str(col).upper()
    if any(t in u for t in ("SUPPORT", "GARN", "GARNISH", "LEVY", "LIEN", "CHILD")):
        return "post_tax", "garnishment", [], "name_heuristic"
    if any(t in u for t in ("ADVANCE", "ADV-", "LOAN", "REPAY", "TAPCHECK", "DAILY")):
        return "post_tax", "advance_or_loan", [], "name_heuristic"
    if any(t in u for t in ("REVERSE", "REV-", "REISSU")):
        return "post_tax", "corrective", [], "name_heuristic"
    if any(t in u for t in ("ROTH",)):
        return "post_tax", "roth", [], "name_heuristic"
    if any(t in u for t in ("MEDICAL", "MED-", "DENTAL", "DEN-", "VISION", "VIS-",
                            "HSA", "FSA")):
        return "pre_tax", "section_125", ["FIT", "FICA", "MEDI", "SIT"], "name_heuristic"
    if CONTRIB_PATTERN.search(u):
        return "pre_tax", "401k_traditional", ["FIT", "SIT"], "name_heuristic"
    return "post_tax", "default_unknown", [], "name_heuristic"


def classify_deductions_pretax(
    df, ded_cols, total_earn_col, fit_taxable_col, fica_taxable_col,
    medi_taxable_col, sit_taxable_col, tol=0.02, max_subset=8,
):
    proven = {c: {"FIT": False, "FICA": False, "MEDI": False, "SIT": False}
              for c in ded_cols}
    sample = {c: [] for c in ded_cols}

    def _try_axis(taxable_col, key):
        if taxable_col is None:
            return
        for _, row in df.iterrows():
            gap = _row_gap(row, total_earn_col, taxable_col)
            if gap <= tol:
                continue
            present = [(c, _num(row.get(c))) for c in ded_cols if _num(row.get(c)) > 0]
            if not present or len(present) > max_subset:
                continue
            cols = [c for c, _ in present]
            amts = [a for _, a in present]
            for combo in _subset_sum_match(amts, gap, tol):
                for i in combo:
                    proven[cols[i]][key] = True
                if key == "FIT":
                    eid = row.get("ASSOCIATE ID") or row.get("Associate ID")
                    for i in combo:
                        if len(sample[cols[i]]) < 3:
                            sample[cols[i]].append({
                                "associate": str(eid) if eid is not None else "",
                                "gap_fit": round(gap, 2),
                                "subset": [cols[j] for j in combo],
                                "subset_sum": round(sum(amts[j] for j in combo), 2),
                            })

    _try_axis(fit_taxable_col, "FIT")
    _try_axis(fica_taxable_col, "FICA")
    _try_axis(medi_taxable_col, "MEDI")
    _try_axis(sit_taxable_col, "SIT")

    rows = []
    for c in ded_cols:
        amounts = df[c].apply(_num)
        total = float(amounts.sum())
        emp_count = int((amounts != 0).sum())
        p = proven[c]
        if p["FIT"] and p["FICA"] and p["MEDI"]:
            verdict = "pre_tax"; flavor = "section_125"
            pre_taxes = ["FIT", "FICA", "MEDI"] + (["SIT"] if p["SIT"] else [])
        elif p["FIT"] and p["SIT"] and not p["FICA"]:
            verdict = "pre_tax"; flavor = "401k_traditional"; pre_taxes = ["FIT", "SIT"]
        elif p["FIT"] and not (p["FICA"] or p["MEDI"]):
            verdict = "pre_tax"; flavor = "pretax_unknown"; pre_taxes = ["FIT"]
        elif p["FIT"] or p["FICA"] or p["MEDI"] or p["SIT"]:
            verdict = "pre_tax"; flavor = "mixed_unusual"
            pre_taxes = [k for k in ("FIT", "FICA", "MEDI", "SIT") if p[k]]
        else:
            verdict = "post_tax"; flavor = ""; pre_taxes = []

        if emp_count == 0:
            verdict, flavor, pre_taxes, confidence = _name_heuristic(c)
        else:
            confidence = "empirical_subset_sum"

        code = _strip_prefix(c, [DED_PREFIX])
        is_contrib = bool(CONTRIB_PATTERN.search(code.upper()))
        rows.append({
            "Source Column": str(c).strip(), "Code": code,
            "Total $": round(total, 2), "Employees": emp_count,
            "Verdict": verdict, "Pre-Tax Of": ",".join(pre_taxes),
            "Pre-Tax Flavor": flavor, "Confidence": confidence,
            "Sample": "; ".join(
                f"{s['associate']}: gap={s['gap_fit']}, subset_sum={s['subset_sum']}"
                for s in sample[c][:2]
            ),
            "_is_contribution": is_contrib,
        })
    return rows


# ---------- bonus classifier (FLSA) ----------
#
# Per-paycheck dual-prediction test. For each testable row we compute BOTH
# possible OT amounts and see which one ADP actually paid:
#   Regular Rate   = REGULAR EARNINGS / REGULAR HOURS
#   Worked Hours   = REGULAR + OVERTIME + DOUBLE OVERTIME + TRAINING hours
#                    (PTO / Station Closure hours are paid-not-worked: excluded)
#   discretionary prediction     = 1.5 x Regular Rate x OT hours
#   Blended Rate   = (Regular Rate x Worked Hours + Bonus) / Worked Hours
#   non-disc OT rate = Blended / 2 + Regular Rate
#   non-disc prediction          = non-disc OT rate x OT hours
# The row's verdict is whichever prediction the actual OVERTIME EARNINGS
# matches (<=1% off). Both far, or predictions indistinguishable (tiny
# bonus) -> row is 'unclear' and doesn't count as evidence.
#
# Every bonus column gets its OWN verdict, two ways:
#   individual  - rows where ONLY that bonus is present (clean evidence)
#   combined    - rows where the bonus appears at all, tested with the SUM
#                 of all bonuses (matches how ADP computes the rate)
# agree -> confirmed; disagree -> needs_review; one side indeterminate ->
# the other side's verdict (noted). Any new column whose name contains
# BONUS / a BN* code is picked up automatically.
#
# CONSOLIDATED FILES ARE REFUSED: on a 1-row-per-employee quarter file the
# bonus may have been earned in a period with no overtime (and vice versa),
# so any verdict would be meaningless — the tool says so instead of guessing.

_BONUS_ROW_TOL = 0.01          # actual must be within 1% of a prediction
_BONUS_MIN_GAP = 0.005         # predictions closer than 0.5% are indistinguishable


def _bonus_verdict_from_rows(row_verdicts):
    """FLSA aggregation: any positive proof of rate inflation is conclusive."""
    tested = [v for v in row_verdicts if v in ("discretionary", "non_discretionary")]
    if not tested:
        return "indeterminate"
    if any(v == "non_discretionary" for v in tested):
        return "non_discretionary"
    return "discretionary"


def classify_bonus(df, earn_cols):
    reg_e = _find_col(df, ["REGULAR EARNINGS"])
    reg_h = _find_col(df, ["REGULAR HOURS"])
    ot_e = _find_col(df, ["OVERTIME EARNINGS"])
    ot_h = _find_col(df, ["OVERTIME HOURS"])
    dot_h = _find_col(df, ["DOUBLE OVERTIME HOURS", "DOUBLE OT HOURS", "DBL OT HOURS"])

    # Training hours count as worked; PTO / Station Closure do not.
    extra_worked_h = []
    for c in df.columns:
        u = str(c).strip().upper()
        if "HOURS" not in u or not u.startswith("ADDITIONAL HOURS"):
            continue
        if any(x in u for x in ("PTO", "PAID TIME", "STATION", "CLOSURE")):
            continue
        if "TRA" in u or "TRAINING" in u:
            extra_worked_h.append(c)

    bonus_cols = []
    for c in earn_cols:
        u = str(c).upper()
        code = _strip_prefix(c, EARN_PREFIXES).upper()
        if "BONUS" in u or re.search(r"\bBN[A-Z0-9]*\b", code) or code.startswith("BN"):
            if "BACKUP" in u or code.startswith("BCK"):
                continue
            bonus_cols.append(c)

    def _empty(verdict, reason, consolidated=False, skipped=0):
        return {
            "verdict": verdict, "reason": reason,
            "bonus_columns_found": [str(c) for c in bonus_cols],
            "rows_tested": 0, "discretionary_rows": 0, "non_discretionary_rows": 0,
            "unclear_rows": 0, "samples": [], "per_bonus": [],
            "consolidated_file": consolidated, "aggregated_rows_skipped": skipped,
        }

    if not bonus_cols or not (reg_e and reg_h and ot_e and ot_h):
        return _empty("indeterminate", "Missing bonus / overtime columns to test")

    # ── ROW-level aggregated-row guard ──────────────────────────────────────
    # A verdict may only come from a genuine single-pay-period row. Uploads can
    # MIX consolidated and per-pay-period files (they get concatenated), so a
    # file-level check is not enough: each row's own PERIOD span decides.
    # Span > 35 days => aggregated (quarter/consolidated) row => never evidence.
    pbeg = _find_col(df, ["PERIOD BEGINNING DATE"])
    pend = _find_col(df, ["PERIOD ENDING DATE"])
    agg_mask = None
    if pbeg and pend:
        _b = pd.to_datetime(df[pbeg], errors="coerce")
        _e = pd.to_datetime(df[pend], errors="coerce")
        _span = (_e - _b).dt.days
        agg_mask = _span.gt(35).fillna(False)
    aggregated_skipped = int(agg_mask.sum()) if agg_mask is not None else 0

    def _is_aggregated(idx):
        return agg_mask is not None and bool(agg_mask.loc[idx])

    def _row_test(r, bonus_amt):
        """('discretionary'|'non_discretionary'|'unclear'|None, details).
        None = row not testable with this bonus amount."""
        re_v = _num(r.get(reg_e)); rh_v = _num(r.get(reg_h))
        oe_v = _num(r.get(ot_e)); oh_v = _num(r.get(ot_h))
        if bonus_amt <= 0 or oh_v <= 0 or rh_v <= 0 or re_v <= 0 or oe_v <= 0:
            return None, None
        worked = rh_v + oh_v + (_num(r.get(dot_h)) if dot_h else 0.0) \
            + sum(_num(r.get(c)) for c in extra_worked_h)
        if worked <= 0:
            return None, None
        regular_rate = re_v / rh_v
        disc_pred = 1.5 * regular_rate * oh_v
        blended = regular_rate + bonus_amt / worked
        nd_pred = (blended / 2 + regular_rate) * oh_v
        if disc_pred <= 0:
            return None, None
        d_disc = abs(oe_v - disc_pred) / disc_pred
        d_nd = abs(oe_v - nd_pred) / nd_pred if nd_pred > 0 else 1.0
        details = {
            "regular_rate": round(regular_rate, 4), "worked_hours": round(worked, 2),
            "blended_rate": round(blended, 4),
            "predicted_ot_discretionary": round(disc_pred, 2),
            "predicted_ot_non_disc": round(nd_pred, 2),
            "actual_ot": round(oe_v, 2), "bonus_amt": round(bonus_amt, 2),
        }
        if abs(nd_pred - disc_pred) / disc_pred < _BONUS_MIN_GAP:
            return "unclear", details        # predictions indistinguishable
        if min(d_disc, d_nd) > _BONUS_ROW_TOL:
            return "unclear", details        # actual matches neither
        return ("discretionary" if d_disc <= d_nd else "non_discretionary"), details

    # ── Combined pass (sum of all bonuses, every non-aggregated bonus row) ──
    combined_verdicts = []
    samples = []
    unclear_rows = 0
    for idx, r in df.iterrows():
        if _is_aggregated(idx):
            continue
        total_bonus = sum(_num(r.get(c)) for c in bonus_cols)
        v, det = _row_test(r, total_bonus)
        if v is None:
            continue
        combined_verdicts.append(v)
        if v == "unclear":
            unclear_rows += 1
        if len(samples) < 5:
            eid = r.get("ASSOCIATE ID") or r.get("Associate ID")
            parts = []
            for c in bonus_cols:
                amt = _num(r.get(c))
                if amt > 0:
                    parts.append(f"{_strip_prefix(c, EARN_PREFIXES)} = {amt:,.2f}")
            samples.append({"associate": str(eid) if eid is not None else "",
                            "verdict_row": v,
                            "bonus_breakdown": "; ".join(parts), **det})

    overall = _bonus_verdict_from_rows(combined_verdicts)
    tested = sum(1 for v in combined_verdicts if v != "unclear")
    n_nd = sum(1 for v in combined_verdicts if v == "non_discretionary")
    n_d = sum(1 for v in combined_verdicts if v == "discretionary")

    # ── Per-bonus pass: individual (solo rows) + combined (rows featuring it) ──
    per_bonus = []
    for bc in bonus_cols:
        others = [c for c in bonus_cols if c != bc]
        ind_verdicts, comb_verdicts = [], []
        for idx, r in df.iterrows():
            if _is_aggregated(idx):
                continue
            own = _num(r.get(bc))
            if own <= 0:
                continue
            others_amt = sum(_num(r.get(c)) for c in others)
            total = own + others_amt
            if others_amt <= 0:
                v, _d = _row_test(r, own)
                if v is not None:
                    ind_verdicts.append(v)
            v2, _d2 = _row_test(r, total)
            if v2 is not None:
                comb_verdicts.append(v2)

        ind = _bonus_verdict_from_rows(ind_verdicts)
        comb = _bonus_verdict_from_rows(comb_verdicts)
        if ind == comb:
            final, note = ind, ("confirmed by both individual and combined tests"
                                if ind != "indeterminate" else "no testable rows")
        elif ind == "indeterminate":
            final, note = comb, "combined evidence only (bonus never appears alone)"
        elif comb == "indeterminate":
            final, note = ind, "individual evidence only"
        else:
            final, note = "needs_review", (
                f"individual test says {ind}, combined says {comb} — review manually")

        code = _strip_prefix(bc, EARN_PREFIXES)
        code = code.split("-", 1)[0].strip().upper() if "-" in code else code.strip().upper()
        per_bonus.append({
            "column": str(bc), "code": code,
            "individual_verdict": ind,
            "individual_rows": sum(1 for v in ind_verdicts if v != "unclear"),
            "combined_verdict": comb,
            "combined_rows": sum(1 for v in comb_verdicts if v != "unclear"),
            "final_verdict": final, "note": note,
        })

    skipped_note = (f" ({aggregated_skipped} aggregated/consolidated row(s) with a "
                    f">35-day period span were excluded from evidence.)"
                    if aggregated_skipped else "")
    all_aggregated = tested == 0 and aggregated_skipped > 0

    if all_aggregated:
        reason = (f"All rows look CONSOLIDATED (period span > 35 days, "
                  f"{aggregated_skipped} row(s) skipped). The bonus may sit in a "
                  f"pay period with no overtime (and vice versa), so aggregate math "
                  f"would be wrong. Upload the per-pay-period file for bonus "
                  f"classification.")
    elif tested == 0:
        reason = ("No single-pay-period row had bonus + overtime together "
                  "(or all rows were unclear)." + skipped_note)
    elif overall == "non_discretionary":
        reason = (f"{n_nd} of {tested} testable single-pay-period rows match the "
                  f"blended-rate (non-discretionary) OT prediction — any positive "
                  f"proof is conclusive under FLSA. See per-bonus verdicts."
                  + skipped_note)
    else:
        reason = (f"All {tested} testable single-pay-period rows match the plain "
                  f"1.5x (discretionary) OT prediction. See per-bonus verdicts."
                  + skipped_note)

    return {
        "verdict": overall, "reason": reason,
        "bonus_columns_found": [str(c) for c in bonus_cols],
        "rows_tested": tested,
        "discretionary_rows": n_d,
        "non_discretionary_rows": n_nd,
        "unclear_rows": unclear_rows,
        "samples": samples,
        "per_bonus": per_bonus,
        "consolidated_file": all_aggregated,
        "aggregated_rows_skipped": aggregated_skipped,
    }


# ---------- tax mapping ----------

TAX_TOKEN_MAP = {
    "FEDERAL INCOME - EMPLOYEE TAX":          ("FED", "FIT"),
    "MEDICARE - EMPLOYEE TAX":                ("FED", "MEDI"),
    "MEDICARE - EMPLOYER TAX":                ("FED", "ER_MEDI"),
    "SOCIAL SECURITY - EMPLOYEE TAX":         ("FED", "FICA"),
    "SOCIAL SECURITY - EMPLOYER TAX":         ("FED", "ER_FICA"),
    "FUTA - EMPLOYER TAX":                    ("FED", "ER_FUTA"),
    "WORKED IN STATE - EMPLOYEE TAX":         ("STATE", "SIT"),
    "SUI/SDI - EMPLOYEE TAX":                 ("STATE", "SDI"),
    "SUI/SDI - EMPLOYER TAX":                 ("STATE", "ER_SUTA"),
    "FAMILY LEAVE INSURANCE - EMPLOYEE TAX":  ("STATE", "FLI"),
}


def lookup_canonical_tax(master_df, state_abbr, type_code):
    if master_df is None:
        return None
    pat = re.compile(rf"^\d{{2}}-000-0000-{re.escape(type_code)}-000$")
    sub = master_df[master_df["state_abbreviation"].astype(str).str.upper()
                    == state_abbr.upper()]
    if sub.empty:
        return None
    sub2 = sub[sub["unique_tax_id"].astype(str).apply(lambda s: bool(pat.match(s)))]
    if sub2.empty:
        broad = master_df[
            (master_df["state_abbreviation"].astype(str).str.upper() == state_abbr.upper())
            & master_df["unique_tax_id"].astype(str).str.contains(f"-{type_code}-", regex=False)
        ]
        if broad.empty:
            return None
        primary = broad[broad["sub_tax_desc"].fillna("").astype(str).str.strip() == ""]
        return primary.iloc[0] if not primary.empty else broad.iloc[0]
    primary = sub2[sub2["sub_tax_desc"].fillna("").astype(str).str.strip() == ""]
    return primary.iloc[0] if not primary.empty else sub2.iloc[0]


def build_tax_mapping(df, tax_cols, master_df):
    state_col = _find_col(df, ["WORKED IN STATE", "Worked In State", "State"])
    states = []
    if state_col:
        for v in df[state_col].dropna().astype(str):
            s = v.strip().upper()
            if s and s not in states and len(s) == 2:
                states.append(s)
    if not states:
        states = ["NY"]

    out_rows = []; not_found = []
    for tcol in tax_cols:
        key = str(tcol).strip().upper()
        scope_type = TAX_TOKEN_MAP.get(key)
        if not scope_type:
            not_found.append({"tax_column": str(tcol), "reason": "no rule defined"})
            continue
        scope, type_code = scope_type
        targets = ["FED"] if scope == "FED" else states
        for st_code in targets:
            rec = lookup_canonical_tax(master_df, st_code, type_code)
            if rec is None:
                not_found.append({"tax_column": str(tcol),
                                  "reason": f"{st_code} {type_code} not in master"})
                continue
            out_rows.append({
                "Source Tax Code": "",
                "Source Tax Code Name": str(tcol),
                "Source Tax Code Description": "",
                "Uzio Tax Code": rec.get("tax_code", ""),
                "Unique Tax ID": rec.get("unique_tax_id", ""),
                "Uzio Tax Code Description": rec.get("tax_name", ""),
                "Uzio Sub-Tax Description": rec.get("sub_tax_desc", "") or "",
            })
    return out_rows, states, not_found


def tax_mapping_to_csv_bytes(rows):
    cols = ["Source Tax Code", "Source Tax Code Name", "Source Tax Code Description",
            "Uzio Tax Code", "Unique Tax ID", "Uzio Tax Code Description",
            "Uzio Sub-Tax Description"]
    return pd.DataFrame(rows, columns=cols).to_csv(index=False).encode("utf-8")


# ---------- orchestrator ----------

def run_setup_helper(adp_files, master_csv_file=None):
    """Run the analysis over one ADP file or a list of them (frames are
    concatenated). Returns (results_dict_of_lists, tax_csv_bytes)."""
    if not isinstance(adp_files, (list, tuple)):
        adp_files = [adp_files]
    frames = []
    for f in adp_files:
        d, _, _ = read_input_file(f)
        frames.append(d.reset_index(drop=True))
    df = pd.concat(frames, ignore_index=True) if len(frames) > 1 else frames[0]

    cats = categorize_columns(df)
    if len(frames) > 1:
        # pd.concat appends columns unique to later files at the END of the
        # combined frame — outside the positional TOTAL HOURS..TOTAL EARNINGS
        # window — so detect earnings/hours per source frame and union them
        # (first-seen order preserved).
        earn_cols, hour_cols = [], []
        for fr in frames:
            c = categorize_columns(fr)
            for col in c["earnings"]:
                if col not in earn_cols:
                    earn_cols.append(col)
            for col in c["hours"]:
                if col not in hour_cols:
                    hour_cols.append(col)
        cats["earnings"], cats["hours"] = earn_cols, hour_cols
    earn_rows = build_earnings_catalog(df, cats["earnings"], cats["hours"])

    total_earn_col = _find_col(df, ["TOTAL EARNINGS"]) or _find_col(df, ["GROSS PAY"])
    fit_taxable = _find_col(df, ["FEDERAL INCOME - EMPLOYEE TAXABLE"])
    fica_taxable = _find_col(df, ["SOCIAL SECURITY - EMPLOYEE TAXABLE"])
    medi_taxable = _find_col(df, ["MEDICARE - EMPLOYEE TAXABLE"])
    sit_taxable = _find_col(df, ["WORKED IN STATE - EMPLOYEE TAXABLE"])

    ded_rows = classify_deductions_pretax(
        df, cats["deductions"], total_earn_col,
        fit_taxable, fica_taxable, medi_taxable, sit_taxable,
    )
    contributions = [r for r in ded_rows if r.pop("_is_contribution", False)]
    deductions = [r for r in ded_rows if r not in contributions]
    for r in deductions:
        r.pop("_is_contribution", None)

    tax_rows = [{
        "Source Column": str(c).strip(),
        "Total $": round(float(df[c].apply(_num).sum()), 2),
        "Employees": int((df[c].apply(_num) != 0).sum()),
    } for c in cats["taxes"]]

    master_df = None
    if master_csv_file is not None:
        master_csv_file.seek(0)
        master_df = pd.read_csv(master_csv_file, dtype=str)
    tax_mapping_rows, states, missing = build_tax_mapping(df, cats["taxes"], master_df)

    bonus_info = classify_bonus(df, cats["earnings"])

    summary = [
        {"Metric": "Rows in file", "Value": len(df)},
        {"Metric": "Distinct earnings codes", "Value": len(earn_rows)},
        {"Metric": "Distinct contribution codes", "Value": len(contributions)},
        {"Metric": "Distinct deduction codes", "Value": len(deductions)},
        {"Metric": "Distinct tax columns", "Value": len(tax_rows)},
        {"Metric": "States detected", "Value": ", ".join(states) if states else "(none)"},
        {"Metric": "Tax mapping rows produced", "Value": len(tax_mapping_rows)},
        {"Metric": "Tax mapping rows missing from master", "Value": len(missing)},
        {"Metric": "Bonus classification verdict", "Value": bonus_info["verdict"]},
        {"Metric": "Bonus rows tested", "Value": bonus_info["rows_tested"]},
        {"Metric": "Bonus columns detected",
         "Value": ", ".join(bonus_info["bonus_columns_found"]) or "(none)"},
    ]

    bonus_rows = [{
        "Verdict": bonus_info["verdict"], "Reason": bonus_info["reason"],
        "Rows Tested": bonus_info["rows_tested"],
        "Discretionary Rows": bonus_info["discretionary_rows"],
        "Non-Discretionary Rows": bonus_info["non_discretionary_rows"],
        "Unclear Rows": bonus_info.get("unclear_rows", 0),
        "Consolidated File": "Yes" if bonus_info.get("consolidated_file") else "No",
        "Aggregated Rows Skipped": bonus_info.get("aggregated_rows_skipped", 0),
        "Bonus Columns": ", ".join(bonus_info["bonus_columns_found"]),
    }]

    results = {
        "Summary": summary,
        "Earnings_Codes": earn_rows,
        "Contributions": contributions,
        "Deductions": deductions,
        "Taxes_Discovered": tax_rows,
        "Tax_Mapping": tax_mapping_rows,
        "Tax_Mapping_Missing": missing,
        "States_Detected": [{"State": s} for s in states],
        "Bonus_Classification": bonus_rows,
        "Bonus_Sample_Rows": bonus_info["samples"],
        "Bonus_Per_Column": bonus_info.get("per_bonus", []),
        # New UZIO-setup sections: every MEMO candidate column (with auto-detect
        # flags) for the Contributions picker, and the structured tax rows for
        # the catalog-based mapping UI.
        "Memo_Candidates": build_memo_candidate_rows(df),
        "ADP_Taxes": extract_adp_taxes(df, cats["taxes"]),
    }
    csv_bytes = tax_mapping_to_csv_bytes(tax_mapping_rows)
    return results, csv_bytes


def _results_to_xlsx_bytes(results):
    """Three-tab simplified xlsx that answers exactly:
      Tab 1 - What to set up in Uzio (Earnings | Contributions | Deductions)
      Tab 2 - Bonus discretionary or non-discretionary (verdict + one example)
      Tab 3 - Each deduction: pre-tax or post-tax + plain-English why
    Nothing else. Tax mapping is offered as a separate CSV download.
    """
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        wb = writer.book
        header_fmt = wb.add_format({
            "bold": True, "bg_color": "#1F4E78", "font_color": "white",
            "border": 1, "align": "left", "valign": "vcenter",
        })
        wrap_fmt = wb.add_format({"valign": "top", "text_wrap": True})
        verdict_pre = wb.add_format({
            "bold": True, "bg_color": "#C6EFCE", "font_color": "#006100",
            "align": "center", "valign": "vcenter",
        })
        verdict_post = wb.add_format({
            "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006",
            "align": "center", "valign": "vcenter",
        })
        verdict_nondisc = wb.add_format({
            "bold": True, "bg_color": "#FFC7CE", "font_color": "#9C0006",
            "align": "left", "valign": "vcenter", "font_size": 14,
        })
        verdict_disc = wb.add_format({
            "bold": True, "bg_color": "#C6EFCE", "font_color": "#006100",
            "align": "left", "valign": "vcenter", "font_size": 14,
        })

        # ---- Tab 1: What to Set Up ----
        earn_codes = [r["Code"] for r in results["Earnings_Codes"]]
        contrib_codes = [r["Code"] for r in results["Contributions"]]
        ded_codes = [r["Code"] for r in results["Deductions"]]
        max_n = max(len(earn_codes), len(contrib_codes), len(ded_codes), 1)
        rows1 = []
        for i in range(max_n):
            rows1.append({
                "Earnings": earn_codes[i] if i < len(earn_codes) else "",
                "Contributions": contrib_codes[i] if i < len(contrib_codes) else "",
                "Deductions": ded_codes[i] if i < len(ded_codes) else "",
            })
        df1 = pd.DataFrame(rows1)
        df1.to_excel(writer, sheet_name="1. What to Set Up", index=False)
        ws1 = writer.sheets["1. What to Set Up"]
        ws1.set_column("A:A", 32); ws1.set_column("B:B", 24); ws1.set_column("C:C", 32)
        for i, col in enumerate(df1.columns):
            ws1.write(0, i, col, header_fmt)
        ws1.set_row(0, 24)

        # ---- Tab 2: Bonus ----
        bonus = results["Bonus_Classification"][0]
        sample = _pick_bonus_example(results["Bonus_Sample_Rows"], bonus["Verdict"])
        verdict_label = bonus["Verdict"].upper().replace("_", "-")
        rows2 = [
            ("Verdict", verdict_label),
            ("Reason", bonus["Reason"]),
            ("Bonus codes detected in file", bonus["Bonus Columns"]),
            ("Rows that had both bonus AND overtime", bonus["Rows Tested"]),
            ("    of which discretionary", bonus["Discretionary Rows"]),
            ("    of which non-discretionary", bonus["Non-Discretionary Rows"]),
        ]
        if sample:
            rows2 += [
                ("", ""),
                ("---- Example row that proves the verdict ----", ""),
                ("Associate ID", sample["associate"]),
                ("Regular rate ($/hr)", f"${sample['regular_rate']}"),
                ("Worked hours (Reg + OT + Dbl OT + Training)", sample["worked_hours"]),
                ("Blended rate with bonus ($/hr)", f"${sample['blended_rate']}"),
                ("Predicted OT if DISCRETIONARY (1.5 x regular)", f"${sample['predicted_ot_discretionary']:,}"),
                ("Predicted OT if NON-DISCRETIONARY (blended)", f"${sample['predicted_ot_non_disc']:,}"),
                ("Actual OT paid in this row", f"${sample['actual_ot']:,}"),
                ("Bonus paid in this row", f"${sample['bonus_amt']:,}"),
                ("", ""),
                ("Plain-English explanation",
                    "Actual OT matches the BLENDED-rate prediction (bonus spread over all "
                    "worked hours, half-time premium on the blended rate) => the bonus was "
                    "rolled into the regular rate before computing OT => NON-DISCRETIONARY (FLSA)."
                    if bonus["Verdict"] == "non_discretionary" else
                    "Actual OT matches the plain 1.5 x regular-rate prediction => the bonus "
                    "did NOT inflate the regular rate basis => DISCRETIONARY."
                    if bonus["Verdict"] == "discretionary" else
                    bonus["Reason"]),
            ]
        per_bonus = results.get("Bonus_Per_Column") or []
        if per_bonus:
            rows2 += [("", ""), ("---- Per-bonus verdicts ----", "")]
            for pb in per_bonus:
                rows2.append((
                    pb["column"],
                    f"{str(pb['final_verdict']).upper().replace('_', '-')} — {pb['note']} "
                    f"(individual: {pb['individual_verdict']} on {pb['individual_rows']} row(s); "
                    f"combined: {pb['combined_verdict']} on {pb['combined_rows']} row(s))",
                ))
        df2 = pd.DataFrame(rows2, columns=["Field", "Value"])
        df2.to_excel(writer, sheet_name="2. Bonus Verdict", index=False)
        ws2 = writer.sheets["2. Bonus Verdict"]
        ws2.set_column("A:A", 44); ws2.set_column("B:B", 80, wrap_fmt)
        for i, col in enumerate(df2.columns):
            ws2.write(0, i, col, header_fmt)
        ws2.set_row(0, 24)
        # Highlight the verdict cell (row 1, col B)
        if bonus["Verdict"] == "non_discretionary":
            ws2.write(1, 1, verdict_label, verdict_nondisc)
        elif bonus["Verdict"] == "discretionary":
            ws2.write(1, 1, verdict_label, verdict_disc)
        ws2.set_row(1, 28)

        # ---- Tab 3: Deductions Pre/Post-Tax ----
        rows3 = []
        for r in results["Contributions"] + results["Deductions"]:
            rows3.append({
                "Code": r["Code"],
                "Verdict": "PRE-TAX" if r["Verdict"] == "pre_tax" else "POST-TAX",
                "Why": _deduction_reason(r),
            })
        if not rows3:
            rows3 = [{"Code": "(none)", "Verdict": "",
                      "Why": "No voluntary deductions or contributions found in this file."}]
        df3 = pd.DataFrame(rows3)
        df3.to_excel(writer, sheet_name="3. Pre-Tax vs Post-Tax", index=False)
        ws3 = writer.sheets["3. Pre-Tax vs Post-Tax"]
        ws3.set_column("A:A", 26); ws3.set_column("B:B", 14)
        ws3.set_column("C:C", 110, wrap_fmt)
        for i, col in enumerate(df3.columns):
            ws3.write(0, i, col, header_fmt)
        ws3.set_row(0, 24)
        # Color the verdict cells
        for ri, r in enumerate(rows3, start=1):
            if r["Verdict"] == "PRE-TAX":
                ws3.write(ri, 1, "PRE-TAX", verdict_pre)
            elif r["Verdict"] == "POST-TAX":
                ws3.write(ri, 1, "POST-TAX", verdict_post)
            ws3.set_row(ri, 30)

    return buf.getvalue()


def _pick_bonus_example(samples, verdict):
    """Pick the single most illustrative row for the chosen verdict."""
    if not samples:
        return None

    def gap(s):
        return abs(s["actual_ot"] - s["predicted_ot_discretionary"])

    if verdict == "non_discretionary":
        candidates = [s for s in samples if s["verdict_row"] == "non_discretionary"]
        return max(candidates, key=gap) if candidates else samples[0]
    if verdict == "discretionary":
        candidates = [s for s in samples if s["verdict_row"] == "discretionary"]
        return min(candidates, key=gap) if candidates else samples[0]
    return samples[0]


def _deduction_reason(row):
    """Plain-English reason for a deduction's pre/post-tax verdict."""
    verdict = row["Verdict"]
    flavor = row.get("Pre-Tax Flavor", "")
    sample = row.get("Sample", "")
    if verdict == "post_tax":
        return "No row in the file showed taxable wages being reduced by this deduction's amount, so it does NOT shrink the tax base."
    if flavor == "section_125":
        first = sample.split(";")[0].strip() if sample else ""
        return ("Reduces FIT, FICA, Medicare, and state-income taxable wages by the deduction amount — Section 125 cafeteria plan." +
                (f" Example row: {first}" if first else ""))
    if flavor == "401k_traditional":
        first = sample.split(";")[0].strip() if sample else ""
        return ("Reduces FIT and state-income taxable wages but NOT FICA/Medicare — traditional 401(k)/403(b) pattern." +
                (f" Example row: {first}" if first else ""))
    if flavor == "pretax_unknown":
        return "Reduces FIT taxable wages only (no FICA/Medicare reduction observed)."
    if flavor == "mixed_unusual":
        return f"Pre-tax for: {row.get('Pre-Tax Of', '')} (unusual mix — review)."
    return "Pre-tax (see sample column for the matching row)."


# ---------- Streamlit UI ----------

def _render_name_editor(title, rows, name_field, key_prefix, caption=None):
    """Collapsible accordion of editable UZIO-name text boxes — one per row.

    Writes the chosen name back into each row[name_field] IN PLACE so the edit
    flows into the on-screen table AND the setup Excel. Each box pre-fills with
    the computed name and keeps following it on reruns until the user types a
    custom value, after which the user's value sticks."""
    if not rows:
        return
    with st.expander(title, expanded=False):
        if caption:
            st.caption(caption)
        for r in rows:
            code = r.get("Type Code", "")
            td = r.get("Type Description", "")
            default = str(r.get(name_field, td) or td)
            wkey = f"{key_prefix}::{autosync_row_key(code, td)}"
            defkey = wkey + "::__def__"
            prev_def = st.session_state.get(defkey)
            if wkey not in st.session_state or st.session_state.get(wkey) == prev_def:
                st.session_state[wkey] = default
            st.session_state[defkey] = default
            nm_col, ed_col = st.columns([1, 1])
            with nm_col:
                st.markdown(f"**{code}** — {td}" if code else f"**{td}**")
            with ed_col:
                st.text_input(
                    f"{key_prefix}-{wkey}", key=wkey, label_visibility="collapsed",
                )
            val = st.session_state.get(wkey, default)
            r[name_field] = val.strip() if isinstance(val, str) and val.strip() else default


def _render_earning_setup_section(results, src_name):
    """UZIO Earning Setup — mirrors the Paycom helper: filters UZIO-default
    earnings, lets the user (a) flag bonuses non-discretionary, (b) override the
    Earning Type, (c) rename each earning, then renders the enriched table and a
    download for the single-tab Earning Setup .xlsx."""
    setup_rows = adp_earnings_to_setup_rows(results.get("Earnings_Codes", []))
    kept_earnings, skipped_earnings = filter_default_uzio_earnings(setup_rows)

    st.markdown("## UZIO Earning Setup")
    st.caption(
        f"Discovered from the earnings block between `TOTAL HOURS` and "
        f"`TOTAL EARNINGS` (REGULAR / OVERTIME + additional earnings). Codes are "
        f"split off the labels (`BNH-BONUS HOURS` → **BNH** / *Bonus Hours*).  \n"
        f"**{len(kept_earnings)} earning(s) to create**, "
        f"**{len(skipped_earnings)} skipped** as UZIO defaults (auto-created). "
        f"Earning Type drives Hourly/Taxability; `{EARNING_TYPE_OTHER}` rows stay "
        f"editable in the Excel."
    )
    if skipped_earnings:
        skip_lines = "; ".join(
            (f"{r['Type Code']}-" if r.get("Type Code") else "") + r["Type Description"]
            + f" → {r['UZIO Default Earning']}"
            for r in skipped_earnings
        )
        st.info(
            f"**Skipped {len(skipped_earnings)} UZIO default earning(s)** "
            f"(already auto-created by UZIO; will not be re-created): `{skip_lines}`"
        )

    unknown_codes = [r for r in kept_earnings if r.get("_Unknown Code")]
    if unknown_codes:
        st.warning(
            f"**{len(unknown_codes)} code-only earning(s) not in the catalog**: "
            f"`{'; '.join(r['Type Code'] for r in unknown_codes)}` — the file gives "
            "no description and the earning-code catalog doesn't know the code yet, "
            "so the code itself is used as the name below. Once you know what it "
            "means, add it to `apps/adp/adp_earning_code_catalog.json` (or rename it "
            "in the editors here) — known pairs from every analyzed file are saved "
            "there automatically."
        )

    def _effective_earn_type(code, td):
        return st.session_state.get(
            f"adp_ppsh_earntype::{autosync_row_key(code, td)}"
        ) or map_to_earning_type(code, td)

    # ── "Is the earning Non Discretionary?" toggles (Bonus types only) ──────
    bonus_rows = [
        (r["Type Code"], r["Type Description"])
        for r in kept_earnings
        if _effective_earn_type(r["Type Code"], r["Type Description"]) == "Bonus"
    ]
    btd_counts = {}
    for _, td in bonus_rows:
        btd_counts[td] = btd_counts.get(td, 0) + 1

    include_in_ot_map = {}
    if bonus_rows:
        # Pre-fill each toggle from the per-bonus FLSA verdict (code-matched).
        pb_by_code = {str(p.get("code", "")).upper(): p.get("final_verdict")
                      for p in (results.get("Bonus_Per_Column") or [])}
        review_codes = [c for c, v in pb_by_code.items() if v == "needs_review"]

        st.markdown("### Is the earning Non Discretionary?")
        st.caption(
            f"**{len(bonus_rows)} bonus earning(s)**. Toggle **ON = Non-Discretionary** "
            "(included in the overtime rate → *Include Bonus in Overtime = Yes*); "
            "**OFF = Discretionary** (*= No*). Defaults are pre-filled from the "
            "per-bonus FLSA test above — override if you know better."
        )
        if review_codes:
            st.warning(
                "FLSA test could not settle these bonus code(s) — individual and "
                "combined evidence disagree, please decide manually: "
                + ", ".join(f"`{c}`" for c in review_codes)
            )
        cols = st.columns(2)
        for i, (code, td) in enumerate(bonus_rows):
            label = f"{td} ({code})" if (code and btd_counts[td] > 1) else td
            default_on = pb_by_code.get(str(code).upper()) == "non_discretionary"
            tkey = f"adp_ppsh_nondisc::{autosync_row_key(code, td)}"
            dkey = f"{tkey}::computed_default"
            # Streamlit remembers the widget's state across reruns, so a verdict
            # that CHANGES (e.g. new files uploaded mid-session) would otherwise
            # keep showing the stale default. Follow the new verdict — but only
            # when the user hasn't manually overridden the old one.
            prev_default = st.session_state.get(dkey)
            if tkey not in st.session_state:
                st.session_state[tkey] = default_on
            elif prev_default is not None and prev_default != default_on \
                    and st.session_state[tkey] == prev_default:
                st.session_state[tkey] = default_on
            st.session_state[dkey] = default_on
            with cols[i % 2]:
                on = st.toggle(label, key=tkey)
            include_in_ot_map[autosync_row_key(code, td)] = "Yes" if on else "No"

    # ── Map earnings to an Earning Type (collapsed accordion; all earnings) ──
    earning_type_override_map = {}
    if kept_earnings:
        et_opts = list(UZIO_EARNING_TYPES)
        with st.expander("Map earnings to an earning type", expanded=False):
            st.caption(
                f"**{len(kept_earnings)} earning(s)**. Each dropdown defaults to the "
                "Earning Type the tool inferred — change it to override. "
                f"`{EARNING_TYPE_OTHER}` exposes the editable fields in UZIO."
            )
            for code, td in [(r["Type Code"], r["Type Description"]) for r in kept_earnings]:
                computed = map_to_earning_type(code, td)
                idx = et_opts.index(computed) if computed in et_opts else et_opts.index(EARNING_TYPE_OTHER)
                nm_col, pk_col = st.columns([1, 1])
                with nm_col:
                    st.markdown(f"**{code}** — {td}" if code else f"**{td}**")
                with pk_col:
                    sel = st.selectbox(
                        f"EarnType {td}", options=et_opts, index=idx,
                        key=f"adp_ppsh_earntype::{autosync_row_key(code, td)}",
                        label_visibility="collapsed",
                    )
                if sel != computed:
                    earning_type_override_map[autosync_row_key(code, td)] = sel

    enriched_earnings = enrich_earnings_for_uzio(
        kept_earnings, include_in_ot_map=include_in_ot_map,
        earning_type_override_map=earning_type_override_map,
    )

    _render_name_editor(
        "✏️ Edit UZIO Earning names (optional)",
        enriched_earnings, "Earning Name", "adp_ppsh_earnname",
        caption=(
            "Set the exact name each earning should have in UZIO. Flows into the "
            "setup table below and the Earning Setup Excel. (Skipped UZIO-default "
            "earnings aren't listed; their names must match UZIO's seeded spellings.)"
        ),
    )

    st.markdown("### UZIO Earning Setup (all form fields)")
    if not enriched_earnings:
        st.caption("(no new earnings to create — all were UZIO defaults, or none found)")
    else:
        st.dataframe(
            pd.DataFrame(enriched_earnings, columns=EARNING_OUTPUT_COLUMNS),
            hide_index=True, use_container_width=True,
        )
    st.markdown("---")
    return enriched_earnings, skipped_earnings


def _render_deduction_setup_section(results, src_name):
    """UZIO Deduction Setup — mirrors the Paycom helper: filters UZIO-default
    deductions, shows the empirical pre/post-tax verdicts, lets the user (a) set
    Auto-Sync per benefit-type deduction, (b) override the inferred Master
    Deduction, (c) rename each deduction, then renders the enriched table.
    Returns the enriched rows for the combined setup workbook."""
    # Employee-paid retirement deferrals (401k / Roth / 401k Loan) are UZIO
    # deduction masters, so the legacy contribution split is folded back in here.
    # (Employer-side match money — true Contributions — will get its own section.)
    setup_rows = adp_deductions_to_setup_rows(
        results.get("Deductions", []) + results.get("Contributions", []))
    kept_deds, skipped_deds = filter_default_uzio_deductions(setup_rows)

    st.markdown("## UZIO Deduction Setup")
    st.caption(
        f"Discovered from the `VOLUNTARY DEDUCTION` columns; codes split off the "
        f"labels (`73-GARNISHMENT` → **73** / *Garnishment*). Employee retirement "
        f"deferrals (401k / Roth / 401k Loan / HSA) are included — in UZIO they're "
        f"deduction masters.  \n"
        f"**{len(kept_deds)} deduction(s) to create**. **Pre/Post Tax** is proven "
        f"empirically: the classifier finds a subset of each row's deductions "
        f"whose sum exactly explains the gap between Total Earnings and each "
        f"taxable wage base — one positive proof anywhere in the file locks the "
        f"verdict (`empirical_subset_sum`); rows with no proof fall back to name "
        f"heuristics."
    )
    if skipped_deds:
        skip_lines = "; ".join(
            (f"{r['Type Code']}-" if r.get("Type Code") else "") + r["Type Description"]
            for r in skipped_deds
        )
        st.info(
            f"**Skipped {len(skipped_deds)} default UZIO deduction(s)** "
            f"(auto-created by UZIO; will not be re-created): `{skip_lines}`"
        )

    if not kept_deds:
        st.caption("(no voluntary deductions found in the file(s))")
        st.markdown("---")
        return [], skipped_deds

    # ── Auto-Sync toggle bar (benefit-type masters only) ────────────────────
    st.markdown("### Auto-Sync from Uzio Benefits")
    benefit_rows = [
        (r["Type Code"], r["Type Description"])
        for r in enrich_deductions_for_uzio(kept_deds)
        if r.get("Is Benefit Type")
    ]
    td_counts = {}
    for _, td in benefit_rows:
        td_counts[td] = td_counts.get(td, 0) + 1

    def _widget_key(code, td):
        return f"adp_ppsh_autosync::{autosync_row_key(code, td)}"

    auto_sync_map = {}
    if not benefit_rows:
        st.caption(
            "No benefit-type deductions detected, so Auto-Sync is `N/A` for "
            "every deduction. (Benefit types: dental, medical, vision, voluntary "
            "life, critical illness, accident/cancer insurance, hospital "
            "indemnity, STD, AD&D.)"
        )
    else:
        def _on_select_all():
            v = st.session_state.get("adp_ppsh_autosync_all", False)
            for code, td in benefit_rows:
                st.session_state[_widget_key(code, td)] = v

        st.caption(
            f"**{len(benefit_rows)} benefit-type deduction(s)** can Auto-Sync. "
            "Toggle ON = Auto-Sync **Yes**, OFF = **No** (default OFF). "
            "**Select All** flips every one at once."
        )
        st.toggle("Select All", value=False, key="adp_ppsh_autosync_all",
                  on_change=_on_select_all)
        cols = st.columns(2)
        for i, (code, td) in enumerate(benefit_rows):
            label = f"{td} ({code})" if td_counts[td] > 1 else td
            with cols[i % 2]:
                on = st.toggle(label, value=False, key=_widget_key(code, td))
            auto_sync_map[autosync_row_key(code, td)] = "Yes" if on else "No"

    # ── Map deductions to a Master Deduction (collapsed accordion; ALL) ─────
    base_deds = enrich_deductions_for_uzio(kept_deds, auto_sync_map)
    needs_review = [r for r in base_deds
                    if r.get("UZIO Master Deductions List") == NEEDS_REVIEW]
    if needs_review:
        st.warning(
            f"**{len(needs_review)} deduction(s) couldn't be mapped to a UZIO "
            f"master** and are marked `{NEEDS_REVIEW}`: "
            f"`{'; '.join((r['Type Code'] + '-' if r['Type Code'] else '') + r['Type Description'] for r in needs_review)}`. "
            "Pick the right master in the accordion below — unmapped rows must "
            "not be shipped to the automation."
        )
    master_override_map = {}
    if base_deds:
        master_opts = [NEEDS_REVIEW] + UZIO_MASTER_DEDUCTIONS
        with st.expander("Map deductions to a master deduction", expanded=False):
            st.caption(
                f"**{len(base_deds)} deduction(s)** (UZIO auto-created defaults "
                "already excluded). Each dropdown defaults to the master the tool "
                f"inferred — change it to override. `{NEEDS_REVIEW}` means the "
                "tool couldn't match one."
            )
            for r in base_deds:
                code, td = r["Type Code"], r["Type Description"]
                computed = r.get("UZIO Master Deductions List", NEEDS_REVIEW)
                idx = master_opts.index(computed) if computed in master_opts else 0
                nm_col, pk_col = st.columns([1, 1])
                with nm_col:
                    st.markdown(f"**{code}** — {td}" if code else f"**{td}**")
                with pk_col:
                    sel = st.selectbox(
                        f"Master {code or td}", options=master_opts, index=idx,
                        key=f"adp_ppsh_master::{autosync_row_key(code, td)}",
                        label_visibility="collapsed",
                    )
                if sel != computed:
                    master_override_map[autosync_row_key(code, td)] = sel

    enriched_deds = enrich_deductions_for_uzio(
        kept_deds, auto_sync_map, master_override_map=master_override_map
    )

    # ── Edit UZIO Deduction names (creatable rows only) ─────────────────────
    _render_name_editor(
        "✏️ Edit UZIO Deduction names (optional)",
        [r for r in enriched_deds
         if r.get("UZIO Master Deductions List") != NEEDS_REVIEW],
        "UZIO Deduction Name", "adp_ppsh_dedname",
        caption=(
            "Set the exact display name each deduction should have in UZIO (the "
            "Master Deduction above keeps driving the form). Flows into the setup "
            "table below and the setup Excel."
        ),
    )

    st.markdown("### UZIO Deduction Setup (all form fields)")
    st.dataframe(
        pd.DataFrame(enriched_deds, columns=DEDUCTION_OUTPUT_COLUMNS),
        hide_index=True, use_container_width=True,
    )
    st.markdown("---")
    return enriched_deds, skipped_deds


def _render_contribution_setup_section(results, enriched_deds):
    """UZIO Contribution Setup — employer match money from the MEMO columns.
    Lets the user link each contribution to a company deduction (defaults for
    401k / Roth 401k / Medical), rename it, and review the form fields.
    Returns the enriched rows for the combined setup workbook."""
    candidates = results.get("Memo_Candidates") or []

    st.markdown("## UZIO Contribution Setup")
    if not candidates:
        st.caption("(no MEMO columns found in the file(s))")
        st.markdown("---")
        return []

    # ── Which MEMO columns are employer contributions? (auto + manual override)
    # Auto-detection (name "MATCH" / Roth: split column / value-based small-%-of-
    # gross-tied-to-a-deferral) pre-selects the picker; the user corrects it when
    # the tool guesses wrong (e.g. a gross-pay memo) or misses one.
    detected_cols = [c["_Source Column"] for c in candidates if c.get("_Detected")]
    all_cols = [c["_Source Column"] for c in candidates]

    def _opt_label(col):
        c = next((x for x in candidates if x["_Source Column"] == col), None)
        if not c:
            return col
        tag = "  ·  auto-detected" if c.get("_Detected") else ""
        return f"{col}   (${c.get('_Total $', 0):,.2f}, {c.get('_Employees', 0)} emp){tag}"

    st.caption(
        "Employer match money lives in `MEMO` columns. The tool auto-detects them "
        "by name (contains **MATCH**), by `Roth:` split columns, or by **value** "
        "(a small % of gross that lines up with a 401k/Roth deferral). If the "
        "auto-pick is wrong, fix it here — every selected column becomes a "
        "contribution; deselect ones that aren't (e.g. a gross-pay or PTO memo)."
    )
    selected_cols = st.multiselect(
        "Memo columns to treat as employer contributions",
        options=all_cols,
        default=detected_cols,
        format_func=_opt_label,
        key="adp_ppsh_contrib_cols",
        help="Pre-selected with what the tool detected. Add a column it missed, "
             "or remove a wrong one.",
    )
    contrib_rows = build_contributions_from_memo_cols(candidates, selected_cols)
    if not contrib_rows:
        if detected_cols:
            st.caption("(no memo columns selected — pick at least one above)")
        else:
            st.caption(
                "(nothing auto-detected as an employer match — if this client has "
                "a 401k/Roth match, pick its MEMO column above)"
            )
        st.markdown("---")
        return []

    # ── Link each contribution to a company deduction ───────────────────────
    creatable_deds = [
        r for r in (enriched_deds or [])
        if r.get("UZIO Master Deductions List")
        and r["UZIO Master Deductions List"] != NEEDS_REVIEW
    ]
    available_masters = sorted({r["UZIO Master Deductions List"] for r in creatable_deds})
    available_display = sorted({
        r.get("UZIO Deduction Name") for r in creatable_deds if r.get("UZIO Deduction Name")
    })
    master_to_display = {}
    for r in creatable_deds:
        master_to_display.setdefault(
            r["UZIO Master Deductions List"], r.get("UZIO Deduction Name", ""))
    link_options = [CONTRIB_LINK_NONE] + available_display

    link_map = {}
    with st.expander("Map contributions to a company deduction", expanded=False):
        st.caption(
            f"**{len(contrib_rows)} contribution(s).** Pick the company deduction "
            f"to link each one to (default-mapped for 401k / Roth 401k / Medical). "
            f"Choose `{CONTRIB_LINK_NONE}` to create the contribution without a "
            f"link. **401k / Roth 401k** use Method **{CONTRIB_METHOD_FORMULA}** "
            f"({_format_formula(CONTRIB_FORMULA_TIERS)}); all others use "
            f"**{CONTRIB_METHOD_FIXED}**."
        )
        for code, td in [(c["Type Code"], c["Type Description"]) for c in contrib_rows]:
            default_master = map_contribution_to_deduction(code, td, available_masters)
            default_opt = master_to_display.get(default_master, CONTRIB_LINK_NONE)
            if default_opt not in link_options:
                default_opt = CONTRIB_LINK_NONE
            name_col, pick_col = st.columns([1, 1])
            with name_col:
                st.markdown(f"**{code}** — {td}" if code else f"**{td}**")
            with pick_col:
                sel = st.selectbox(
                    f"Link {code or td}", options=link_options,
                    index=link_options.index(default_opt),
                    key=f"adp_ppsh_contriblink::{autosync_row_key(code, td)}",
                    label_visibility="collapsed",
                )
            link_map[autosync_row_key(code, td)] = "" if sel == CONTRIB_LINK_NONE else sel

    enriched_contribs = enrich_contributions_for_uzio(contrib_rows, link_map)

    _render_name_editor(
        "✏️ Edit UZIO Contribution names (optional)",
        enriched_contribs, "Contribution Name", "adp_ppsh_conname",
        caption=(
            "Set the exact name each contribution should have in UZIO. Flows "
            "into the setup table below and the setup Excel."
        ),
    )

    st.markdown("### UZIO Contribution Setup (all form fields)")
    st.dataframe(
        pd.DataFrame(enriched_contribs, columns=CONTRIBUTION_OUTPUT_COLUMNS),
        hide_index=True, use_container_width=True,
    )
    st.markdown("---")
    return enriched_contribs


def _render_tax_mapping_section(results, src_name):
    """Taxes — ADP tax columns → UZIO tax catalog. Mapping only (UZIO already
    owns these taxes); produces the Tax Mapping CSV. Federal & most state taxes
    auto-map; LIVED-IN LOCAL and anything ambiguous gets the finder."""
    taxes = results.get("ADP_Taxes") or []

    st.markdown("## Taxes")
    if not taxes:
        st.caption("No tax columns found in the file(s).")
        st.markdown("---")
        return []

    catalog = _cached_tax_catalog()
    if not catalog:
        st.warning("UZIO tax catalog not found (`uzio_tax_catalog.csv`) — can't map taxes.")
        return []

    ee = [t for t in taxes if t["Side"] == "Employee"]
    er = [t for t in taxes if t["Side"] == "Employer"]
    st.caption(
        f"**{len(taxes)} tax row(s)** — {len(ee)} employee + {len(er)} employer "
        f"(state-scoped columns expand to one row per worked-in state). Federal & "
        f"state auto-map from the UZIO catalog; **LIVED-IN LOCAL needs you to pick "
        f"the city/school** in the finder. Taxes are NOT created in UZIO — this "
        f"only produces the mapping file."
    )

    TAX_LEAVE = "— leave unmapped —"
    TAX_ANY_STATE = "🔍 Search all states"
    state_options = sorted({r["state_abbreviation"].upper() for r in catalog})

    def _opt_label(r):
        base = f"{r['tax_name']}  [{r['tax_code']}]"
        return base + (f" — {r['sub_tax_desc']}" if r["sub_tax_desc"] else "")

    resolved = {}
    with st.expander("Map taxes to a UZIO tax", expanded=False):
        st.caption(
            "Every tax gets the same finder: pick the **State**, optionally type "
            "part of the tax / city / school name, then choose the exact UZIO tax. "
            "Pre-selected to the tool's best guess where it's confident."
        )
        for t in taxes:
            key = adp_tax_key(t)
            best = adp_tax_best_match(t, catalog)
            badge = {"federal": "Federal", "state": "State",
                     "local": "Local — find & confirm",
                     "unknown": "Unmapped — find"}.get(t["Tier"], t["Tier"])

            nm_col, pk_col = st.columns([1, 1])
            with nm_col:
                st.markdown(f"**{t['Tax Name']}**")
                state_note = f" · {t['State']}" if t.get("State") else ""
                st.caption(f"{t['Side']} side · {badge}{state_note}")

            state_key = f"adp_ppsh_taxstate::{key}"
            search_key = f"adp_ppsh_taxsearch::{key}"
            tax_key = f"adp_ppsh_tax::{key}"

            if best:
                guess_state = best.get("state_abbreviation", "").upper()
            elif t.get("State") and t["State"] != "FED":
                guess_state = t["State"]
            elif t.get("State") == "FED":
                guess_state = "FED"
            else:
                guess_state = TAX_ANY_STATE
            if guess_state != TAX_ANY_STATE and guess_state not in state_options:
                guess_state = TAX_ANY_STATE
            if state_key not in st.session_state:
                st.session_state[state_key] = guess_state

            with pk_col:
                state_sel = st.selectbox(
                    f"State {key}", options=[TAX_ANY_STATE] + state_options,
                    key=state_key, label_visibility="collapsed",
                )
                all_states = (state_sel == TAX_ANY_STATE)
                search = st.text_input(
                    f"Search {key}", value="", key=search_key,
                    placeholder="type part of the tax / city / school name…",
                    label_visibility="collapsed",
                )
                q = (search or "").strip().lower()
                toks = [w for w in re.split(r"[^a-z0-9]+", q) if len(w) >= 3 and not w.isdigit()]
                pool = (list(catalog) if all_states
                        else [r for r in catalog if r["state_abbreviation"].upper() == state_sel])
                if toks:
                    strict = [r for r in pool if all(w in r["tax_name"].lower() for w in toks)]
                    pool = strict if strict else [
                        r for r in pool if any(w in r["tax_name"].lower() for w in toks)]
                pool = pool[:400]
                if best and best not in pool:
                    pool = [best] + pool
                labels = [TAX_LEAVE] + [_opt_label(r) for r in pool]
                rowmap = {_opt_label(r): r for r in pool}
                bl = _opt_label(best) if best else None
                default_index = labels.index(bl) if (bl and bl in labels) else 0
                sel = st.selectbox(
                    f"Tax {key}", options=labels, index=default_index,
                    key=tax_key, label_visibility="collapsed",
                )
                resolved[key] = rowmap.get(sel)

    mapping_rows = build_adp_tax_mapping_rows(taxes, resolved)
    st.markdown("### Tax Mapping (preview)")
    st.dataframe(
        pd.DataFrame(mapping_rows, columns=MAPPING_TAX_COLUMNS),
        hide_index=True, use_container_width=True,
    )
    unmapped = [r for r in mapping_rows if not r["Uzio Tax Code"]]
    if unmapped:
        st.warning(
            f"**{len(unmapped)} tax(es) unmapped**: "
            f"`{'; '.join(r['Source Tax Code Name'] for r in unmapped)}` "
            "— pick them in the finder above. (The Tax mapping CSV ships in the "
            "**Download all mapping CSVs** bundle below.)"
        )
    st.markdown("---")
    return mapping_rows


def render_ui():
    st.title("ADP - Prior Payroll Setup Helper")
    st.caption(
        "Upload the ADP Prior Payroll file(s): what to set up in Uzio, "
        "is the bonus discretionary, and which deductions are pre-tax vs post-tax."
    )

    adp_files = st.file_uploader(
        "ADP Prior Payroll file(s) (sanitized)",
        type=["xlsx", "xls", "csv"],
        key="pps_helper_adp",
        accept_multiple_files=True,
    )
    client_name = st.text_input(
        "Client Name",
        value="",
        placeholder="e.g. Happy Delivery",
        key="adp_ppsh_client",
        help=("Used in the downloaded file names: <Client Name>_UZIO_Setup.xlsx, "
              "<Client Name>_Tax_Mapping.csv, and the mapping-files zip."),
    )

    if not adp_files:
        st.info("Upload at least one ADP Prior Payroll file to begin.")
        return

    if len(adp_files) > 1:
        st.caption(f"Combining {len(adp_files)} Prior Payroll file(s) for analysis.")

    # Persist results in session_state so the interactive Earning Setup controls
    # (toggles / dropdowns / name editors) — which each trigger a Streamlit rerun
    # — don't lose the analysis when the Run button is no longer "pressed".
    if st.button("Run", type="primary"):
        with st.spinner("Analyzing..."):
            try:
                results, csv_bytes = run_setup_helper(adp_files)
                st.session_state["adp_ppsh_results"] = results
                st.session_state["adp_ppsh_csv"] = csv_bytes
                st.session_state["adp_ppsh_filename"] = adp_files[0].name
            except Exception as e:
                st.session_state.pop("adp_ppsh_results", None)
                st.error(f"Failed to run analysis: {e}")
                raise

    results = st.session_state.get("adp_ppsh_results")
    if not results:
        return
    src_name = st.session_state.get("adp_ppsh_filename") or "ADP_Prior_Payroll"
    # Client Name (when given) drives every download file name; else the first
    # uploaded file's stem.
    display_base = (client_name or "").strip() or src_name.rsplit(".", 1)[0]

    # ------------------------------------------------------------------
    # UZIO Setup sections (mirror the Paycom helper)
    # ------------------------------------------------------------------
    enriched_earnings, skipped_earnings = _render_earning_setup_section(results, display_base)
    enriched_deds, skipped_deds = _render_deduction_setup_section(results, display_base)
    enriched_contribs = _render_contribution_setup_section(results, enriched_deds)
    tax_mapping_rows = _render_tax_mapping_section(results, display_base)

    # ------------------------------------------------------------------
    # Downloads: the UZIO Setup workbook + the four API mapping CSVs
    # ------------------------------------------------------------------
    st.markdown("## Downloads")
    if enriched_earnings or enriched_deds or enriched_contribs:
        st.download_button(
            "Download UZIO Setup (xlsx) — Earnings + Deductions + Contributions",
            data=build_setup_xlsx(enriched_earnings, enriched_deds, enriched_contribs),
            file_name=f"{display_base}_UZIO_Setup.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="adp_ppsh_setup_dl",
            type="primary",
        )

    st.markdown("### API Mapping Files")
    st.caption(
        "The four Source → UZIO mapping CSVs the API run consumes, downloaded as "
        "separate files. Source names are the original column headers VERBATIM; "
        "UZIO names are exactly what the setup sections above create (renames "
        "flow in automatically). Skipped UZIO defaults are included so payroll "
        "data still uploads against the existing system-created items."
    )
    safe_name = re.sub(r"[^A-Za-z0-9_-]+", "_", display_base).strip("_") or "Client"
    mapping_files = [
        (f"{safe_name}_Earnings_mapping.csv",
         _mapping_csv_bytes(build_earnings_mapping_rows(enriched_earnings, skipped_earnings),
                            MAPPING_EARNING_COLUMNS)),
        (f"{safe_name}_Deductions_mapping.csv",
         _mapping_csv_bytes(build_deductions_mapping_rows(enriched_deds, skipped_deds),
                            MAPPING_DEDUCTION_COLUMNS)),
        (f"{safe_name}_Contributions_mapping.csv",
         _mapping_csv_bytes(build_contributions_mapping_rows(enriched_contribs),
                            MAPPING_CONTRIBUTION_COLUMNS)),
        (f"{safe_name}_Tax_mapping.csv",
         _mapping_csv_bytes(tax_mapping_rows, MAPPING_TAX_COLUMNS)),
    ]
    files_payload = json.dumps(
        [[fn, base64.b64encode(data).decode("ascii")] for fn, data in mapping_files])
    import streamlit.components.v1 as components
    components.html(MAPPING_DOWNLOAD_HTML.replace("__FILES__", files_payload), height=110)
    st.markdown("---")

    # ------------------------------------------------------------------
    # ANSWER 1 — Bonus discretionary or non-discretionary
    # ------------------------------------------------------------------
    st.markdown("## 1. Bonus: discretionary or non-discretionary?")
    bonus = results["Bonus_Classification"][0]
    verdict = bonus["Verdict"]
    sample = _pick_bonus_example(results["Bonus_Sample_Rows"], verdict)

    if verdict == "non_discretionary":
        st.error("**NON-DISCRETIONARY** — overall verdict (combined test across all "
                 "bonus columns). Each earning's own verdict is in the table below.")
    elif verdict == "discretionary":
        st.success("**DISCRETIONARY** — overall verdict (combined test across all "
                   "bonus columns). Each earning's own verdict is in the table below.")
    else:
        st.warning(f"**{verdict.upper()}** — {bonus['Reason']}")

    if sample:
        breakdown = sample.get("bonus_breakdown") or f"{sample['bonus_amt']:,}"
        st.markdown(
            f"""
**Example: Employee `{sample['associate']}`** — bonus(es) in this paycheck: **{breakdown}**

- Regular rate = **{sample['regular_rate']}/hr** · Worked hours (Reg + OT + Dbl OT + Training) = **{sample['worked_hours']}**
- Blended rate with bonus = **{sample['blended_rate']}/hr** (total bonus {sample['bonus_amt']:,} spread over worked hours)
- Predicted OT if **discretionary** (plain 1.5×): **{sample['predicted_ot_discretionary']:,}**
- Predicted OT if **non-discretionary** (blended): **{sample['predicted_ot_non_disc']:,}**
- **Actual OT paid: {sample['actual_ot']:,}**

*(all amounts in $)*
"""
        )
        if verdict == "discretionary":
            st.markdown(
                "→ Actual OT matches the plain **1.5×** prediction — the bonus did **not** "
                "inflate the regular rate basis, so it's **discretionary**."
            )
        elif verdict == "non_discretionary":
            st.markdown(
                "→ Actual OT matches the **blended-rate** prediction — the bonus was rolled "
                "into the regular rate before computing OT, so it's **non-discretionary** under FLSA."
            )

    per_bonus = results.get("Bonus_Per_Column") or []
    if per_bonus:
        st.markdown(
            "**Per-bonus verdicts** — *individual* = paychecks where only that bonus "
            "appears; *combined* = tested with the sum of all bonuses on the paycheck:"
        )
        st.dataframe(pd.DataFrame([{
            "Bonus Column": p["column"],
            "Final Verdict": str(p["final_verdict"]).replace("_", " ").upper(),
            "Individual": f"{p['individual_verdict']} ({p['individual_rows']} rows)",
            "Combined": f"{p['combined_verdict']} ({p['combined_rows']} rows)",
            "Note": p["note"],
        } for p in per_bonus]), hide_index=True, use_container_width=True)

    # ------------------------------------------------------------------
    # ANSWER 2 — Which deductions are pre-tax vs post-tax
    # ------------------------------------------------------------------
    st.markdown("## 2. Pre-tax vs post-tax (per deduction)")
    if not results["Deductions"] and not results["Contributions"]:
        st.caption("No voluntary deductions or contributions found in this file.")
    else:
        rows = []
        for r in results["Contributions"] + results["Deductions"]:
            rows.append({
                "Code": r["Code"],
                "Verdict": "PRE-TAX" if r["Verdict"] == "pre_tax" else "POST-TAX",
                "Why": _deduction_reason(r),
            })
        st.dataframe(pd.DataFrame(rows), hide_index=True, use_container_width=True)

    # ------------------------------------------------------------------
    # Downloads (full report stays available, just out of the way)
    # ------------------------------------------------------------------
    st.markdown("---")
    base = (src_name or "ADP_Prior_Payroll").rsplit(".", 1)[0]
    xlsx_bytes = _results_to_xlsx_bytes(results)
    st.download_button(
        "Download Full Detailed Report (xlsx)",
        data=xlsx_bytes,
        file_name=f"{base}_Setup_Helper.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Full multi-sheet workbook with $ totals, hours, sample rows, etc.",
    )
