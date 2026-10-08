"""Rebuild Uzio's "Employee Profile Change Report" from the history tables.

This is a port of the report Uzio itself generates, not a lookalike. Every label,
every ordering rule and every value format below was read off the generator:

  phix-monolith .../report/impl/EmployeeAuditPoiReport.java      (sheet layout)
  phix-monolith .../report/impl/EmployeeAuditMapping.java        (labels + order)
  phix-monolith .../impl/EmployeeHistoryCustomConvertorImpl.java (value formats)

It reads `employee_history` / `emergency_contact_history` through the NeuronOps
query endpoint, so it can only show what that endpoint returns. Two consequences,
both surfaced to the user rather than papered over:

* SSN, Hourly Pay Rate, Annual Salary, Bonus and Salary Commissions are encrypted
  at rest. Uzio's own report decrypts them inside the application; the query
  endpoint hands back the ciphertext, so those cells read "(encrypted)". The
  effective-date suffix next to them is real and is still shown.
* Work Schedule needs a work-week-schedule lookup, and the Family (dependents)
  section needs the family-member history table. Neither is exposed by the query
  endpoint, so Work Schedule stays blank and the Family section is not emitted.

"Union Classification" is dropped the way Uzio drops it for an Amazon exchange
(`isAmazonExchange` in the generator), which is every DSP client this tool is
used for. It is emitted anyway if any version actually carries a value, so the
row can never hide data.
"""
import json
import re
from datetime import datetime, timezone

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SHEET = "Change History"
ENCRYPTED = "(encrypted)"
VERSION_STR = "Version: V"
EFFECTIVE_PREFIX = " (Effective Date - "

# All four enums below are ordinal-indexed exactly as the Java enums declare them.
GENDER = {0: "Male", 1: "Female", 2: "Intersex"}
STATUS = {0: "UNDER_PROVISION", 1: "ACTIVE", 2: "TERMINATED", 3: "DECEASED",
          4: "DELETED", 5: "FUTURE_HIRE"}
EMPLOYMENT_TYPE = {0: "Part Time", 1: "Full Time", 2: "Both Part Time and Full Time",
                   3: "Others", 4: "All", 5: "Term", 6: "Seasonal"}
TERMINATION = {"DEATH": "Death", "RETIREMENT": "Retirement",
               "DISABILITY": "Permanent Disability",
               "QUIT": "Voluntary Termination of Employment",
               "FIRED": "Involuntary Termination of Employment",
               "TRANSFER": "Transfer", "OTHER": "Other"}
# populateWhoChanged's fallback when a login carries no role: Uzio prints the user
# type, not the role name.
USER_TYPE = {"EMPLOYEE": "EE", "EMPLOYER": "ER", "BROKER": "BR", "SUBBROKER": "BR",
             "ADMIN": "CSR", "BROKERAGENCY": "BR"}
SOURCE = {"USER_INTERFACE": "User Interface", "CENSUS": "Census Template",
          "ENROLLMENT": "Enrollment Update", "PAYROLL_INTEGRATION": "Payroll Integration"}

# (section, label, kind, column, effective-date key).
# The effective-date key is the `job_effective_dates` entry whose date gets appended
# as " (Effective Date - ...)". The set comes from the generator's own map, which is
# why Work Location has none: there the suffix is written onto `workLocationIdentifier`
# while the report prints `workLocation`.
FIELDS = [
    ("Personal", "Employee Prefix", "text", "salutation", None),
    ("Personal", "Employee First Name", "text", "first_name", None),
    ("Personal", "Employee Middle Initial", "text", "middle_name", None),
    ("Personal", "Employee Last Name", "text", "last_name", None),
    ("Personal", "Employee Suffix", "text", "name_suffix", None),
    ("Personal", "Employee Date of Birth", "date", "date_of_birth", None),
    ("Personal", "Employee Gender", "gender", "gender", None),
    ("Personal", "Employee Marital Status", "marital", "marital_status", None),
    ("Personal", "Employee SSN", "encrypted", "ssn", None),
    ("Personal", "Employee Tobacco usage in last 12 months", "bool", "is_smoker", None),

    ("Job", "Employee ID", "text", "ext_employee_code", None),
    ("Job", "Employment Status", "status", "status", None),
    ("Job", "Date of Hire", "date", "date_of_hire", None),
    ("Job", "Adjusted Service Date", "date", "adjusted_service_date", None),
    ("Job", "Termination Date", "date", "date_of_termination", None),
    ("Job", "Termination Reason", "termination", "termination_type", None),
    ("Job", "Employment Type", "employment_type", "employment_type", "EMPLOYEMENT_TYPE_I"),
    ("Job", "Others (Employment Type)", "text", "other_employment_type", None),
    ("Job", "Pay Type", "text", "pay_type", "PAY_TYPE"),
    ("Job", "Pay Group", "text", "pay_group_name", "PAY_GROUP"),
    ("Job", "Annual Salary", "encrypted", "annual_salary", "ANNUAL_SALARY"),
    ("Job", "Salary Effective Date", "date", "salary_effective_date", None),
    ("Job", "Working Hours per Week", "text", "hours", "WORKING_HOURS_PER_WEEK"),
    ("Job", "Hourly Pay Rate", "encrypted", "hourly_rate", "HOURLY_RATE"),
    ("Job", "Benefits Class", "text", "class_name", "CLASS"),
    ("Job", "Job Title", "text", "designation", "JOB_TITLE"),
    ("Job", "Department", "text", "department", "DEPARTMENT"),
    ("Job", "Employee Type", "text", "employment_type_two", None),
    ("Job", "Union Classification", "text", "employment_classification_one", None),
    ("Job", "FLSA Classification", "text", "employment_classification_two", None),
    ("Job", "Work Location", "text", "work_location", None),
    ("Job", "Division", "text", "division", "DIVISION"),
    ("Job", "Bonus", "encrypted", "bonus", "BONUS"),
    ("Job", "Salary Commissions", "encrypted", "commission", "COMMISSION"),
    ("Job", "Reporting Manager", "text", "reporting_to_name", None),
    ("Job", "Special Hire", "bool", "special_hire_flag", None),
    ("Job", "Protected Veteran Status", "text", "protected_veteran_status", None),
    ("Job", "Disability Status", "text", "disability_status", None),
    ("Job", "EEO Job Category", "text", "eeo_job_category", None),
    ("Job", "Race/Ethnicity", "text", "race_ethnicity", None),
    ("Job", "Original DOH", "date", "original_doh", None),
    ("Job", "Work Schedule", "text", "work_week_schedule_name", None),

    ("Contact", "Personal Email", "text", "email", None),
    ("Contact", "Official Email", "text", "emailofficial", None),
    ("Contact", "Phone Number(Digits)", "phone", "phone", None),
    ("Contact", "Employee Address Line 1", "text", "address_line1", "ADDRESS"),
    ("Contact", "Employee Address Line 2", "text", "address_line2", "ADDRESS"),
    ("Contact", "City", "text", "city", "ADDRESS"),
    ("Contact", "Zip Code", "text", "zip", "ADDRESS"),
    ("Contact", "State(Abbreviation)", "text", "state", "ADDRESS"),
    ("Contact", "County", "text", "county", "ADDRESS"),
]

# Uzio drops this row for an Amazon exchange, which is every DSP client.
AMAZON_SKIPPED = "Union Classification"

EMERGENCY_SECTION = "Emergency Contact Details"
EMERGENCY_FIELDS = [
    ("Name", "text", "name"),
    ("Relationship", "text", "relationship"),
    ("Address (Line 1)", "text", "address_line1"),
    ("Address (Line 2)", "text", "address_line2"),
    ("City", "text", "city"),
    ("Zip Code", "text", "zip"),
    ("County", "text", "county"),
    ("State(Abbreviation)", "text", "state"),
    ("Email ID", "text", "email"),
    ("Phone Number(Digits)", "phone", "phone"),
    ("CA Arrest/Detention Notification Authorization", "lower_bool", "detention_arrest_consent"),
]

CUSTOM_SECTION = "Custom Fields"


# --------------------------------------------------------------------------- values

def _parse(value):
    if value in (None, "", "None"):
        return None
    try:
        return datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None


def _date(value) -> str:
    d = _parse(value)
    return d.strftime("%m/%d/%Y") if d else ""


def modified_on(value) -> str:
    """Modified On, as Uzio prints it: 09/09/2026 07:33:57 PM UTC."""
    d = _parse(value)
    if not d:
        return ""
    d = d.astimezone(timezone.utc) if d.tzinfo else d
    return d.strftime("%m/%d/%Y %I:%M:%S %p UTC")


def _json(value):
    if not value:
        return {}
    try:
        return json.loads(value) if isinstance(value, str) else value
    except ValueError:
        return {}


def _phone(value) -> str:
    """The generator's own rule: leave anything containing "-" alone, else split 3/3/rest."""
    text = str(value or "").strip()
    if not text or "-" in text or len(text) < 7:
        return text
    return f"{text[:3]}-{text[3:6]}-{text[6:]}"


def _normalize(value) -> str:
    return str(value).replace("&#96;", "`")


def person_label(email: str) -> str:
    """"tierra.williams1@uzio.com" -> "Tierra Williams (CSR)".

    Uzio builds this from the user profile (first/middle/last plus the role, "CSR"
    for staff and the account's own role names for a client login). Those profile
    tables are not exposed by the query endpoint, so a uzio.com address — always
    firstname.lastname — is expanded to the same "Name (CSR)" the real report shows,
    and any other login keeps its address verbatim rather than being given an
    invented display name or a guessed role.
    """
    email = (email or "").strip()
    if not email or email in ("System", "SCRIPT"):
        return email
    local, _, domain = email.partition("@")
    if not domain.lower().endswith("uzio.com"):
        return email
    parts = [p for p in local.replace("_", ".").split(".") if p]
    if len(parts) >= 2 and all(p.strip("0123456789").isalpha() for p in parts):
        return " ".join(p.strip("0123456789").capitalize() for p in parts) + " (CSR)"
    return f"{email} (CSR)"


def who_changed(row: dict) -> str:
    """The "Who Changed" line, from the user tables when they could be read."""
    return row.get("who_changed") or person_label(row.get("created_by"))


def effective_date(row: dict, key: str) -> str:
    """The date Uzio appends for `key`.

    Missing keys fall back the way the generator falls back: to the salary
    effective date, except for ADDRESS, which falls back to a literal "--".
    """
    found = (_json(row.get("job_effective_dates")) or {}).get(key)
    if found:
        return str(found)
    return "--" if key == "ADDRESS" else _date(row.get("salary_effective_date"))


def cell_value(row: dict, kind: str, column, effective_key=None) -> str:
    """One field of one version, formatted the way the Uzio report formats it."""
    if kind == "custom":
        for item in _json(row.get("custom_field")) or []:
            if isinstance(item, dict) and item.get("key") == column:
                raw = item.get("value")
                return "" if raw in (None, "") else _normalize(raw)
        return ""

    raw = row.get(column)
    blank = raw in (None, "", "None")

    if kind == "encrypted":
        value = ENCRYPTED if not blank else ""
    elif kind == "bool":
        # isSmoker / isSpecialHire are normalised to Yes/No even when unset.
        value = "Yes" if str(raw) in ("1", "True", "true") else "No"
    elif blank:
        value = ""
    elif kind == "date":
        value = _date(raw)
    elif kind == "gender":
        value = GENDER.get(int(raw), "") if str(raw).lstrip("-").isdigit() else _normalize(raw)
    elif kind == "status":
        name = STATUS.get(int(raw), str(raw)) if str(raw).lstrip("-").isdigit() else str(raw)
        value = "ACTIVE" if name == "FUTURE_HIRE" else name
    elif kind == "employment_type":
        value = (EMPLOYMENT_TYPE.get(int(raw), "")
                 if str(raw).lstrip("-").isdigit() else _normalize(raw))
    elif kind == "marital":
        value = str(raw).replace("_", " ").title()
    elif kind == "termination":
        value = TERMINATION.get(str(raw).upper(), _normalize(raw))
    elif kind == "phone":
        value = _phone(raw)
    elif kind == "lower_bool":
        value = "true" if str(raw) in ("1", "True", "true") else "false"
    else:
        value = _normalize(raw)

    if value and effective_key:
        value = f"{value}{EFFECTIVE_PREFIX}{effective_date(row, effective_key)})"
    return value


# --------------------------------------------------------------------------- styles

def _styles(wb):
    wrap = dict(wrap_text=True)
    return {
        "blue": (PatternFill("solid", start_color="FF95B3D7"), Font(bold=True),
                 Alignment(**wrap)),
        "red": (PatternFill("solid", start_color="FFC00000"), None, Alignment(**wrap)),
        "big": (PatternFill("solid", start_color="FFC6D9F0"), Font(size=14),
                Alignment(**wrap)),
        "center": (None, None, Alignment(horizontal="center", **wrap)),
        "green": (PatternFill("solid", start_color="FFD9F0F3"), Font(bold=True),
                  Alignment(**wrap)),
        "green_center": (PatternFill("solid", start_color="FFD9F0F3"), Font(bold=True),
                         Alignment(horizontal="center", **wrap)),
    }


def _paint(cell, style):
    fill, font, alignment = style
    if fill is not None:
        cell.fill = fill
    if font is not None:
        cell.font = font
    if alignment is not None:
        cell.alignment = alignment


# --------------------------------------------------------------------------- sheet

class _Sheet:
    """Writes the grid the Java report writes: label columns A/B, one value column
    per version newest-first, and a narrow red spacer column after each."""

    def __init__(self, ws, count, styles):
        self.ws = ws
        self.count = count
        self.st = styles
        self.row = 1

    def column(self, index):
        return 3 + index * 2

    def spacers(self, row):
        for i in range(self.count):
            _paint(self.ws.cell(row=row, column=self.column(i) + 1), self.st["red"])

    def headers(self, versions):
        """versions newest-first; each contributes five header lines."""
        for i, version in enumerate(versions):
            col = self.column(i)
            lines = [
                "Version: " + VERSION_STR + str(version["_number"]),
                "Who Changed: " + who_changed(version),
                "Modified On: " + modified_on(version.get("created_date")),
                "IP Address: " + (version.get("ip_address") or ""),
                "Source of Change: " + SOURCE.get(str(version.get("source") or ""),
                                                  str(version.get("source") or "")),
            ]
            for offset, text in enumerate(lines):
                _paint(self.ws.cell(row=1 + offset, column=col, value=text), self.st["blue"])
            _paint(self.ws.cell(row=6, column=col), self.st["big"])
        for row in range(1, 7):
            self.spacers(row)
        _paint(self.ws.cell(row=6, column=1, value="Section"), self.st["blue"])
        _paint(self.ws.cell(row=6, column=2, value="EE Profile Details"), self.st["blue"])
        self.row = 7

    def line(self, section, label, values):
        """One label row. Highlighted when any version differs from the newest."""
        row = self.row
        self.row += 1
        head = [self.ws.cell(row=row, column=1, value=section),
                self.ws.cell(row=row, column=2, value=label)]
        changed = len(set(values)) > 1
        for i, value in enumerate(values):
            cell = self.ws.cell(row=row, column=self.column(i), value=value or None)
            _paint(cell, self.st["green_center"] if changed else self.st["center"])
        if changed:
            for cell in head:
                _paint(cell, self.st["green"])
        self.spacers(row)


def build_workbook(versions: list, contacts: dict = None) -> openpyxl.Workbook:
    """`versions` oldest-first (V1 first), as the history tables return them.

    `contacts` maps an employee_history row id to that version's emergency contacts.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET
    styles = _styles(wb)

    for n, row in enumerate(versions):
        row["_number"] = n + 1
    newest_first = list(reversed(versions))

    sheet = _Sheet(ws, len(newest_first), styles)
    sheet.headers(newest_first)

    for section, label, kind, column, key in FIELDS:
        values = [cell_value(v, kind, column, key) for v in newest_first]
        if label == AMAZON_SKIPPED and not any(values):
            continue
        sheet.line(section, label, values)

    keys = []
    for version in newest_first:
        for item in _json(version.get("custom_field")) or []:
            if isinstance(item, dict) and item.get("key") not in keys:
                keys.append(item["key"])
    for key in keys:
        sheet.line(CUSTOM_SECTION, _normalize(key),
                   [cell_value(v, "custom", key) for v in newest_first])

    _emergency(sheet, newest_first, contacts or {})

    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 33
    for i in range(len(newest_first)):
        ws.column_dimensions[get_column_letter(sheet.column(i))].width = 50
        ws.column_dimensions[get_column_letter(sheet.column(i) + 1)].width = 1.5
    ws.freeze_panes = "A2"
    return wb


def _emergency(sheet, newest_first, contacts):
    """One 11-row block per emergency-contact record.

    A contact gets a fresh `emergency_contact_identifier` on every edit, so the same
    person appears once per version. Uzio pads each version with the identifiers it
    is missing and sorts by identifier, which is the order reproduced here.
    """
    per_version = []
    identifiers = set()
    for version in newest_first:
        by_id = {c.get("emergency_contact_identifier"): c
                 for c in contacts.get(version.get("id"), [])}
        per_version.append(by_id)
        identifiers.update(by_id)
    for identifier in sorted(i for i in identifiers if i):
        for label, kind, column in EMERGENCY_FIELDS:
            values = []
            for by_id in per_version:
                contact = by_id.get(identifier)
                values.append(cell_value(contact, kind, column) if contact else "")
            sheet.line(EMERGENCY_SECTION, label, values)


def report_filename(when: datetime = None, employee: str = "") -> str:
    """Uzio's own export name, with the employee added when several are produced."""
    when = when or datetime.now()
    tag = re.sub(r"[^A-Za-z0-9._-]+", "_", employee).strip("_")
    stamp = when.strftime("%Y-%m-%d-%H-%M-%S")
    return (f"Employee Profile Change Report_{tag}_{stamp}.xlsx" if tag
            else f"Employee Profile Change Report_{stamp}.xlsx")


# ----------------------------------------------------------------------- data access

HISTORY_SOURCES = "employee_history / emergency_contact_history"


def _by_historical_id(token, ids, rows, companies, kwargs) -> list:
    """Employees whose Employee ID USED to be one of `ids` but has since changed.

    A census carries the ID of its day; Uzio replaces it later (1020 became
    BH0KS5HPZ). Searching only the current code finds nothing and looks like the
    employee is absent, so any ID that matched nothing is looked up again in the
    history, within the same company.
    """
    from utils import neuronops_client as ops

    seen = {str(r.get("ext_employee_code")) for r in rows}
    missing = ops.sql_in_list([i for i in ids if i not in seen])
    eins = ops.sql_in_list({c.get("ein") for c in companies if c.get("ein")})
    if not missing or not eins:
        return []
    history = ops.query(token, "select distinct employee_code, ext_employee_code from "
                        f"employee_history where ext_employee_code in ({missing}) "
                        f"and ein in ({eins}) and deleted = 0", **kwargs)
    known = {r.get("employee_code") for r in rows}
    was = {h.get("employee_code"): h.get("ext_employee_code")
           for h in history if h.get("employee_code") not in known}
    if not was:
        return []
    codes = ops.sql_in_list(was)
    found = ops.query(token, "select employee_code, ext_employee_code, full_name, status, "
                      "employer_organization_id, date_of_hire, date_of_termination "
                      f"from employee where employee_code in ({codes}) and deleted = 0", **kwargs)
    for row in found:
        row["found_via"] = f"was {was.get(row.get('employee_code'))}"
    return found


def find_companies(token: str, feins, host: str = None) -> list:
    """The companies behind one or more 9-digit FEINs."""
    from utils import neuronops_client as ops

    quoted = ops.sql_in_list(feins)
    if not quoted:
        return []
    return ops.query(token, "select id, ein, fein, company_name, client_code from "
                     f"employer_organization where fein in ({quoted}) and deleted = 0",
                     **({"host": host} if host else {}))


def find_employees(token: str, feins=None, ids=None, host: str = None) -> list:
    """Employees of the given company (FEIN), optionally narrowed by Employee ID.

    The FEIN is what keeps this honest: the same Employee ID exists in more than one
    company, so without it a search for "1020" returns strangers. With a FEIN and no
    other filter, every employee of that company comes back for the caller to pick
    from. Without a FEIN the search still runs, but each row carries its company name
    so a duplicate is visible rather than silently chosen.
    """
    from utils import neuronops_client as ops

    kwargs = {"host": host} if host else {}
    scope = ""
    if feins:
        companies = find_companies(token, feins, host=host)
        if not companies:
            return []
        org_ids = ",".join(str(c["id"]) for c in companies if str(c.get("id", "")).isdigit())
        scope = f"employer_organization_id in ({org_ids})"

    clauses = []
    if ids:
        quoted = ops.sql_in_list(ids)
        if quoted:
            clauses.append(f"ext_employee_code in ({quoted})")

    if not scope and not clauses:
        return []
    where = " and ".join(part for part in (scope, "(" + " or ".join(clauses) + ")" if clauses else "") if part)
    rows = ops.query(token, "select employee_code, ext_employee_code, full_name, status, "
                     "employer_organization_id, date_of_hire, date_of_termination "
                     f"from employee where {where} and deleted = 0 order by full_name", **kwargs)

    if ids and feins:
        rows += _by_historical_id(token, ids, rows, companies, kwargs)

    orgs = {str(r.get("employer_organization_id")) for r in rows if r.get("employer_organization_id")}
    companies = {}
    if orgs:
        listed = ",".join(o for o in orgs if o.isdigit())
        for row in ops.query(token, "select id, fein, company_name from employer_organization "
                             f"where id in ({listed})", **kwargs):
            companies[row.get("id")] = row
    for row in rows:
        company = companies.get(row.get("employer_organization_id")) or {}
        row["company_name"] = company.get("company_name") or ""
        row["fein"] = company.get("fein") or ""
        row["status_label"] = cell_value(row, "status", "status")
    return rows


def resolve_users(token: str, logins, host: str = None) -> dict:
    """{login -> "Tobias Conner (Employer Administrator)"} for the people who made the changes.

    `employee_history.created_by` holds whatever the account signs in with: a CSR's
    email, or a bare user identifier for a client login â€” which is why that column
    alone shows a UUID. Uzio resolves it through the user tables, and so does this:
    the profile's name, plus the role names for an employer login or the user type
    for anyone else, exactly as populateWhoChanged builds it.

    Where that user has more than one profile Uzio takes an unordered set's first
    element; this takes the oldest profile, which is the one its export showed.
    """
    from utils import neuronops_client as ops

    wanted = {str(l).strip() for l in logins if l and str(l).strip() not in ("System", "SCRIPT")}
    quoted = ops.sql_in_list(wanted)
    if not quoted:
        return {}
    rows = ops.query(token, "select u.username, u.user_identifier, p.id as profile_id, "
                     "p.user_type, p.first_name, p.middle_name, p.last_name, "
                     "r.name as role_name from user_data u join user_profile p "
                     "on p.user_id = u.id and p.deleted = 0 "
                     "left join USER_ROLE_MAPPING m on m.user_profile_id = p.id and m.deleted = 0 "
                     "left join USER_ROLES r on r.id = m.role_id "
                     f"where u.username in ({quoted}) or u.user_identifier in ({quoted})",
                     **({"host": host} if host else {}))

    profiles = {}
    for row in rows:
        for key in (row.get("username"), row.get("user_identifier")):
            if not key or key not in wanted:
                continue
            profile = profiles.setdefault((key, row.get("profile_id")),
                                          {"row": row, "roles": []})
            if row.get("role_name") and row["role_name"] not in profile["roles"]:
                profile["roles"].append(row["role_name"])

    labels = {}
    for (key, profile_id), profile in sorted(profiles.items(), key=lambda kv: kv[0][1] or 0):
        if key in labels:                      # the oldest profile wins
            continue
        row, roles = profile["row"], profile["roles"]
        name = " ".join(part for part in (row.get("first_name"), row.get("middle_name"),
                                          row.get("last_name")) if part)
        user_type = str(row.get("user_type") or "")
        if user_type == "EMPLOYER" and roles:
            suffix = ",".join(roles)
        else:
            suffix = USER_TYPE.get(user_type, user_type)
        labels[key] = f"{name} ({suffix})" if name and suffix else name or key
    return labels


def fetch_versions(token: str, employee_code: str, host: str = None) -> list:
    """Every history row for one employee, oldest first, with the lookups resolved."""
    from utils import neuronops_client as ops

    kwargs = {"host": host} if host else {}
    safe = re.sub(r"[^A-Za-z0-9-]", "", str(employee_code))
    rows = ops.query(token, "select * from employee_history where employee_code = "
                     f"'{safe}' and deleted = 0 order by created_date, id", **kwargs)
    if not rows:
        return []

    groups = ops.sql_in_list({r.get("pay_group_identifier") for r in rows if r.get("pay_group_identifier")})
    names = {}
    if groups:
        for row in ops.query(token, "select pay_group_identifier, pay_group_name from "
                             f"EMPLOYER_PAY_GROUP where pay_group_identifier in ({groups})", **kwargs):
            names[row.get("pay_group_identifier")] = row.get("pay_group_name")

    managers = ops.sql_in_list({r.get("reporting_to") for r in rows if r.get("reporting_to")})
    full_names = {}
    if managers:
        for row in ops.query(token, "select employee_code, full_name from employee "
                             f"where employee_code in ({managers})", **kwargs):
            full_names[row.get("employee_code")] = row.get("full_name")

    try:
        people = resolve_users(token, {r.get("created_by") for r in rows}, host=host)
    except ops.NeuronOpsError:
        people = {}                            # fall back to the login itself

    for row in rows:
        row["who_changed"] = people.get(str(row.get("created_by") or "").strip(), "")
        row["pay_group_name"] = names.get(row.get("pay_group_identifier")) or ""
        row["reporting_to_name"] = full_names.get(row.get("reporting_to")) or ""
        # The work-week-schedule lookup is not exposed by the query endpoint, so this
        # stays blank rather than printing the raw identifier.
        row["work_week_schedule_name"] = ""
    return rows


def fetch_contacts(token: str, versions: list, host: str = None) -> dict:
    """Emergency contacts for each history row, keyed by that row's id."""
    from utils import neuronops_client as ops

    ids = [str(v.get("id")) for v in versions if v.get("id") is not None]
    if not ids:
        return {}
    by_version = {}
    sql = ("select * from emergency_contact_history where employee_history_id in ("
           + ",".join(i for i in ids if i.isdigit()) + ") and deleted = 0 order by id")
    for row in ops.query(token, sql, **({"host": host} if host else {})):
        by_version.setdefault(row.get("employee_history_id"), []).append(row)
    return by_version


def build_for_employee(token: str, employee_code: str, host: str = None):
    """(workbook, versions) for one employee, or (None, []) when it has no history."""
    versions = fetch_versions(token, employee_code, host=host)
    if not versions:
        return None, []
    contacts = fetch_contacts(token, versions, host=host)
    return build_workbook(versions, contacts), versions
