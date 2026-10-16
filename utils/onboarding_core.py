"""Uzio onboarding-service (production) API -- transport and parsing only.

No Streamlit and no UI here, so both the Streamlit app
(utils/onboarding_api_client.py) and the census-migration skill's scripts run the
exact same code. Keep it that way: anything that draws on screen belongs in the
UI layer, anything that talks to Uzio or shapes its answer belongs here.

What this does:
  * login()                -> JWT, from username + password + FEIN
  * push_employee_census() -> multipart upload of the corrected census
  * push_payment_method()  -> multipart upload of the corrected direct deposit file
  * push_employee_deductions() -> multipart upload of the corrected deduction
    file plus its mapping
  * fetch_run_log() / find_run_since() -> read that run's row back through the
    read-only /app/onboarding/query endpoint (PHIX-98714)

The push response only carries COUNTS (TotalMap / SuccessMap / FailureMap); the
per-employee errors and warnings are written to that run's row in
onboarding_automation_history. When the push gets no answer at all (a big census
outlasts the request timeout while Uzio keeps processing), the run is found by
FEIN + user + start time instead.

Credentials are never persisted here: login() takes them as plain arguments and
the JWT lives only as long as the caller keeps it.
"""
import json
import re
import time
from datetime import datetime, timedelta, timezone

import requests

DEFAULT_PROD_HOST = "https://app.uzio.com"
IST = timezone(timedelta(hours=5, minutes=30))

RUN_LOG_COLS = "id, start_time, end_time, error_messages, optional_validations"
# Only the id is interpolated, and only after int() -- this is never free-form SQL.
RUN_LOG_SQL = "select " + RUN_LOG_COLS + " from onboarding_automation_history where id = {run_id:d}"

# A gateway that gives up on a long request answers with one of these while the
# onboarding service carries on processing behind it.
GATEWAY_TIMEOUTS = {502, 503, 504}
# Allowance for clock drift between this machine and the onboarding DB when
# looking a run up by its start time.
LOOKUP_SKEW = timedelta(seconds=60)
# How long after the push a run may start and still be taken for it. Processing
# begins seconds after the upload finishes, so this is generous -- its job is to
# stop an unrelated later run for the same employer from being reported as this
# push's outcome when the push never arrived at all.
LOOKUP_WINDOW = timedelta(minutes=30)


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


def push_payment_method(jwt_token: str, vendor: str,
                       payment_bytes: bytes, payment_filename: str,
                       prod_host: str = DEFAULT_PROD_HOST, timeout: int = 120):
    """POST {prod_host}/app/onboarding/api/employee/payment-method (multipart).

    The endpoint declares both file parts optional and then decides per vendor which
    one is mandatory (MasterRequestDTO.validateInputFilesForProcess, PAYMENT_METHOD_SETUP):

      * ADP sends `direct_deposit_file` — its Direct Deposit export, which the API
        parses with the SAME EmployeeMasterADP bindings as the census, reading
        `Routing Number`, `Account Number`, `Deposit Type`, `Deposit Amount` and
        `Deposit Percent`. The Payment Method Sanity tool's Corrected Source carries
        exactly those columns, so it is the file to send.
      * PAYCOM sends `employee_census_adv_file`, because Paycom keeps the bank columns
        inside the census itself and the API reads the payment data from there.

    Returns the raw requests.Response so the caller can inspect status/body.
    """
    url = f"{prod_host.rstrip('/')}/app/onboarding/api/employee/payment-method"
    headers = {"Accept": "application/json", "AuthorizationHeader": jwt_token}
    field = "direct_deposit_file" if str(vendor).upper() == "ADP" else "employee_census_adv_file"
    try:
        return requests.post(url, headers=headers,
                             files={field: (payment_filename, payment_bytes)},
                             timeout=timeout)
    except requests.RequestException as e:
        raise OnboardingAPIError(f"Could not reach {url}: {e}") from e


def push_employee_deductions(jwt_token: str,
                             deduction_bytes: bytes, deduction_filename: str,
                             mapping_bytes: bytes, mapping_filename: str,
                             prod_host: str = DEFAULT_PROD_HOST, timeout: int = 120):
    """POST {prod_host}/app/onboarding/api/employee/deductions (multipart).

    BOTH parts are mandatory for both vendors — the endpoint declares neither
    optional and `MasterRequestDTO.validateInputFilesForProcess(EMPLOYEE_DEDUCTIONS)`
    requires them without a vendor branch. The mapping is what turns a source
    deduction description into a Uzio deduction code, and the API looks it up by
    description EXACTLY, so an unmapped description silently takes its employees
    with it.

    Returns the raw requests.Response so the caller can inspect status/body.
    """
    url = f"{prod_host.rstrip('/')}/app/onboarding/api/employee/deductions"
    headers = {"Accept": "application/json", "AuthorizationHeader": jwt_token}
    files = {
        "employee_deduction_file": (deduction_filename, deduction_bytes),
        "deduction_mapping_file": (mapping_filename, mapping_bytes),
    }
    try:
        return requests.post(url, headers=headers, files=files, timeout=timeout)
    except requests.RequestException as e:
        raise OnboardingAPIError(f"Could not reach {url}: {e}") from e


def automation_id(resp) -> int | None:
    """The run id (onboarding_automation_history.id) from a push response, if any."""
    try:
        rid = (resp.json().get("data") or {}).get("OnboardingAutomationId")
        return int(rid) if rid is not None else None
    except (ValueError, AttributeError, TypeError):
        return None


def query(jwt_token: str, sql: str, prod_host: str = DEFAULT_PROD_HOST,
          timeout: int = 60) -> list[dict]:
    """One SELECT against the read-only /app/onboarding/query endpoint (first page)."""
    url = f"{prod_host.rstrip('/')}/app/onboarding/query"
    headers = {"Accept": "application/json", "AuthorizationHeader": jwt_token}
    try:
        resp = requests.post(url, json={"sql": sql, "page": 0, "size": 1}, headers=headers, timeout=timeout)
    except requests.RequestException as e:
        raise OnboardingAPIError(f"Could not reach {url}: {e}") from e
    if not resp.ok:
        raise OnboardingAPIError(f"Run log lookup failed (HTTP {resp.status_code}): {resp.text[:300]}")
    try:
        return resp.json().get("data") or []
    except (ValueError, AttributeError):
        raise OnboardingAPIError(f"Run log lookup returned something that isn't JSON: {resp.text[:300]}")


def fetch_run_log(jwt_token: str, run_id: int, prod_host: str = DEFAULT_PROD_HOST,
                  wait_seconds: int = 15) -> dict:
    """One run's row from onboarding_automation_history.

    The API writes error_messages / optional_validations together with end_time as
    its last step before answering the push, so they are normally there already;
    the short wait only covers a row read back before that write is visible.
    Returns the row even if end_time is still empty -- the caller says so.
    """
    deadline = time.time() + wait_seconds
    while True:
        rows = query(jwt_token, RUN_LOG_SQL.format(run_id=int(run_id)), prod_host)
        if not rows:
            raise OnboardingAPIError(f"No run with id {run_id} in the onboarding log.")
        if rows[0].get("end_time") or time.time() >= deadline:
            return rows[0]
        time.sleep(3)


def find_run_since(jwt_token: str, fein: str, username: str, since_utc: datetime,
                   prod_host: str = DEFAULT_PROD_HOST) -> dict | None:
    """The FIRST run this user started for this FEIN just after `since_utc`, or None.

    Used when a push gets no answer: a big census outlasts the request timeout
    (run 1376 -- 873 employees -- took 8 minutes) but Uzio keeps processing, and
    its row exists from the moment processing starts. First rather than newest,
    because a later run by the same user (another module) must not be taken for it
    -- a newest-first lookup for run 1376 returned the Payment run 1377 that
    followed it. The window closes after LOOKUP_WINDOW for the same reason: when
    the push never arrived, an unrelated later run for that employer must not be
    reported as its outcome.
    FEIN and username are checked against a strict character set before they go
    into the SQL, so no quote can reach the query.
    """
    digits = re.sub(r"\D", "", fein or "")
    user = (username or "").strip().lower()
    if not digits or not re.fullmatch(r"[a-z0-9._%+@-]+", user):
        raise OnboardingAPIError("Can't look up the run: the FEIN or username has unexpected characters.")
    start = (since_utc - LOOKUP_SKEW).astimezone(timezone.utc)
    until = (since_utc + LOOKUP_WINDOW).astimezone(timezone.utc)
    rows = query(jwt_token,
                 f"select {RUN_LOG_COLS} from onboarding_automation_history "
                 f"where fein = '{digits}' and lower(created_by) = '{user}' "
                 f"and start_time >= '{start:%Y-%m-%d %H:%M:%S}' "
                 f"and start_time < '{until:%Y-%m-%d %H:%M:%S}' "
                 f"order by id asc limit 1", prod_host)
    return rows[0] if rows else None


def new_result(fein: str, user: str) -> dict:
    """The shape every caller passes around: one push attempt and what came of it."""
    return {"ok": False, "status": None, "text": "", "push_error": None,
            "run_id": None, "run": None, "run_error": None, "found_by_lookup": False,
            "fein": (fein or "").strip(), "user": (user or "").strip(),
            "pushed_at": datetime.now(timezone.utc).isoformat()}


def lost_answer(result: dict) -> bool:
    """The push got no usable answer, so Uzio may still be processing it."""
    return result.get("push_error") is not None or result.get("status") in GATEWAY_TIMEOUTS


def attach_run_log(result: dict, jwt_token: str, prod_host: str = DEFAULT_PROD_HOST) -> dict:
    """Fill result["run"] -- and result["run_id"] when the push gave no answer."""
    try:
        if result.get("run_id") is not None:
            wait = 0 if result.get("found_by_lookup") else 15
            result["run"] = fetch_run_log(jwt_token, result["run_id"], prod_host, wait_seconds=wait)
        elif lost_answer(result):
            row = find_run_since(jwt_token, result["fein"], result["user"],
                                 datetime.fromisoformat(result["pushed_at"]), prod_host)
            if row:
                result.update(run_id=int(row["id"]), run=row, found_by_lookup=True)
        result["run_error"] = None
    except OnboardingAPIError as e:
        result["run_error"] = str(e)
    return result


def needs_recheck(result: dict) -> bool:
    """Something is still missing that a later look could fill in."""
    if result.get("run_error"):
        return True
    if result.get("run_id") is None:
        return lost_answer(result)
    return not (result.get("run") or {}).get("end_time")


def parse_ts(value):
    """DB timestamps come back as ISO strings; the column is UTC without a zone."""
    try:
        d = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    return d if d.tzinfo else d.replace(tzinfo=timezone.utc)


def ist(value) -> str:
    d = parse_ts(value)
    return d.astimezone(IST).strftime("%d-%b %H:%M") if d else "?"


def run_issues(blob) -> list[dict]:
    """error_messages / optional_validations -> one dict per employee row.

    The column holds JSON text {"<Module>": [{rowNumber, employeeId, errorType,
    errorDetails}, ...]}. When the API could not serialise its failures it holds a
    plain sentence instead, which is kept as a single row rather than dropped.
    """
    if not blob:
        return []
    obj = blob
    if isinstance(blob, str):
        try:
            obj = json.loads(blob)
        except ValueError:
            return [{"Module": "-", "Row": None, "Employee ID": None, "Type": None, "Reason": blob.strip()}]
    out = []
    if isinstance(obj, dict):
        for module, items in obj.items():
            for it in items or []:
                if isinstance(it, dict):
                    out.append({"Module": module, "Row": it.get("rowNumber"),
                                "Employee ID": it.get("employeeId"), "Type": it.get("errorType"),
                                "Reason": (it.get("errorDetails") or "").strip()})
                else:
                    out.append({"Module": module, "Row": None, "Employee ID": None,
                                "Type": None, "Reason": str(it)})
    return out


def group_issues(issues: list[dict], sample: int = 5) -> list[dict]:
    """One row per distinct reason, biggest first, with the affected Employee IDs."""
    groups = {}
    for i in issues:
        groups.setdefault((i["Module"], i["Reason"]), []).append(str(i["Employee ID"] or "-"))
    rows = []
    for (module, reason), ids in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        shown = ", ".join(ids[:sample]) + (f" (+{len(ids) - sample} more)" if len(ids) > sample else "")
        rows.append({"Employees": len(ids), "Module": module, "Reason": reason, "Employee IDs": shown})
    return rows
