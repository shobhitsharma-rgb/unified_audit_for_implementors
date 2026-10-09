"""Read-only access to the onboarding DB through /app/onboarding/query (PHIX-98714).

This is the transport only: sign in, run one SELECT, hand back rows. It is kept
apart from utils/onboarding_core.py on purpose — that module can also PUSH a
census, and the implementors build must never carry the push path. A logs reader
needs none of it.

The endpoint runs a single SELECT in a read-only transaction and audits it
against whoever signed in, so each user brings their own credentials. Nothing is
cached or written to disk here; the token lives only as long as the caller keeps
it.

The token is minted against one FEIN, but the rows it returns are not limited to
that employer — any FEIN the caller has access to is enough to read the log.
"""
from datetime import datetime, timedelta, timezone

import requests

DEFAULT_HOST = "https://api.uzio.com"
IST = timezone(timedelta(hours=5, minutes=30))

TABLE = "onboarding_automation_history"
# Listing columns. error_messages / optional_validations are left out on purpose:
# one row can carry 3 MB of them, and response_body already holds the counts.
LIGHT_COLUMNS = ("id, vendor, fein, start_time, end_time, created_by, created_date, "
                 "response_body")

# A finished run that touched no employees at all. Not a pass, not a failure.
NOTHING_PROCESSED = "Nothing processed"

MODULE_SHORT = {
    "EmployeeCensus": "Census", "PaymentMethodSetup": "Payment",
    "FedTaxWithholding": "FedTax", "StateTaxWithholding": "StateTax",
    "EmployeeDeductions": "Deductions", "EmployeeContributions": "Contributions",
    "WorkerCompensation": "WorkersComp", "PriorPayroll": "PriorPayroll",
    "CompanyJobTitle": "JobTitle", "SocCode": "SocCode", "W2DeliveryMethod": "W2Delivery",
}

# Real prod rows belonging to test employers; they have no client name anywhere.
SANDBOX_FEINS = {
    "232332223": "AA prod (sandbox)", "927387483": "AA prod 01 (sandbox)",
    "769465445": "A Mobile Company 1 (sandbox)", "990000001": "DSP 101 (sandbox)",
    "232432324": "DSP 102 (sandbox)", "876876982": "DSP Test (sandbox)",
    "991182990": "DSP Trial (sandbox)", "565656565": "Vatica Health Sandbox (sandbox)",
}


class OnboardingQueryError(Exception):
    """Raised when login or the query endpoint refuses, or cannot be reached."""


def login(username: str, password: str, fein: str, host: str = DEFAULT_HOST,
          timeout: int = 30) -> str:
    """POST {host}/app/onboarding/token -> JWT."""
    url = f"{host.rstrip('/')}/app/onboarding/token"
    try:
        resp = requests.post(url, json={"username": username, "password": password,
                                        "fein": fein}, timeout=timeout)
    except requests.RequestException as e:
        raise OnboardingQueryError(
            f"Could not reach {url}: {e}. If you are off the office network or the VPN "
            "is down, this endpoint is not reachable from here.") from e
    if not resp.ok:
        raise OnboardingQueryError(f"Login failed (HTTP {resp.status_code}): {resp.text[:300]}")

    token = resp.text.strip()
    if token.startswith("{"):
        try:
            body = resp.json()
            token = body.get("token") or body.get("access_token") or body.get("jwt") or ""
        except ValueError:
            token = ""
    if not token or token.count(".") != 2:
        raise OnboardingQueryError(
            "Login succeeded but no token came back. Check the FEIN — it has to be one "
            "you have access to.")
    return token


def query(token: str, sql: str, host: str = DEFAULT_HOST, size: int = 200,
          max_pages: int = 50, timeout: int = 180) -> list:
    """Run one SELECT and return every page of rows.

    A query needing more than `max_pages` is too broad to be read on a screen, so
    it stops there rather than pulling production data indefinitely.
    """
    url = f"{host.rstrip('/')}/app/onboarding/query"
    headers = {"Accept": "application/json", "Content-Type": "application/json",
               "AuthorizationHeader": token}
    rows = []
    for page in range(max_pages):
        try:
            resp = requests.post(url, json={"sql": sql, "page": page, "size": size},
                                 headers=headers, timeout=timeout)
        except requests.RequestException as e:
            raise OnboardingQueryError(f"Could not reach {url}: {e}") from e
        if resp.status_code in (401, 403):
            raise OnboardingQueryError(
                "The query endpoint rejected your token (HTTP "
                f"{resp.status_code}). Sign in again, or ask whether your account is "
                "allowed to use it.")
        if not resp.ok:
            raise OnboardingQueryError(f"Query failed (HTTP {resp.status_code}): {resp.text[:300]}")
        try:
            body = resp.json()
        except ValueError:
            raise OnboardingQueryError(f"The endpoint returned something that is not JSON: "
                                       f"{resp.text[:200]}")
        rows.extend(body.get("data") or [])
        if not body.get("hasMore"):
            break
    return rows


# ------------------------------------------------------------------ formatting

def parse_ts(value):
    """The DB columns are UTC without a zone; make that explicit."""
    if not value:
        return None
    try:
        parsed = datetime.fromisoformat(str(value).replace("Z", "+00:00"))
    except ValueError:
        return None
    return parsed if parsed.tzinfo else parsed.replace(tzinfo=timezone.utc)


def ist(value, fmt: str = "%d-%b %H:%M") -> str:
    parsed = parse_ts(value)
    return parsed.astimezone(IST).strftime(fmt) if parsed else ""


def day_bounds_utc(start_day, end_day) -> tuple:
    """IST calendar days -> the UTC range to compare start_time against.

    Everyone reads these logs in IST while the column is UTC, so a filter for
    "8 Oct" has to mean the IST day, not the UTC one.
    """
    begin = datetime.combine(start_day, datetime.min.time(), tzinfo=IST)
    finish = datetime.combine(end_day, datetime.min.time(), tzinfo=IST) + timedelta(days=1)
    fmt = "%Y-%m-%d %H:%M:%S"
    return (begin.astimezone(timezone.utc).strftime(fmt),
            finish.astimezone(timezone.utc).strftime(fmt))


def duration(start, end) -> str:
    first, last = parse_ts(start), parse_ts(end)
    if not (first and last):
        return ""
    seconds = int((last - first).total_seconds())
    if seconds < 3600:
        return f"{seconds // 60}:{seconds % 60:02d}"
    return f"{seconds // 3600}h{seconds % 3600 // 60:02d}"


def summarize(row: dict) -> dict:
    """Counts and a status for one run, read out of response_body.

    An empty end_time means the run is still going or died without writing a
    result; the age says which is more likely. A run that finished having
    processed nobody is reported as such rather than as a success.
    """
    import json

    try:
        body = json.loads(row.get("response_body") or "") or {}
    except (ValueError, TypeError):
        body = {}
    totals = body.get("TotalMap") or {}
    passed = body.get("SuccessMap") or {}
    failed = body.get("FailureMap") or {}
    fail_count = sum(int(v or 0) for v in failed.values())

    processed = sum(int(v or 0) for v in totals.values())
    if row.get("end_time"):
        if processed == 0:
            # Finished, but nobody went in: no employees matched, or the file was
            # empty. Calling that OK reads as "it worked", which it did not.
            status = NOTHING_PROCESSED
        else:
            status = "OK" if fail_count == 0 else f"FAIL {fail_count}"
    else:
        started = parse_ts(row.get("start_time"))
        minutes = int((datetime.now(timezone.utc) - started).total_seconds() // 60) if started else 0
        if minutes <= 120:
            status = f"RUNNING {minutes}m"
        elif minutes < 48 * 60:
            status = f"NO RESULT {minutes // 60}h"
        else:
            status = f"NO RESULT {minutes // 1440}d"

    return {
        "modules": list(totals.keys()),
        "total": sum(int(v or 0) for v in totals.values()) if totals else None,
        "ok": sum(int(v or 0) for v in passed.values()) if passed else None,
        "fail": fail_count if totals else None,
        "totals": totals, "passed": passed, "failed": failed,
        "duration": duration(row.get("start_time"), row.get("end_time")),
        "status": status,
    }


def issue_rows(blob) -> list:
    """error_messages / optional_validations -> one dict per employee row.

    The column holds {"<Module>": [{rowNumber, employeeId, errorType, errorDetails}]}.
    When the API could not serialise its failures it holds a plain sentence
    instead, which is kept as a single row rather than dropped.
    """
    import json

    try:
        parsed = json.loads(blob) if blob else None
    except (ValueError, TypeError):
        parsed = None
    out = []
    if isinstance(parsed, dict):
        for module, items in parsed.items():
            for item in items or []:
                if isinstance(item, dict):
                    out.append({"Module": MODULE_SHORT.get(module, module),
                                "Row": item.get("rowNumber"),
                                "Employee ID": item.get("employeeId"),
                                "Type": item.get("errorType"),
                                "Reason": (item.get("errorDetails") or "").strip()})
                else:
                    out.append({"Module": MODULE_SHORT.get(module, module), "Row": None,
                                "Employee ID": None, "Type": None, "Reason": str(item)})
    elif parsed:
        out.append({"Module": "", "Row": None, "Employee ID": None, "Type": None,
                    "Reason": str(parsed)})
    return out


def group_issues(issues: list, sample: int = 3) -> list:
    """The same reason hit by 200 employees is one line, not 200."""
    from collections import Counter

    counts = Counter((i["Module"], i["Type"], i["Reason"]) for i in issues)
    grouped = []
    for (module, kind, reason), count in counts.most_common():
        who = [str(i["Employee ID"]) for i in issues
               if (i["Module"], i["Type"], i["Reason"]) == (module, kind, reason)
               and i["Employee ID"]][:sample]
        grouped.append({"Employees": count, "Module": module, "Type": kind or "",
                        "Reason": reason,
                        "e.g.": ", ".join(who) + (" …" if count > len(who) else "")})
    return grouped


def digits(value, length: int = 9) -> str:
    """A FEIN as the column stores it, or "" when the text is not one."""
    kept = "".join(c for c in str(value or "") if c.isdigit())
    return kept if len(kept) == length else ""


def sql_literal(value, limit: int = 64) -> str:
    """A user's text, safe to put inside quotes in a LIKE or = comparison."""
    kept = [c for c in str(value or "").strip()[:limit]
            if c.isalnum() or c in " ._@+-"]
    return "".join(kept)
