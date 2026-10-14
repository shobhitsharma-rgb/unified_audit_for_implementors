"""Read-only access to NeuronDB through the NeuronOps /query endpoint (PHIX-98325).

The endpoint takes a single SELECT and runs it in a read-only transaction, so
nothing here can change data even by accident. Every read is audited server-side
against the caller's own login, which is why this module never holds a shared
service account: each user signs in with their own credentials and their own name
is what the audit records.

Credentials are passed in and used once to mint a token; nothing is written to
disk and nothing is cached beyond the caller's own variable.
"""
import requests

DEFAULT_HOST = "https://api.uzio.com"


class NeuronOpsError(Exception):
    """Raised when login or the query endpoint refuses, or cannot be reached."""


def login(username: str, password: str, host: str = DEFAULT_HOST, timeout: int = 30) -> str:
    """POST {host}/api/auth/token -> JWT."""
    url = f"{host.rstrip('/')}/api/auth/token"
    try:
        resp = requests.post(url, json={"username": username, "password": password}, timeout=timeout)
    except requests.RequestException as e:
        raise NeuronOpsError(
            f"Could not reach {url}: {e}. If you are off the office network, this endpoint "
            "may not be reachable from here.") from e
    if not resp.ok:
        raise NeuronOpsError(f"Login failed (HTTP {resp.status_code}): {resp.text[:300]}")
    try:
        token = (resp.json() or {}).get("token")
    except ValueError:
        token = None
    if not token:
        raise NeuronOpsError(f"Login succeeded but no token came back: {resp.text[:300]}")
    return token


def query(token: str, sql: str, host: str = DEFAULT_HOST, size: int = 500,
          max_pages: int = 20, timeout: int = 120) -> list[dict]:
    """Run one SELECT and return every page of rows.

    `size` is capped at 5000 by the server. A query that needs more than
    `max_pages` pages is almost certainly too broad for this tool, so it stops
    rather than pulling prod data indefinitely.
    """
    url = f"{host.rstrip('/')}/api/neuronops/query"
    headers = {"Content-Type": "application/json",
               "Authorization": f"Bearer {token}",
               "X-Auth-Type": "bearer"}
    rows: list[dict] = []
    for page in range(max_pages):
        try:
            resp = requests.post(url, json={"sql": sql, "page": page, "size": size},
                                 headers=headers, timeout=timeout)
        except requests.RequestException as e:
            raise NeuronOpsError(f"Could not reach {url}: {e}") from e
        if resp.status_code == 401:
            raise NeuronOpsError("The login was accepted but the query endpoint rejected the token "
                                 "(HTTP 401). Sign in again, or ask whether your account is allowed "
                                 "to use the reporting endpoint.")
        if resp.status_code == 403:
            raise NeuronOpsError("Your account is not permitted to use the reporting endpoint "
                                 "(HTTP 403). Ask the platform team for access.")
        if not resp.ok:
            raise NeuronOpsError(f"Query failed (HTTP {resp.status_code}): {resp.text[:300]}")
        body = resp.json()
        rows.extend(body.get("data") or [])
        if not body.get("hasMore"):
            break
    return rows


def sql_in_list(values) -> str:
    """Quote a list of identifiers for an IN (...) clause.

    Only letters, digits and a few separators survive, so nothing a user types can
    close the quote and change the statement. Values that do not fit are dropped.
    """
    import re
    safe = [v.strip() for v in values if v and re.fullmatch(r"[A-Za-z0-9._@+-]{1,64}", v.strip())]
    return ",".join("'" + v + "'" for v in safe)


# Every Amazon DSP employer sits on this exchange; it is what makes a client list
# "the Amazon clients" rather than every employer in the database.
AMAZON_EXCHANGE_ID = "EX-20243277-1b50-4035-821d-d0fcd9b895a9"


def amazon_clients(token: str, host: str = DEFAULT_HOST) -> list:
    """[{fein, company_name, live_status, is_test_client}] for the Amazon exchange.

    Sorted by name, because every caller shows it to someone who knows the client by
    name and not by FEIN. Rows without both a name and a FEIN are dropped: they
    cannot be picked or pushed to.
    """
    rows = query(token, "select fein, company_name, live_status, is_test_client "
                 "from employer_organization where exchange_id = "
                 f"'{AMAZON_EXCHANGE_ID}' and deleted = 0 order by company_name",
                 host=host)
    return [r for r in rows if r.get("fein") and r.get("company_name")]


def is_true(value) -> bool:
    """A 0/1 column as this endpoint returns it: a STRING, so '0' is truthy in Python.

    Every boolean-ish column comes back quoted, which makes `if row["is_test_client"]`
    silently true for everyone. Compare the text instead.
    """
    return str(value).strip().lower() in ("1", "true", "yes", "y")
