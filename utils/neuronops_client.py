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
