#!/usr/bin/env python3
"""
Export Netskope application events (category = GenerativeAI) to CSV.

Endpoint : GET /api/v2/events/data/application
Docs     : Netskope REST API v2 - Event Data / DataSearch

Dependencies: requests, pandas (stdlib: typing, time, csv, json)
"""

import csv
import json
import time
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# CONFIGURATION  -- edit this block only
# ---------------------------------------------------------------------------

# Your API token. Netskope issues these under Settings > Tools > REST API v2.
API_TOKEN: str = "PUT_YOUR_API_TOKEN_HERE"

# Tenant base URL, no trailing slash. e.g. "https://myorg.goskope.com"
BASE_URL: str = "https://YOUR_TENANT.goskope.com"

# Endpoint path (kept separate so other event types are a one-line change).
ENDPOINT_PATH: str = "/api/v2/events/data/application"

# Authentication style.
#   "bearer"   -> Authorization: Bearer <token>        (requested / generic)
#   "netskope" -> Netskope-Api-Token: <token>          (what Netskope expects)
# If you get 401s with "bearer", flip this to "netskope".
AUTH_MODE: str = "bearer"

# Output file. If DATE_STAMP_FILENAME is True the day being exported is
# appended, e.g. GenerativeAI_Applications_2026-07-19.csv -- useful when a
# folder of dated files is being picked up downstream.
OUTPUT_FILE: str = "GenerativeAI_Applications.csv"
DATE_STAMP_FILENAME: bool = True

# --- Which single day to pull ---------------------------------------------
# DAY_MODE:
#   "yesterday" -> the full previous calendar day, 00:00:00 to 23:59:59 local
#   "today"     -> midnight local through right now
#   "date"      -> the specific calendar day named in TARGET_DATE
#   "rolling"   -> the trailing 24 hours from this moment
DAY_MODE: str = "yesterday"

# Only used when DAY_MODE == "date". Format: YYYY-MM-DD
TARGET_DATE: str = "2026-07-19"

# One calendar day in seconds.
ONE_DAY: int = 86400

# Where the day ends.
#   0 -> 00:00:00 yesterday through 00:00:00 today (a clean 86400s span)
#   1 -> stops at 23:59:59 instead
# Netskope treats endtime as INCLUSIVE, so with 0 an event landing exactly on
# 00:00:00 appears in both this file and tomorrow's. Set to 1 if the
# downstream consumer can't tolerate that duplicate.
BOUNDARY_TRIM: int = 0

# Records per page. Netskope caps this at 10000 for event data.
PAGE_SIZE: int = 5000

# Offset pagination on this API becomes unreliable past roughly 10k records
# in a single query (the backend re-sorts between pages, producing duplicates
# and gaps). To stay correct we split the lookback into windows and paginate
# inside each window. Set to 0 to disable windowing and use one flat query.
TIME_WINDOW_SECONDS: int = 3600

# Network behaviour.
REQUEST_TIMEOUT: int = 60      # seconds per HTTP request
MAX_RETRIES: int = 5           # retries for 429 / 5xx / network errors
BACKOFF_BASE: float = 2.0      # exponential backoff base, seconds

# Print the exact request (URL + params) and a snippet of every raw response
# to the console. Turn this on FIRST if you're getting zero records -- it
# will show you immediately whether the problem is auth, the filter, or the
# response shape not matching what the script expects.
DEBUG: bool = True

# Name of the pagination query parameter. Netskope's own documented example
# for this endpoint family is "skip" (e.g. limit=100&skip=0). If your tenant
# actually uses "offset" instead, change it here.
OFFSET_PARAM_NAME: str = "skip"

# The mandatory filter. Netskope stores this value lowercased.
# category is frequently a MULTI-VALUE field on an app record (an app can
# sit in several categories at once), and "eq" only matches when the field
# holds exactly that one value -- against a list it typically matches
# nothing, which is the most common reason this filter silently returns
# zero rows. "in" checks membership instead, so it matches whether the
# field is a single value or a list containing it. If your tenant's schema
# really does store category as a single scalar, "equals" is still fine.
BASE_FILTER_FIELD: str = "category"
BASE_FILTER_VALUE: str = "Generative AI"
BASE_FILTER_OPERATOR: str = "in"

# ---------------------------------------------------------------------------
# FILTER SYNTAX
# ---------------------------------------------------------------------------
# Netskope Query Language (NQL) is passed as the `query` QUERY STRING
# parameter -- this endpoint is a GET and takes no JSON body.
#
#   query=category eq 'generativeai' and appname eq 'ChatGPT'
#
# Terms are combined with `and` / `or`. If your tenant's dialect differs,
# OPERATOR_MAP below is the only place you need to change.
#
# Each entry maps a friendly operator name to a template. {f} = field,
# {v} = the quoted value.
# ---------------------------------------------------------------------------

OPERATOR_MAP: Dict[str, str] = {
    "equals":      "{f} eq {v}",
    "notequals":   "{f} ne {v}",
    "contains":    "{f} like {v}",        # NQL `like` supports * wildcards
    "startswith":  "{f} like {v}",        # value gets a trailing *
    "endswith":    "{f} like {v}",        # value gets a leading *
    "in":          "{f} in {v}",          # value = comma separated list
    "greaterthan": "{f} gt {v}",
    "lessthan":    "{f} lt {v}",
}

# Response keys that might hold the record array, in priority order.
RESULT_KEYS: Tuple[str, ...] = ("result", "data", "events", "records", "items")

# Response keys that might hold a pagination cursor / next pointer.
CURSOR_KEYS: Tuple[str, ...] = (
    "next", "nextUrl", "next_url", "next_token", "nextToken",
    "cursor", "scroll_id", "wait_for_marker",
)


# ---------------------------------------------------------------------------
# Filter construction
# ---------------------------------------------------------------------------

def quote_value(value: str) -> str:
    """Single-quote a value for NQL, escaping any embedded quotes."""
    return "'" + str(value).replace("'", "\\'") + "'"


def build_term(field: str, operator: str, value: str) -> str:
    """
    Turn one (field, operator, value) triple into an NQL term.

    Raises ValueError if the operator is not recognised.
    """
    key = operator.strip().lower().replace("_", "").replace(" ", "")
    if key not in OPERATOR_MAP:
        raise ValueError(
            f"Unsupported operator '{operator}'. "
            f"Supported: {', '.join(sorted(OPERATOR_MAP))}"
        )

    raw = str(value).strip()

    # Wildcard shaping for the pattern-matching operators.
    if key == "contains":
        raw = f"*{raw}*"
    elif key == "startswith":
        raw = f"{raw}*"
    elif key == "endswith":
        raw = f"*{raw}"

    if key == "in":
        # "a, b, c" -> ('a','b','c')
        parts = [quote_value(p.strip()) for p in raw.split(",") if p.strip()]
        quoted = "[" + ", ".join(parts) + "]"
    else:
        quoted = quote_value(raw)

    return OPERATOR_MAP[key].format(f=field.strip(), v=quoted)


def build_query(extra: Optional[Tuple[str, str, str]]) -> str:
    """
    Build the full NQL query: the mandatory category filter, ANDed with the
    user-supplied filter if one was given.
    """
    terms = [build_term(BASE_FILTER_FIELD, BASE_FILTER_OPERATOR, BASE_FILTER_VALUE)]
    if extra:
        terms.append(build_term(*extra))
    return " and ".join(terms)


def prompt_for_filter() -> Optional[Tuple[str, str, str]]:
    """
    Ask the user for an additional filter. Returns None if they skip it.
    Re-prompts on an unknown operator rather than failing the whole run.
    """
    print("\nAdd a second filter (ANDed with category = GenerativeAI).")
    print("Press Enter at the field prompt to skip.\n")

    field = input("  Field name (e.g. appname, user, app): ").strip()
    if not field:
        print("  No extra filter -- pulling all GenerativeAI events.")
        return None

    print(f"  Operators: {', '.join(sorted(OPERATOR_MAP))}")
    while True:
        operator = input("  Operator [equals]: ").strip() or "equals"
        key = operator.lower().replace("_", "").replace(" ", "")
        if key in OPERATOR_MAP:
            break
        print(f"  '{operator}' is not supported. Try again.")

    value = input("  Value (e.g. ChatGPT): ").strip()
    if not value:
        print("  Empty value -- skipping the extra filter.")
        return None

    return field, operator, value


# ---------------------------------------------------------------------------
# HTTP layer
# ---------------------------------------------------------------------------

def build_session() -> requests.Session:
    """Create a session with auth headers applied once."""
    session = requests.Session()
    if AUTH_MODE == "netskope":
        session.headers["Netskope-Api-Token"] = API_TOKEN
    else:
        session.headers["Authorization"] = f"Bearer {API_TOKEN}"
    session.headers["Accept"] = "application/json"
    return session


class ApiError(RuntimeError):
    """Raised for non-recoverable API problems."""


def request_page(
    session: requests.Session,
    url: str,
    params: Optional[Dict[str, Any]],
    stats: Dict[str, int],
) -> Dict[str, Any]:
    """
    Perform one GET with retry on 429, 5xx and transient network errors.

    Returns the decoded JSON body. Raises ApiError on auth failure, bad
    request, or exhausted retries.
    """
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            stats["calls"] += 1
            if DEBUG:
                print(f"  [DEBUG] GET {url}")
                print(f"  [DEBUG] params: {params}")
                print(f"  [DEBUG] headers: {dict(session.headers)}")
            response = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if DEBUG:
                print(f"  [DEBUG] status: {response.status_code}")
                print(f"  [DEBUG] body  : {response.text[:800]}")

        except requests.exceptions.Timeout:
            wait = BACKOFF_BASE ** attempt
            print(f"  ! Timeout after {REQUEST_TIMEOUT}s "
                  f"(attempt {attempt}/{MAX_RETRIES}) -- retrying in {wait:.0f}s")
            time.sleep(wait)
            continue

        except requests.exceptions.ConnectionError as exc:
            wait = BACKOFF_BASE ** attempt
            print(f"  ! Connection error: {exc} "
                  f"(attempt {attempt}/{MAX_RETRIES}) -- retrying in {wait:.0f}s")
            time.sleep(wait)
            continue

        # --- Rate limited -------------------------------------------------
        if response.status_code == 429:
            # Netskope returns remaining-quota headers; honour Retry-After
            # when present, otherwise back off exponentially.
            retry_after = response.headers.get("Retry-After")
            wait = float(retry_after) if retry_after and retry_after.isdigit() \
                else BACKOFF_BASE ** attempt
            print(f"  ! Rate limited (429) -- sleeping {wait:.0f}s "
                  f"(attempt {attempt}/{MAX_RETRIES})")
            time.sleep(wait)
            continue

        # --- Auth ---------------------------------------------------------
        if response.status_code in (401, 403):
            raise ApiError(
                f"Authentication failed (HTTP {response.status_code}). "
                f"Check API_TOKEN, token scopes, and AUTH_MODE "
                f"(currently '{AUTH_MODE}' -- try the other value). "
                f"Body: {response.text[:300]}"
            )

        # --- Bad request: usually a malformed query --------------------
        if response.status_code == 400:
            raise ApiError(
                f"HTTP 400 -- the API rejected the request. This is most often "
                f"a filter syntax problem; check OPERATOR_MAP and your field "
                f"name. Body: {response.text[:500]}"
            )

        if response.status_code == 404:
            raise ApiError(
                f"HTTP 404 for {url} -- check BASE_URL and ENDPOINT_PATH."
            )

        # --- Server side --------------------------------------------------
        if response.status_code >= 500:
            wait = BACKOFF_BASE ** attempt
            print(f"  ! Server error {response.status_code} "
                  f"(attempt {attempt}/{MAX_RETRIES}) -- retrying in {wait:.0f}s")
            time.sleep(wait)
            continue

        if not response.ok:
            raise ApiError(
                f"Unexpected HTTP {response.status_code}: {response.text[:300]}"
            )

        # --- Success ------------------------------------------------------
        if not response.text.strip():
            print("  ! Empty response body -- treating as no more data.")
            return {}

        try:
            return response.json()
        except json.JSONDecodeError:
            raise ApiError(
                f"Response was not valid JSON: {response.text[:300]}"
            )

    raise ApiError(f"Giving up after {MAX_RETRIES} retries for {url}")


def extract_records(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    Pull the record list out of a response, whatever key it hides under.

    If the payload is non-empty but none of RESULT_KEYS match, this is
    almost always the reason "no data" happens with a healthy HTTP 200 --
    the response shape doesn't match what we're looking for. Rather than
    silently returning [], we scan every top-level value for the first
    list we can find, and warn loudly either way so it's never a silent
    failure.
    """
    if not payload:
        return []

    if isinstance(payload, list):
        return payload

    for key in RESULT_KEYS:
        value = payload.get(key)
        if isinstance(value, list):
            return value

    # Nothing matched the expected keys -- fall back to scanning for any
    # list value anywhere in the payload before giving up.
    for key, value in payload.items():
        if isinstance(value, list):
            print(f"  [WARN] Records found under unexpected key '{key}' "
                  f"-- consider adding it to RESULT_KEYS.")
            return value

    print(f"  [WARN] Response had no recognizable record list. "
          f"Top-level keys were: {list(payload.keys())}. "
          f"If this looks like status/metadata only, the filter or auth "
          f"is likely the issue -- check the DEBUG output above.")
    return []


def extract_cursor(payload: Dict[str, Any]) -> Optional[str]:
    """Look for a next-page pointer. Returns None if the API uses offsets."""
    if not payload:
        return None
    for key in CURSOR_KEYS:
        value = payload.get(key)
        if value and isinstance(value, str):
            return value
    return None


# ---------------------------------------------------------------------------
# Pagination
# ---------------------------------------------------------------------------

def fetch_window(
    session: requests.Session,
    query: str,
    starttime: int,
    endtime: int,
    stats: Dict[str, int],
) -> List[Dict[str, Any]]:
    """
    Retrieve every record for one time window.

    Prefers a server-supplied cursor if the response exposes one; otherwise
    falls back to limit/offset paging, stopping on a short or empty page.
    """
    url = BASE_URL.rstrip("/") + ENDPOINT_PATH
    records: List[Dict[str, Any]] = []
    offset = 0
    next_url: Optional[str] = None

    while True:
        if next_url:
            # Cursor mode: the API handed us a full URL or token.
            target = next_url if next_url.startswith("http") else url
            params = None if next_url.startswith("http") else {"cursor": next_url}
        else:
            target = url
            params = {
                "query": query,
                "starttime": starttime,
                "endtime": endtime,
                "limit": PAGE_SIZE,
                OFFSET_PARAM_NAME: offset,
            }

        payload = request_page(session, target, params, stats)

        # Netskope wraps results with its own status block, e.g.
        # {"result": [...], "status": {"execution": "SUCCESS", ...}}.
        # A 200 HTTP response can still carry a query-level failure here
        # (bad filter syntax, unknown field, etc.) -- surface it plainly.
        status_block = payload.get("status") if isinstance(payload, dict) else None
        if isinstance(status_block, dict) and status_block.get("execution") not in (None, "SUCCESS"):
            print(f"  [WARN] API reported execution="
                  f"{status_block.get('execution')!r}: "
                  f"{status_block.get('message', '')}")

        page = extract_records(payload)

        if not page:
            break

        records.extend(page)
        print(f"    page +{len(page):>5}  (window total {len(records)})")

        cursor = extract_cursor(payload)
        if cursor:
            next_url = cursor
            continue

        # Offset mode: a short page means we've reached the end.
        if len(page) < PAGE_SIZE:
            break
        offset += PAGE_SIZE

    return records


def local_midnight(epoch: float) -> int:
    """
    Snap an epoch timestamp back to 00:00:00 of that day in LOCAL time.

    Netskope stores starttime/endtime as epoch seconds, so a "calendar day"
    is defined by your machine's timezone. If you need UTC day boundaries
    instead, swap time.localtime for time.gmtime and time.mktime for
    calendar.timegm.
    """
    parts = list(time.localtime(epoch))
    parts[3] = parts[4] = parts[5] = 0   # hour, minute, second
    parts[8] = -1                        # let mktime work out DST
    return int(time.mktime(time.struct_time(parts)))


def resolve_day_range() -> Tuple[int, int, str]:
    """
    Work out the single day to export.

    Returns (starttime, endtime, label) where label is a YYYY-MM-DD string
    used for the filename and console output.
    """
    now = time.time()

    if DAY_MODE == "today":
        start = local_midnight(now)
        end = int(now)

    elif DAY_MODE == "yesterday":
        start = local_midnight(now) - ONE_DAY
        end = local_midnight(now) - BOUNDARY_TRIM

    elif DAY_MODE == "date":
        # strptime gives us the naive date; mktime turns it into local epoch.
        parsed = time.strptime(TARGET_DATE, "%Y-%m-%d")
        start = int(time.mktime(parsed))
        end = start + ONE_DAY - BOUNDARY_TRIM

    elif DAY_MODE == "rolling":
        end = int(now)
        start = end - ONE_DAY

    else:
        raise ValueError(
            f"DAY_MODE '{DAY_MODE}' is not valid. "
            f"Use: yesterday, today, date, rolling"
        )

    label = time.strftime("%Y-%m-%d", time.localtime(start))
    return start, end, label


def prompt_for_date_range() -> Tuple[int, int, str]:
    """
    Ask the user which day to export. Enter is a plain date pull for the
    calendar day 00:00:00 -> 00:00:00 next day. Leaving it blank falls back
    to the DAY_MODE constant at the top of the script.
    """
    print("\nWhich day to extract?")
    print(f"  Press Enter to use the default ({DAY_MODE}).")
    entry = input("  Start date (YYYY-MM-DD): ").strip()

    if not entry:
        return resolve_day_range()

    try:
        parsed = time.strptime(entry, "%Y-%m-%d")
    except ValueError:
        print(f"  '{entry}' isn't a valid date (expected YYYY-MM-DD). "
              f"Falling back to the default.")
        return resolve_day_range()

    start = int(time.mktime(parsed))
    end = start + ONE_DAY - BOUNDARY_TRIM
    return start, end, entry


def resolve_output_path(label: str) -> str:
    """Append the day being exported to the filename when configured."""
    if not DATE_STAMP_FILENAME:
        return OUTPUT_FILE
    if OUTPUT_FILE.lower().endswith(".csv"):
        return f"{OUTPUT_FILE[:-4]}_{label}.csv"
    return f"{OUTPUT_FILE}_{label}.csv"


def fetch_all(
    session: requests.Session,
    query: str,
    start: int,
    now: int,
    stats: Dict[str, int],
) -> List[Dict[str, Any]]:
    """
    Retrieve everything across the given day, subdividing into time windows
    to keep offset pagination correct, and deduplicating by _id.
    """
    if TIME_WINDOW_SECONDS > 0:
        bounds = list(range(start, now, TIME_WINDOW_SECONDS)) + [now]
        windows = list(zip(bounds[:-1], bounds[1:]))
    else:
        windows = [(start, now)]

    seen: set = set()
    all_records: List[Dict[str, Any]] = []

    for index, (w_start, w_end) in enumerate(windows, 1):
        print(f"  window {index}/{len(windows)}  "
              f"[{w_start} -> {w_end}]")
        for record in fetch_window(session, query, w_start, w_end, stats):
            # _id is Netskope's per-event unique key. Fall back to the whole
            # record if it's absent so we never drop legitimate rows.
            key = record.get("_id") or json.dumps(record, sort_keys=True)
            if key in seen:
                continue
            seen.add(key)
            all_records.append(record)

    return all_records


# ---------------------------------------------------------------------------
# Flattening + column discovery
# ---------------------------------------------------------------------------

def flatten_record(
    obj: Any,
    prefix: str = "",
    out: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Recursively flatten nested JSON into a single-level dict.

      {"application": {"name": "ChatGPT"}}  ->  {"application.name": "ChatGPT"}

    Lists of scalars become comma-separated strings. Lists of dicts are
    expanded with an index so nothing is lost:  policies.0.name
    """
    if out is None:
        out = {}

    if isinstance(obj, dict):
        for key, value in obj.items():
            path = f"{prefix}.{key}" if prefix else str(key)
            flatten_record(value, path, out)

    elif isinstance(obj, list):
        if not obj:
            out[prefix] = ""
        elif all(not isinstance(item, (dict, list)) for item in obj):
            # Scalars -> comma joined.
            out[prefix] = ", ".join("" if i is None else str(i) for i in obj)
        else:
            for i, item in enumerate(obj):
                flatten_record(item, f"{prefix}.{i}", out)

    else:
        out[prefix] = obj

    return out


def discover_columns(rows: Iterable[Dict[str, Any]]) -> List[str]:
    """
    Union of every key across every flattened record, in first-seen order.
    No field names are hardcoded anywhere.
    """
    columns: Dict[str, None] = {}
    for row in rows:
        for key in row:
            columns.setdefault(key, None)
    return list(columns)


# ---------------------------------------------------------------------------
# Export
# ---------------------------------------------------------------------------

def export_csv(rows: List[Dict[str, Any]], columns: List[str], path: str) -> None:
    """
    Write to CSV with every discovered field as a column. Missing values are
    written blank. QUOTE_MINIMAL keeps embedded commas safe.
    """
    frame = pd.DataFrame(rows, columns=columns)
    frame = frame.where(pd.notnull(frame), "")
    frame.to_csv(path, index=False, quoting=csv.QUOTE_MINIMAL, encoding="utf-8")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    stats: Dict[str, int] = {"calls": 0}

    if "PUT_YOUR_API_TOKEN" in API_TOKEN or "YOUR_TENANT" in BASE_URL:
        print("Set API_TOKEN and BASE_URL at the top of the script first.")
        return

    extra = prompt_for_filter()

    try:
        query = build_query(extra)
    except ValueError as exc:
        print(f"Filter error: {exc}")
        return

    try:
        start, end, label = prompt_for_date_range()
    except ValueError as exc:
        print(f"Date error: {exc}")
        return

    output_path = resolve_output_path(label)

    print(f"\nQuery : {query}")
    print(f"Target: {BASE_URL.rstrip('/')}{ENDPOINT_PATH}")
    print(f"Day   : {label}  ({DAY_MODE})")
    print(f"Range : {start} -> {end}  "
          f"[{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(start))} to "
          f"{time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end))}]\n")

    session = build_session()

    try:
        records = fetch_all(session, query, start, end, stats)
    except ApiError as exc:
        print(f"\nFAILED: {exc}")
        return
    except KeyboardInterrupt:
        print("\nInterrupted by user.")
        return

    if not records:
        print("\nNo records matched. Nothing was written.")
        print(f"API calls made: {stats['calls']}")
        return

    flat_rows = [flatten_record(r) for r in records]
    columns = discover_columns(flat_rows)
    export_csv(flat_rows, columns, output_path)

    print("\n" + "-" * 52)
    print(f"Day exported            : {label}")
    print(f"Total records retrieved : {len(records)}")
    print(f"Unique columns found    : {len(columns)}")
    print(f"API calls made          : {stats['calls']}")
    print(f"Exported to             : {output_path}")
    print("-" * 52)


if __name__ == "__main__":
    main()
