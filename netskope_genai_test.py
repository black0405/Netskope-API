#!/usr/bin/env python3
"""
Netskope GenAI Application Export -- standalone test version.

All configuration is in the CONFIGURATION block below; there is no
config.env, no logging framework, and no screen wrapper. Edit the block,
run the file, check the CSV.

    python netskope_genai_test.py                 # yesterday
    python netskope_genai_test.py 2026-07-20      # a specific day
    python netskope_genai_test.py --peek          # show raw field names only

Dependencies: requests, pandas   (msal only if UPLOAD_TO_SHAREPOINT = True)
"""

import csv
import json
import os
import sys
import time
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import requests

# ===========================================================================
# CONFIGURATION -- edit this block
# ===========================================================================

# --- Netskope --------------------------------------------------------------
API_TOKEN: str = "PUT_YOUR_API_TOKEN_HERE"
BASE_URL: str = "https://YOUR_TENANT.goskope.com"

# "bearer"   -> Authorization: Bearer <token>
# "netskope" -> Netskope-Api-Token: <token>
# If you get 401s, try the other one.
AUTH_MODE: str = "bearer"

# The datasearch endpoint is the one that supports the `query` filter.
ENDPOINT_PATH: str = "/api/v2/events/datasearch/application"

# --- Filter ----------------------------------------------------------------
# appcategory is the CCI *application* category. `category` is the web/URL
# category and is a different taxonomy -- don't confuse them.
FILTER_FIELD: str = "appcategory"
FILTER_VALUE: str = "Generative AI"
FILTER_OPERATOR: str = "in"          # in | equals

# --- Which day -------------------------------------------------------------
# "yesterday" | "today" | "date"
DAY_MODE: str = "yesterday"
TARGET_DATE: str = "2026-07-20"      # only used when DAY_MODE = "date"

# --- Output ----------------------------------------------------------------
OUTPUT_DIR: str = "."                # "." = next to this script
OUTPUT_FILE: str = "GenerativeAI_Applications.csv"
DATE_STAMP_FILENAME: bool = True     # -> GenerativeAI_Applications_2026-07-20.csv

DATE_FORMAT: str = "%m/%d/%Y"        # 07/29/2026
MB_DECIMALS: int = 2

# --- Behaviour -------------------------------------------------------------
PAGE_SIZE: int = 5000                # Netskope caps at 10000
TIME_WINDOW_SECONDS: int = 3600      # split the day; 0 = one flat query
QUERY_TIMEOUT: int = 180             # server-side query timeout
REQUEST_TIMEOUT: int = 60            # per-HTTP-request timeout
MAX_RETRIES: int = 5
DEBUG: bool = False                  # print every request/response

# One row per user+application, byte columns summed. [] = every raw event.
DEDUPE_ON: List[str] = ["user", "app"]

# --- SharePoint (off by default for local testing) --------------------------
UPLOAD_TO_SHAREPOINT: bool = False
SPN_TENANT_ID: str = ""
SPN_CLIENT_ID: str = ""
SPN_CLIENT_SECRET: str = ""
SITE_HOSTNAME: str = "yourtenant.sharepoint.com"
SITE_PATH: str = "/sites/YourSiteName"
LIBRARY_NAME: str = "Documents"
TARGET_FOLDER: str = "Netskope/GenAI"

# ===========================================================================
# COLUMNS
# ===========================================================================

# Field name -> CSV header, in output order.
COLUMNS: List[Tuple[str, str]] = [
    ("app",               "Application"),
    ("user",              "User"),
    ("url",               "URL"),
    ("timestamp",         "Event Date"),
    ("usergroup",         "User Group"),
    ("organization_unit", "Organization Unit"),
    ("numbytes",          "Sum - Total Bytes (MB)"),
    ("server_bytes",      "Sum - Bytes Downloaded (MB)"),
    ("client_bytes",      "Sum - Bytes Uploaded (MB)"),
]

# Leftmost row-counter column. Set ROW_NUMBER_HEADER = "" for a blank header.
ADD_ROW_NUMBER: bool = True
ROW_NUMBER_HEADER: str = "S.No."

# Byte field names differ between Netskope schemas. The script asks for
# every alias and uses whichever actually comes back.
#   client_bytes = sent BY the client -> uploaded
#   server_bytes = sent BY the server -> downloaded
BYTE_ALIASES: Dict[str, List[str]] = {
    "numbytes":     ["numbytes", "numbytes_total", "total_bytes", "bytes"],
    "server_bytes": ["server_bytes", "server_bytes_total",
                     "bytes_downloaded", "download_bytes"],
    "client_bytes": ["client_bytes", "client_bytes_total",
                     "bytes_uploaded", "upload_bytes"],
}

SUM_FIELDS = set(BYTE_ALIASES)       # summed when rows collapse
BYTES_FIELDS = set(BYTE_ALIASES)     # converted bytes -> MB
EPOCH_FIELDS = {"timestamp"}         # epoch -> DATE_FORMAT

ONE_DAY = 86400

# ===========================================================================
# Query building
# ===========================================================================


def build_query() -> str:
    """
    Build the NQL filter, e.g.  appcategory in ['Generative AI']

    Passed as the `query` query-string parameter -- this endpoint is a GET
    and takes no JSON body.
    """
    value = str(FILTER_VALUE).replace("'", "\\'")
    if FILTER_OPERATOR == "in":
        return f"{FILTER_FIELD} in ['{value}']"
    return f"{FILTER_FIELD} eq '{value}'"


def fields_param() -> str:
    """
    Comma-separated `fields` list, so the API returns only what we export
    instead of ~100 columns per row. Every byte alias is requested; unknown
    names are ignored by the API.
    """
    wanted: List[str] = []
    for field, _header in COLUMNS:
        for name in BYTE_ALIASES.get(field, [field]):
            if name not in wanted:
                wanted.append(name)
    return ",".join(wanted)


# ===========================================================================
# Day range
# ===========================================================================


def local_midnight(epoch: float) -> int:
    """Snap a timestamp back to 00:00:00 that day, in LOCAL time."""
    parts = list(time.localtime(epoch))
    parts[3] = parts[4] = parts[5] = 0
    parts[8] = -1                      # let mktime work out DST
    return int(time.mktime(time.struct_time(parts)))


def resolve_day() -> Tuple[int, int, str]:
    """
    Work out the day to export: a command-line date wins, then DAY_MODE.
    Returns (start_epoch, end_epoch, label).
    """
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if args:
        parsed = time.strptime(args[0], "%Y-%m-%d")
        start = int(time.mktime(parsed))
        return start, start + ONE_DAY, args[0]

    now = time.time()
    if DAY_MODE == "today":
        start, end = local_midnight(now), int(now)
    elif DAY_MODE == "date":
        start = int(time.mktime(time.strptime(TARGET_DATE, "%Y-%m-%d")))
        end = start + ONE_DAY
    else:                                            # yesterday
        start = local_midnight(now) - ONE_DAY
        end = local_midnight(now)

    return start, end, time.strftime("%Y-%m-%d", time.localtime(start))


def output_path(label: str) -> str:
    """Full path of the CSV for a given day."""
    name = OUTPUT_FILE
    if DATE_STAMP_FILENAME and "." in OUTPUT_FILE:
        stem, _, ext = OUTPUT_FILE.rpartition(".")
        name = f"{stem}_{label}.{ext}"

    folder = OUTPUT_DIR
    if not os.path.isabs(folder):
        folder = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              folder)
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, name)


# ===========================================================================
# API
# ===========================================================================


class ApiError(RuntimeError):
    """Non-recoverable API problem."""


def build_session() -> requests.Session:
    """Session with the auth header applied once."""
    session = requests.Session()
    if AUTH_MODE == "netskope":
        session.headers["Netskope-Api-Token"] = API_TOKEN
    else:
        session.headers["Authorization"] = f"Bearer {API_TOKEN}"
    session.headers["Accept"] = "application/json"
    return session


def request_page(session: requests.Session, params: Dict[str, Any],
                 stats: Dict[str, int]) -> Dict[str, Any]:
    """One GET, with retry on 429 / 5xx / timeouts."""
    url = BASE_URL.rstrip("/") + ENDPOINT_PATH

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            stats["calls"] += 1
            if DEBUG:
                print(f"  [DEBUG] GET {url}")
                print(f"  [DEBUG] params {params}")
            response = session.get(url, params=params, timeout=REQUEST_TIMEOUT)
            if DEBUG:
                print(f"  [DEBUG] {response.status_code}: "
                      f"{response.text[:400]}")

        except (requests.exceptions.Timeout,
                requests.exceptions.ConnectionError) as exc:
            wait = 2 ** attempt
            print(f"  ! {type(exc).__name__} "
                  f"(attempt {attempt}/{MAX_RETRIES}) -- retry in {wait}s")
            time.sleep(wait)
            continue

        if response.status_code == 429:
            retry_after = response.headers.get("Retry-After", "")
            wait = int(retry_after) if retry_after.isdigit() else 2 ** attempt
            print(f"  ! rate limited -- sleeping {wait}s")
            time.sleep(wait)
            continue

        if response.status_code in (401, 403):
            raise ApiError(
                f"Auth failed (HTTP {response.status_code}). Check API_TOKEN, "
                f"and try flipping AUTH_MODE (currently '{AUTH_MODE}'). "
                f"{response.text[:200]}"
            )

        if response.status_code == 400:
            raise ApiError(
                f"HTTP 400 -- the API rejected the query. Usually a filter "
                f"syntax or field-name problem. {response.text[:400]}"
            )

        if response.status_code >= 500:
            wait = 2 ** attempt
            print(f"  ! server error {response.status_code} -- retry in {wait}s")
            time.sleep(wait)
            continue

        if not response.ok:
            raise ApiError(f"HTTP {response.status_code}: "
                           f"{response.text[:200]}")

        if not response.text.strip():
            return {}

        try:
            return response.json()
        except json.JSONDecodeError:
            raise ApiError(f"Response wasn't JSON: {response.text[:200]}")

    raise ApiError(f"Gave up after {MAX_RETRIES} retries")


def extract_records(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    """Pull the record list out, whatever key it's under."""
    if not payload:
        return []
    if isinstance(payload, list):
        return payload
    for key in ("result", "data", "events", "records", "items"):
        value = payload.get(key)
        if isinstance(value, list):
            return value
    print(f"  [WARN] no record list in response; top-level keys: "
          f"{list(payload.keys())}")
    return []


def fetch_all(session: requests.Session, query: str,
              start: int, end: int, stats: Dict[str, int]
              ) -> List[Dict[str, Any]]:
    """
    Pull every record for the day.

    The day is split into windows because offset pagination on this API
    degrades past ~10k rows in one query (the backend re-sorts between
    pages, producing duplicates and gaps). Records are deduped on _id.
    """
    if TIME_WINDOW_SECONDS > 0:
        bounds = list(range(start, end, TIME_WINDOW_SECONDS)) + [end]
        windows = list(zip(bounds[:-1], bounds[1:]))
    else:
        windows = [(start, end)]

    seen: set = set()
    all_records: List[Dict[str, Any]] = []

    for index, (w_start, w_end) in enumerate(windows, 1):
        offset = 0
        while True:
            params = {
                "query": query,
                "starttime": w_start,
                "endtime": w_end,
                "limit": PAGE_SIZE,
                "offset": offset,
                "timeout": QUERY_TIMEOUT,
                "fields": fields_param(),
            }
            page = extract_records(request_page(session, params, stats))
            if not page:
                break

            for record in page:
                key = record.get("_id") or json.dumps(record, sort_keys=True)
                if key not in seen:
                    seen.add(key)
                    all_records.append(record)

            if len(page) < PAGE_SIZE:
                break
            offset += PAGE_SIZE

        print(f"\r  window {index}/{len(windows)}  "
              f"{len(all_records):>7} records  {stats['calls']:>4} calls",
              end="", flush=True)

    print()
    return all_records


# ===========================================================================
# Shaping
# ===========================================================================


def flatten(obj: Any, prefix: str = "",
            out: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
    """Flatten nested JSON to dotted keys; scalar lists become CSV strings."""
    if out is None:
        out = {}

    if isinstance(obj, dict):
        for key, value in obj.items():
            flatten(value, f"{prefix}.{key}" if prefix else str(key), out)
    elif isinstance(obj, list):
        if not obj:
            out[prefix] = ""
        elif all(not isinstance(i, (dict, list)) for i in obj):
            out[prefix] = ", ".join("" if i is None else str(i) for i in obj)
        else:
            for i, item in enumerate(obj):
                flatten(item, f"{prefix}.{i}", out)
    else:
        out[prefix] = obj

    return out


def resolve_byte_aliases(rows: List[Dict[str, Any]]) -> None:
    """
    Normalise byte-field naming in place.

    Copies whichever alias the tenant returned onto the canonical name, so
    the export doesn't care if the schema says numbytes or numbytes_total.
    Reports what it found -- an empty byte column is then diagnosable.
    """
    if not rows:
        return

    present: set = set()
    for row in rows[:50]:
        present.update(row.keys())

    for canonical, aliases in BYTE_ALIASES.items():
        if canonical in present:
            print(f"  byte field  {canonical}: found")
            continue
        match = next((a for a in aliases if a in present), None)
        if match:
            print(f"  byte field  {canonical}: using '{match}'")
            for row in rows:
                if match in row:
                    row[canonical] = row[match]
        else:
            print(f"  byte field  {canonical}: NOT FOUND "
                  f"(tried {', '.join(aliases)}) -- will export blank")


def to_number(value: Any) -> float:
    """Numeric coercion; anything unparseable counts as 0."""
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def dedupe(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Collapse to one row per DEDUPE_ON combination, summing SUM_FIELDS.

    Keys are casefolded and whitespace-collapsed so 'Alice@Corp.com' and
    'alice@corp.com ' don't count as two different users.
    """
    if not DEDUPE_ON:
        return rows

    index: Dict[Tuple[str, ...], Dict[str, Any]] = {}
    unique: List[Dict[str, Any]] = []

    for row in rows:
        key = tuple(" ".join(str(row.get(f, "")).split()).casefold()
                    for f in DEDUPE_ON)
        kept = index.get(key)

        if kept is None:
            row = dict(row)
            for field in SUM_FIELDS:
                if field in row:
                    row[field] = to_number(row[field])
            index[key] = row
            unique.append(row)
        else:
            for field in SUM_FIELDS:
                if field in row or field in kept:
                    kept[field] = to_number(kept.get(field)) + \
                        to_number(row.get(field))

    return unique


def format_epoch(value: Any) -> Any:
    """Epoch seconds -> DATE_FORMAT. Non-epoch values pass through."""
    if value is None or value == "":
        return ""
    try:
        epoch = float(value)
    except (TypeError, ValueError):
        return value
    if epoch <= 0:
        return ""
    try:
        return time.strftime(DATE_FORMAT, time.localtime(epoch))
    except (ValueError, OSError, OverflowError):
        return value


def bytes_to_mb(value: Any) -> Any:
    """Byte count -> MB, rounded."""
    if value is None or value == "":
        return ""
    try:
        size = float(value)
    except (TypeError, ValueError):
        return value
    return round(size / (1024 * 1024), MB_DECIMALS) if size > 0 else 0


def export_csv(rows: List[Dict[str, Any]], path: str) -> None:
    """Write the CSV with exactly the configured columns and headers."""
    fields = [f for f, _h in COLUMNS]
    frame = pd.DataFrame(rows, columns=fields)
    frame = frame.where(pd.notnull(frame), "")

    for column in frame.columns:
        if column in EPOCH_FIELDS:
            frame[column] = frame[column].apply(format_epoch)
        elif column in BYTES_FIELDS:
            frame[column] = frame[column].apply(bytes_to_mb)
        elif frame[column].dtype == object:
            frame[column] = frame[column].apply(
                lambda v: v.strip() if isinstance(v, str) else v)

    frame = frame.rename(columns=dict(COLUMNS))

    if ADD_ROW_NUMBER:
        frame[ROW_NUMBER_HEADER] = list(range(1, len(frame) + 1))
        ordered = [ROW_NUMBER_HEADER] + [c for c in frame.columns
                                         if c != ROW_NUMBER_HEADER]
        frame = frame[ordered]

    frame.to_csv(path, index=False, quoting=csv.QUOTE_MINIMAL,
                 encoding="utf-8")


# ===========================================================================
# Peek mode -- what does the data actually look like?
# ===========================================================================


def peek(session: requests.Session, query: str, start: int, end: int,
         stats: Dict[str, int]) -> None:
    """
    Pull a handful of unfiltered-by-fields records and print every key,
    so you can confirm the real field names before trusting a full run.
    """
    print("\nPeek: 5 records, all fields (no `fields` restriction)\n")
    params = {
        "query": query,
        "starttime": start,
        "endtime": end,
        "limit": 5,
        "offset": 0,
        "timeout": QUERY_TIMEOUT,
    }
    rows = [flatten(r) for r in extract_records(
        request_page(session, params, stats))]

    if not rows:
        print("  No records returned for this filter/day.")
        return

    keys = sorted({k for row in rows for k in row})
    print(f"  {len(rows)} record(s), {len(keys)} distinct fields:\n")
    for key in keys:
        sample = next((row[key] for row in rows
                       if row.get(key) not in (None, "")), "")
        print(f"    {key:<34} {str(sample)[:44]}")

    print("\n  Byte-ish fields present:")
    hits = [k for k in keys if "byte" in k.lower()]
    print("   ", ", ".join(hits) if hits else "(none found)")


# ===========================================================================
# Main
# ===========================================================================


def main() -> int:
    stats: Dict[str, int] = {"calls": 0}
    started = time.time()

    if "PUT_YOUR_API_TOKEN" in API_TOKEN or "YOUR_TENANT" in BASE_URL:
        print("Set API_TOKEN and BASE_URL in the CONFIGURATION block first.")
        return 1

    query = build_query()
    start, end, label = resolve_day()

    print(f"\nDay      : {label}")
    print(f"Query    : {query}")
    print(f"Endpoint : {BASE_URL.rstrip('/')}{ENDPOINT_PATH}")

    session = build_session()

    try:
        if "--peek" in sys.argv:
            peek(session, query, start, end, stats)
            return 0

        print("\nFetching...")
        records = fetch_all(session, query, start, end, stats)
    except ApiError as exc:
        print(f"\nFAILED: {exc}")
        return 2
    except KeyboardInterrupt:
        print("\nInterrupted.")
        return 1

    if not records:
        print("\nNo records matched -- nothing written.")
        print("Try --peek to see what the endpoint actually returns.")
        return 3

    rows = [flatten(r) for r in records]
    print()
    resolve_byte_aliases(rows)

    raw_count = len(rows)
    rows = dedupe(rows)

    path = output_path(label)
    export_csv(rows, path)

    print("\n" + "-" * 54)
    print(f"Day exported     : {label}")
    print(f"Records fetched  : {raw_count}")
    if DEDUPE_ON:
        print(f"Rows exported    : {len(rows)} "
              f"(unique {'+'.join(DEDUPE_ON)})")
    print(f"API calls        : {stats['calls']}")
    print(f"Saved to         : {path}")
    print(f"Elapsed          : {time.time() - started:.0f}s")
    print("-" * 54)

    if UPLOAD_TO_SHAREPOINT:
        try:
            print("\nUploading to SharePoint...")
            print(f"  Uploaded: {upload_to_sharepoint(path)}")
        except Exception as exc:
            print(f"  UPLOAD FAILED: {exc}")
            print(f"  CSV is still saved locally at {path}")
            return 4

    return 0


# ===========================================================================
# SharePoint upload (only used when UPLOAD_TO_SHAREPOINT = True)
# ===========================================================================


def upload_to_sharepoint(local_path: str) -> str:
    """
    Upload via Microsoft Graph using the client credentials flow.
    Requires:  pip install msal

    The SPN needs Graph application permission Sites.Selected PLUS a
    site-level grant on the target site -- the API permission alone grants
    nothing, which shows up as a confusing 401/spException.
    """
    try:
        import msal
    except ImportError:
        raise RuntimeError("pip install msal")

    if not (SPN_TENANT_ID and SPN_CLIENT_ID and SPN_CLIENT_SECRET):
        raise RuntimeError("SPN_TENANT_ID / SPN_CLIENT_ID / "
                           "SPN_CLIENT_SECRET are not set")

    app = msal.ConfidentialClientApplication(
        client_id=SPN_CLIENT_ID,
        client_credential=SPN_CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{SPN_TENANT_ID}",
    )
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in result:
        raise RuntimeError(f"token request failed: {result.get('error')} "
                           f"{result.get('error_description', '')[:200]}")

    graph = "https://graph.microsoft.com/v1.0"
    session = requests.Session()
    session.headers["Authorization"] = f"Bearer {result['access_token']}"

    path = SITE_PATH if SITE_PATH.startswith("/") else "/" + SITE_PATH
    site = session.get(f"{graph}/sites/{SITE_HOSTNAME}:{path}",
                       timeout=REQUEST_TIMEOUT)
    if not site.ok:
        raise RuntimeError(
            f"site lookup failed ({site.status_code}). With Sites.Selected "
            f"this usually means the site-level grant is missing. "
            f"{site.text[:200]}")
    site_id = site.json()["id"]

    drives = session.get(f"{graph}/sites/{site_id}/drives",
                         timeout=REQUEST_TIMEOUT).json().get("value", [])
    drive = next((d for d in drives
                  if d.get("name", "").lower() == LIBRARY_NAME.lower()), None)
    if not drive:
        raise RuntimeError(
            f"no library named '{LIBRARY_NAME}'. Available: "
            f"{', '.join(d.get('name', '?') for d in drives)}")

    folder = TARGET_FOLDER.strip("/")
    filename = os.path.basename(local_path)
    destination = f"{folder}/{filename}" if folder else filename

    with open(local_path, "rb") as handle:
        response = session.put(
            f"{graph}/drives/{drive['id']}/root:/{destination}:/content"
            f"?@microsoft.graph.conflictBehavior=replace",
            data=handle, timeout=REQUEST_TIMEOUT)
    if not response.ok:
        raise RuntimeError(f"upload failed ({response.status_code}): "
                           f"{response.text[:200]}")
    return response.json().get("webUrl", "(uploaded)")


if __name__ == "__main__":
    sys.exit(main())
