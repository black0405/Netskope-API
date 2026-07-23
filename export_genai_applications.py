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
# NOTE: use the *datasearch* endpoint, not /events/data/. Per Netskope's
# docs, the datasearch endpoints are the ones built on the Skope IT Query
# Language and are what accept the `query` filter, `fields` column
# selection, and groupbys/orderbys aggregation. This is the API equivalent
# of Skope IT > Events & Alerts > Application Events.
ENDPOINT_PATH: str = "/api/v2/events/datasearch/application"

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

# The datasearch endpoint takes its own server-side query timeout, in
# seconds (Swagger marks it required, default 180). Raise it if large
# windows start returning timeouts.
QUERY_TIMEOUT: int = 180

# Print the exact request (URL + params) and a snippet of every raw response
# to the console. Turn this on FIRST if you're getting zero records -- it
# will show you immediately whether the problem is auth, the filter, or the
# response shape not matching what the script expects.
DEBUG: bool = False

# Name of the pagination query parameter. The datasearch endpoint documents
# "offset" (number of rows to skip before presenting results). The older
# /events/data/ endpoints used "skip" -- change this if you switch back.
OFFSET_PARAM_NAME: str = "offset"

# The mandatory filter.
# IMPORTANT: the field is `appcategory` (the CCI *application* category,
# e.g. 'Cloud Storage', 'Generative AI'), NOT `category` -- `category` is
# the web/URL category and is a different taxonomy. The tenant's Swagger
# page documents the filter example as:
#     query=appcategory eq 'Cloud Storage' and app eq 'Microsoft Office'
BASE_FILTER_FIELD: str = "appcategory"
BASE_FILTER_VALUE: str = "Generative AI"
BASE_FILTER_OPERATOR: str = "in"

# How many sample records to pull in preview mode (one API call, no time-
# window subdivision, so it comes back in a couple of seconds regardless of
# how big TIME_WINDOW_SECONDS/PAGE_SIZE are for the full run).
PREVIEW_LIMIT: int = 20

# Restrict and reorder the exported CSV to exactly these fields. Leave empty
# to keep the default behaviour (every discovered field becomes a column).
# A column listed here that never appears in the data is still exported,
# blank, as required.
#
# Mapped from the requested business columns to the REAL Netskope field
# names in this event schema. Three had no clean match -- see the comments;
# swap them if a preview run shows a better field for your tenant:
#   Object Type       -> "object_type"  (File / Folder / Message ...)
#   Object Name       -> "object"       (the named object, when the event
#                                        involves one -- blank otherwise)
DESIRED_COLUMNS: List[str] = [
    "user",                # User
    "app",                 # Application
    "url",                 # URL
    "activity",            # Activity
    "object_type",         # Object Type  (NOT `type` -- that's the event
                           #   type: nspolicy/connection/etc.)
    "object",              # Object Name  (the named file/folder/message)
    "timestamp",          # Event Date     (epoch; see RENAME note)
    "organization_unit",   # Organization Unit
    "file_size",           # Sum - File Size (MB)  (raw field is BYTES;
                           #   summed across the user+app pair, then
                           #   converted to MB on export)
    # Uncomment to also see whether each row was allowed/blocked and whether
    # it raised an alert -- useful for confirming a block-only dataset:
    # "action",
    # "alert",
]

# Columns in DESIRED_COLUMNS that do NOT exist in the API schema. They are
# still written to the CSV (blank) but are never sent in the `fields`
# parameter, since asking the API for an unknown field can fail the query.
PLACEHOLDER_COLUMNS: set = set()

# Ask the API to return ONLY the fields we actually export, via the
# datasearch `fields` parameter (e.g. fields=app,category). This cuts the
# response from ~100 fields per row down to the handful you need, which is
# the single biggest speed-up available for a full-day pull. Set to False
# to pull every field (useful when exploring the schema in preview mode).
SEND_FIELDS_PARAM: bool = True

# Collapse the export to unique combinations of these fields. With
# ["user", "app"] each user/application pair appears once, keeping the
# values from that pair's first event for the remaining columns. Set to an
# empty list to export every raw event row instead.
#
# NOTE: if you'd rather have a bare two-column list of distinct user/app
# pairs with no other columns, also trim DESIRED_COLUMNS down to
# ["user", "app"] -- the dedupe key and the exported columns are
# independent on purpose.
DEDUPE_ON: List[str] = ["user", "app"]

# Epoch fields (like `timestamp`) are converted to readable dates on export.
# Month/day/year: "%m/%d/%Y"           -> 07/23/2026
# Day/month/year: "%d/%m/%Y"           -> 23/07/2026
# Year first:     "%Y-%m-%d"           -> 2026-07-23
# With time:      "%m/%d/%Y %H:%M:%S"  -> 07/23/2026 14:05:11
# Times are rendered in the machine's LOCAL timezone, matching how the
# day boundaries in resolve_day_range are calculated.
DATE_FORMAT: str = "%m/%d/%Y"

# Fields holding epoch seconds that should be formatted with DATE_FORMAT.
EPOCH_FIELDS: set = {"timestamp", "_insertion_epoch_timestamp",
                     "_creation_timestamp", "src_time"}

# Numeric fields that should be ADDED UP across rows collapsed by DEDUPE_ON,
# rather than taking the first row's value. This is what makes the file size
# column a genuine "Sum" per user+application pair instead of a single
# event's size.
SUM_FIELDS: set = {"file_size"}

# Fields holding a byte count that should be converted to MB on export.
BYTES_FIELDS: set = {"file_size"}

# Decimal places for the MB conversion.
MB_DECIMALS: int = 2

# Prepend an unnamed leftmost column holding a row counter, so the first
# data row (spreadsheet row 2, since row 1 is the header) is numbered
# ROW_NUMBER_START. Set ROW_NUMBER_START = 2 if you want the values to
# match the spreadsheet's own row numbers instead of counting from 1.
ADD_ROW_NUMBER: bool = True
ROW_NUMBER_START: int = 1

# Rename the technical field names to the friendly headers you asked for.
# Only applied to columns actually present in DESIRED_COLUMNS above.
RENAME_COLUMNS: Dict[str, str] = {
    "user": "User",
    "app": "Application",
    "url": "URL",
    "activity": "Activity",
    "object_type": "Object Type",
    "object": "Object Name",
    "file_size": "Sum - File Size (MB)",
    "timestamp": "Event Date",
    "organization_unit": "Organization Unit",
    "action": "Action",
    "alert": "Alert",
}

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
        # Accepts either bare "a, b, c" or a pasted-in ['a', 'b'] / ["a","b"]
        # form -- strip any wrapping brackets and per-item quotes first so
        # the value never gets double-wrapped into something like
        # in ['[\'a\']'], which matches nothing.
        cleaned = raw.strip()
        if cleaned.startswith("[") and cleaned.endswith("]"):
            cleaned = cleaned[1:-1]
        parts = []
        for item in cleaned.split(","):
            item = item.strip().strip("'").strip('"').strip()
            if item:
                parts.append(quote_value(item))
        quoted = "[" + ", ".join(parts) + "]"
    else:
        quoted = quote_value(raw)

    return OPERATOR_MAP[key].format(f=field.strip(), v=quoted)


def build_query(extra_term: Optional[str]) -> str:
    """
    Build the full NQL query: the mandatory category filter, ANDed with the
    chosen region filter's already-built term, if one was chosen.
    """
    terms = [build_term(BASE_FILTER_FIELD, BASE_FILTER_OPERATOR, BASE_FILTER_VALUE)]
    if extra_term:
        terms.append(extra_term)
    return " and ".join(terms)


# ---------------------------------------------------------------------------
# Region filter presets
# ---------------------------------------------------------------------------
# The second filter is no longer free text -- it's always exactly one of
# these, ANDed with the base category filter, and nothing else. Each entry
# is (label, field, operator, values). "values" is a list even for a
# single-value filter; build_term joins them the same way "in" always has.
# To add/change a region, edit this list only.
REGION_FILTERS: List[Tuple[str, str, str, List[str]]] = [
    ("APAC IN", "usergroup", "in", [
        "AP1.OFC.LOC/India/Groups/IN_Global_netskope_AP1User",
        "india.asia.gcn.local/IN/Administrative/Groups - Security/IN_Global_netskope_Alluser",
        "ONE.OFC.LOC/Zone-ABI Global/Managed Objects/Managed Groups/SONEG-Netskope-ABI-GOA Team",
    ]),
    ("APAC VN", "organization_unit", "in", [
        "AP1.OFC.LOC/Vietnam/Users",
    ]),
    ("APAC KR", "usergroup", "in", [
        "ob.co.kr/OB/Groups/User/manual/S001",
    ]),
    ("APAC CN", "access_method", "in", [
        "Local Proxy",
    ]),
    ("EUR", "usergroup", "in", [
        "we.interbrew.net/ABI-IBS-EU-EuropeDC1 (Amsterdam)/Groups/User Groups/Applications/G-SWED-Netskope_Users",
        "rus.efesmoscow/GFolders/Netskope/G-SRUSG-NetskopeUsers_Exeptions-RUK",
        "rus.efesmoscow/GFolders/Netskope/G-SRUSG-NetskopeUsers-UA",
        "we.interbrew.net/ABI-IBS-EU-EuropeDC1 (Amsterdam)/Groups/User Groups/Applications/G-SWEG-DE-Netskope_Users",
        "we.interbrew.net/ABI-IBS-EU-EuropeDC1 (Amsterdam)/Groups/User Groups/Applications/SONED-ABI-EUR-Netskope-DigitalSolutions-Users",
        "we.interbrew.net/ABI-IBS-EU-EuropeDC1 (Amsterdam)/Groups/User Groups/Applications/G-SWEG-EUR-NPA-VPN",
        "rus.efesmoscow/GFolders/Netskope/G-SRUSG-NetskopeUsers-RU",
        "we.interbrew.net/ABI-IBS-EU-EuropeDC1 (Amsterdam)/Groups/User Groups/Applications/G-SWED-Netskope_UAT",
        "ccc.europe.gcn.local/Applications/Netscope/G-NPA-Users",
        "ccc.europe.gcn.local/Applications/Zscaler/G Users Proxy ZScaler",
        "ccc.europe.gcn.local/Applications/Zscaler/G Users Proxy HyperCare_Netskope",
        "we.interbrew.net/ABI-IBS-EU-EuropeDC1 (Amsterdam)/Groups/User Groups/Applications/G-SWEG-BE-Netskope_Users",
    ]),
    ("AFR", "usergroup", "in", [
        "beerdivision.africa.gcn.local/Groups/ZScaler Proxy/SG - AFR Proxy - Marketing Internet Access",
        "beerdivision.africa.gcn.local/Groups/ZScaler Proxy/SG - AFR Proxy - Standard Internet Access",
        "ONE.OFC.LOC/Zone-ABI Global/Managed Objects/Managed Groups/SONED-Workplace-MUstdinternetaccess",
        "beerdivision.africa.gcn.local/Groups/ZScaler Proxy/SG - AFR Proxy - Hypercare Admin Only",
    ]),
    ("GHQ", "usergroup", "equals", [
        "ONE.OFC.LOC/Zone-ABI Global/Managed Objects/Managed Groups/SONEG-NetskopePolicy-GHQ-AllUsers",
    ]),
    ("MAZ", "usergroup", "in", [
        "modelo.gmodelo.com.mx/Corporativo/Grupos/Diblo/GPO Groups/MAZ_Proxy_Mancom",
        "modelo.gmodelo.com.mx/Corporativo/Grupos/Diblo/GPO Groups/MAZ_Proxy_Advanced",
        "modelo.gmodelo.com.mx/Groups/Security/MAZ_Proxy_Service",
        "modelo.gmodelo.com.mx/Corporativo/Grupos/Diblo/GPO Groups/MAZ_Proxy_Marketing",
        "modelo.gmodelo.com.mx/Corporativo/Grupos/Diblo/GPO Groups/MAZ_Proxy_Basic",
    ]),
    ("NAZ", "usergroup", "in", [
        "abc.corp.anheuser-busch.com/Security Groups/Application Groups/NetSkope/APP ROLE-NAZ-Netskope_ABC",
        "na.interbrew.net/Security Groups/Application Groups/Netskope/APP ROLE-NAZ-Netskope_NA",
        "abpg.corp.anheuser-busch.com/Security Groups/Application Groups/Netskope/APP ROLE-NAZ-Netskope_ABPG",
        "ONE.OFC.LOC/Zone-North America/Security Groups/Application Groups/Netskope/APP ROLE-NAZ-Netskope_ONE",
    ]),
    ("SAZ", "usergroup", "in", [
        "cbn.com.bo/Users/SLASG_Netskope_AllUsers_BO",
        "cmqpnt.bue.bemberg.com.ar/Users/SLASG_Netskope_AllUsers_AR",
        "cervepar.local/Users/SLASG_Netskope_AllUsers_PY",
        "quinsacl.sgo.quilmes.cl/Users/SLASG_Netskope_AllUsers_CL",
        "quinsauy.mvd.bemberg.com.uy/Users/SLASG_Netskope_AllUsers_UY",
        "ONE.OFC.LOC/Zone-Latin America South/Managed Objects/Managed Groups/Netskope/SONEG_Netskope_AllUsers_LAS",
        "la.interbrew.net/Corporativo/ADM_CENTRAL_AC/Service_Groups/SAZ_LAN_Netskope_Advanced",
        "la.interbrew.net/Corporativo/ADM_CENTRAL_AC/Service_Groups/SAZ_LAN_Netskope_Intermediate",
    ]),
]


def prompt_for_region_filter() -> Optional[str]:
    """
    Menu of pre-approved regional filters. This is the ONLY second filter --
    there is no free-text field/operator/value entry anymore. Returns the
    fully-built NQL term for the chosen region, or None for category-only.
    """
    print("\nSecond filter -- pick a region (ANDed with category = GenerativeAI).")
    for i, (label, _field, _operator, _values) in enumerate(REGION_FILTERS, 1):
        print(f"  {i}. {label}")
    print("  0. None (category filter only)")

    by_label = {entry[0].lower(): entry for entry in REGION_FILTERS}

    while True:
        choice = input("Choice: ").strip()
        if choice in ("", "0"):
            print("  No region filter -- pulling all GenerativeAI events.")
            return None

        if choice.isdigit() and 1 <= int(choice) <= len(REGION_FILTERS):
            label, field, operator, values = REGION_FILTERS[int(choice) - 1]
        elif choice.lower() in by_label:
            label, field, operator, values = by_label[choice.lower()]
        else:
            print(f"  '{choice}' isn't valid. Enter a number 1-{len(REGION_FILTERS)}, "
                  f"a region name, or 0 for none.")
            continue

        print(f"  Using: {label}")
        return build_term(field, operator, ", ".join(values))


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

def fields_param() -> Optional[str]:
    """
    Comma-separated field list for the datasearch `fields` parameter, or
    None to let the API return everything. Placeholder columns that don't
    exist in the schema are excluded so they can't fail the query.
    """
    if not SEND_FIELDS_PARAM or not DESIRED_COLUMNS:
        return None
    real = [c for c in DESIRED_COLUMNS if c not in PLACEHOLDER_COLUMNS]
    return ",".join(real) if real else None


def render_progress(stats: Dict[str, int], bar_width: int = 28) -> None:
    """
    Draw a single-line, in-place progress bar (overwrites itself with \r
    rather than scrolling the console with one line per page/window).
    """
    total = stats.get("total_windows", 1) or 1
    done = stats.get("windows_done", 0)
    frac = min(done / total, 1.0)
    filled = int(bar_width * frac)
    bar = "#" * filled + "-" * (bar_width - filled)
    print(
        f"\r  [{bar}] {int(frac * 100):3d}%  "
        f"{stats.get('records', 0):>7} records  "
        f"{stats.get('calls', 0):>4} calls",
        end="",
        flush=True,
    )


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
                "timeout": QUERY_TIMEOUT,
            }
            selected = fields_param()
            if selected:
                params["fields"] = selected

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
        stats["records"] = stats.get("records", 0) + len(page)
        render_progress(stats)

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

    stats["total_windows"] = len(windows)
    stats["windows_done"] = 0
    stats.setdefault("records", 0)

    seen: set = set()
    all_records: List[Dict[str, Any]] = []

    for index, (w_start, w_end) in enumerate(windows, 1):
        for record in fetch_window(session, query, w_start, w_end, stats):
            # _id is Netskope's per-event unique key. Fall back to the whole
            # record if it's absent so we never drop legitimate rows.
            key = record.get("_id") or json.dumps(record, sort_keys=True)
            if key in seen:
                continue
            seen.add(key)
            all_records.append(record)
        stats["windows_done"] = index
        render_progress(stats)

    print()  # move off the progress bar line once the pull is complete
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

def dedupe_key(row: Dict[str, Any], fields: List[str]) -> Tuple[str, ...]:
    """
    Build a normalized comparison key so trivially-different spellings of
    the same user/app collapse together. Real Netskope data routinely mixes
    'Alice@ABI.com' with 'alice@abi.com' and 'ChatGPT ' with 'ChatGPT', and
    a raw string compare would treat those as distinct rows.

    Normalization: casefold, strip, and collapse internal whitespace runs.
    """
    key: List[str] = []
    for field in fields:
        value = row.get(field, "")
        text = "" if value is None else str(value)
        key.append(" ".join(text.split()).casefold())
    return tuple(key)


def to_number(value: Any) -> float:
    """Best-effort numeric coercion; anything unparseable counts as 0."""
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def dedupe_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Collapse rows to unique combinations of DEDUPE_ON, keeping the first
    occurrence of each. Order is preserved. Rows missing a key field are
    treated as having a blank value for it rather than being dropped.

    Fields listed in SUM_FIELDS are accumulated across every row that
    collapses into the same key, so the kept row carries the total for that
    user/application pair rather than just its first event's value.
    """
    if not DEDUPE_ON:
        return rows

    index: Dict[Tuple[str, ...], Dict[str, Any]] = {}
    unique: List[Dict[str, Any]] = []
    for row in rows:
        key = dedupe_key(row, DEDUPE_ON)
        kept = index.get(key)

        if kept is None:
            row = dict(row)   # don't mutate the caller's data
            for field in SUM_FIELDS:
                if field in row:
                    row[field] = to_number(row.get(field))
            index[key] = row
            unique.append(row)
            continue

        # Duplicate: fold its summable values into the row we're keeping.
        for field in SUM_FIELDS:
            if field in row or field in kept:
                kept[field] = to_number(kept.get(field)) + to_number(row.get(field))

    return unique


def format_epoch(value: Any) -> Any:
    """
    Convert an epoch-seconds value to a readable date using DATE_FORMAT.

    Returns the original value untouched if it isn't a usable epoch --
    blanks stay blank, and any field that already holds a formatted string
    (some Netskope time fields do) passes through rather than being
    mangled into a wrong date.
    """
    if value is None or value == "":
        return ""
    try:
        epoch = float(value)
    except (TypeError, ValueError):
        return value          # already a string date, or something else
    if epoch <= 0:
        return ""
    try:
        return time.strftime(DATE_FORMAT, time.localtime(epoch))
    except (ValueError, OSError, OverflowError):
        return value          # out of range for the platform's localtime


def bytes_to_mb(value: Any) -> Any:
    """Convert a byte count to megabytes, rounded to MB_DECIMALS."""
    if value is None or value == "":
        return ""
    try:
        size = float(value)
    except (TypeError, ValueError):
        return value
    if size <= 0:
        return 0
    return round(size / (1024 * 1024), MB_DECIMALS)


def export_csv(rows: List[Dict[str, Any]], columns: List[str], path: str) -> None:
    """
    Write to CSV with every discovered field as a column. Missing values are
    written blank. QUOTE_MINIMAL keeps embedded commas safe. RENAME_COLUMNS,
    if set, relabels headers for columns that are actually present.

    Dedupe is applied HERE rather than at the call sites, so that every
    export path (full run, preview, anything added later) is guaranteed to
    get it. dedupe_rows is idempotent, so calling it twice is harmless.
    """
    rows = dedupe_rows(rows)
    frame = pd.DataFrame(rows, columns=columns)
    # Trim stray leading/trailing whitespace so the CSV doesn't carry the
    # padding that some source fields arrive with.
    for column in frame.columns:
        if frame[column].dtype == object:
            frame[column] = frame[column].apply(
                lambda v: v.strip() if isinstance(v, str) else v
            )
    frame = frame.where(pd.notnull(frame), "")
    # Convert epoch fields to readable dates before headers get renamed.
    for column in frame.columns:
        if column in EPOCH_FIELDS:
            frame[column] = frame[column].apply(format_epoch)
        elif column in BYTES_FIELDS:
            frame[column] = frame[column].apply(bytes_to_mb)

    if RENAME_COLUMNS:
        frame = frame.rename(columns=RENAME_COLUMNS)

    # Unnamed leftmost counter column. Inserted last so it isn't affected by
    # the rename map, and given an empty header name so the CSV's first
    # column heading is blank.
    if ADD_ROW_NUMBER:
        frame.insert(
            0, "",
            range(ROW_NUMBER_START, ROW_NUMBER_START + len(frame)),
        )

    frame.to_csv(path, index=False, quoting=csv.QUOTE_MINIMAL, encoding="utf-8")


# ---------------------------------------------------------------------------
# Zero-result diagnostics
# ---------------------------------------------------------------------------

def diagnose_zero_results(
    session: requests.Session,
    start: int,
    end: int,
    stats: Dict[str, int],
) -> None:
    """
    Runs automatically when the filtered pull comes back empty. Makes one
    lightweight UNFILTERED call over the same day to answer two questions
    without another round of manual DEBUG digging:

      1. Is there any data at all for this day/tenant/token? If not, the
         problem is auth, date range, or tenant-side -- not the filter.
      2. What do the real field names and category-ish values look like?
         A wrong filter field name or a category label that doesn't match
         what's actually stored is the most common reason for 0 rows.
    """
    url = BASE_URL.rstrip("/") + ENDPOINT_PATH
    params = {"starttime": start, "endtime": end, "limit": 5, OFFSET_PARAM_NAME: 0}

    print("\nFilter matched nothing -- running one unfiltered check on the "
          "same day to see what's actually there...")
    try:
        payload = request_page(session, url, params, stats)
    except ApiError as exc:
        print(f"  Unfiltered check also failed: {exc}")
        return

    sample = extract_records(payload)
    if not sample:
        print("  No records at all for this day, even without a filter. "
              "That points to auth, the date range, or tenant retention -- "
              "not the filter text. Try DEBUG = True and a wider date range.")
        return

    flat_sample = [flatten_record(r) for r in sample]
    cols = discover_columns(flat_sample)
    shown = cols[:40]
    print(f"  Data exists: found {len(sample)} unfiltered sample record(s).")
    print(f"  Available fields: {', '.join(shown)}"
          f"{' ...' if len(cols) > len(shown) else ''}")

    # Surface anything that looks like it could be the field(s) being
    # filtered on, along with real sample values, so a wrong field name or
    # wrong expected value is obvious at a glance.
    candidates = [c for c in cols
                  if any(term in c.lower() for term in ("categ", "org", "unit", "ou"))]
    for field in candidates:
        values = sorted({str(row[field]) for row in flat_sample
                          if row.get(field) not in (None, "")})
        if values:
            print(f"    {field} sample value(s): {', '.join(values[:5])}")


# ---------------------------------------------------------------------------
# Preview mode
# ---------------------------------------------------------------------------

def run_preview(
    session: requests.Session,
    query: str,
    start: int,
    end: int,
    stats: Dict[str, int],
) -> None:
    """
    One fast API call -- no time-window subdivision, small limit -- so you
    can check the query and field names in seconds instead of waiting
    through a full day's worth of windowed pagination. Writes a small
    preview_ CSV too, in case the console truncates long field lists.
    """
    url = BASE_URL.rstrip("/") + ENDPOINT_PATH
    # Deliberately NOT sending `fields` here -- preview should show the full
    # schema so you can confirm real field names and value distributions.
    params = {
        "query": query,
        "starttime": start,
        "endtime": end,
        "limit": PREVIEW_LIMIT,
        OFFSET_PARAM_NAME: 0,
        "timeout": QUERY_TIMEOUT,
    }
    print(f"\nPreview: single call, limit={PREVIEW_LIMIT}, no windowing...")
    payload = request_page(session, url, params, stats)
    records = extract_records(payload)

    if not records:
        print("  No records for this filter/day. Run the full pull instead -- "
              "it automatically runs an unfiltered diagnostic on zero results.")
        return

    flat_rows = [flatten_record(r) for r in records]
    columns = discover_columns(flat_rows)
    print(f"  Got {len(records)} sample record(s).")
    print(f"  Fields found: {', '.join(columns)}")

    # Value distribution for the fields that explain "why am I only seeing X".
    # Nothing in the query filters on these -- so whatever shows up here is
    # simply what the endpoint returns for the category filter. If action is
    # 100% block, that's the tenant's policy/feed, not this script.
    for field in ("action", "alert", "activity", "traffic_type", "type"):
        if field not in columns:
            continue
        counts: Dict[str, int] = {}
        for row in flat_rows:
            value = str(row.get(field, "")).strip() or "(blank)"
            counts[value] = counts.get(value, 0) + 1
        breakdown = ", ".join(f"{v}={n}" for v, n in
                              sorted(counts.items(), key=lambda kv: -kv[1]))
        print(f"    {field}: {breakdown}")

    preview_path = "preview_" + OUTPUT_FILE
    export_csv(flat_rows, DESIRED_COLUMNS or columns, preview_path)
    print(f"  Sample written to: {preview_path}")
    print("  Once the field names/values look right, rerun and choose the "
          "full day pull.")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    stats: Dict[str, int] = {"calls": 0}

    if "PUT_YOUR_API_TOKEN" in API_TOKEN or "YOUR_TENANT" in BASE_URL:
        print("Set API_TOKEN and BASE_URL at the top of the script first.")
        return

    # Second filter (region presets) disabled for now -- category-only pull.
    # Re-enable by uncommenting the line below.
    # extra_term = prompt_for_region_filter()
    extra_term = None

    try:
        query = build_query(extra_term)
    except ValueError as exc:
        print(f"Filter error: {exc}")
        return

    try:
        start, end, label = prompt_for_date_range()
    except ValueError as exc:
        print(f"Date error: {exc}")
        return

    output_path = resolve_output_path(label)

    print(f"\n{label} ({DAY_MODE})  |  {query}")
    print(f"{BASE_URL.rstrip('/')}{ENDPOINT_PATH}")

    session = build_session()

    mode = input("\nRun mode -- [P]review (fast, "
                 f"{PREVIEW_LIMIT} records) or [F]ull day? [F]: ").strip().lower()
    if mode.startswith("p"):
        try:
            run_preview(session, query, start, end, stats)
        except ApiError as exc:
            print(f"\nFAILED: {exc}")
        return

    print("Fetching...")
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
        diagnose_zero_results(session, start, end, stats)
        return

    flat_rows = [flatten_record(r) for r in records]
    discovered = discover_columns(flat_rows)
    raw_row_count = len(flat_rows)
    flat_rows = dedupe_rows(flat_rows)
    if DESIRED_COLUMNS:
        missing = [c for c in DESIRED_COLUMNS if c not in discovered]
        if missing:
            print(f"  [WARN] Requested columns not found in the data "
                  f"(will export blank): {', '.join(missing)}")
        columns = DESIRED_COLUMNS
    else:
        columns = discovered
    export_csv(flat_rows, columns, output_path)

    print("\n" + "-" * 52)
    print(f"Day exported            : {label}")
    print(f"Total records retrieved : {raw_row_count}")
    if DEDUPE_ON:
        print(f"Unique rows exported    : {len(flat_rows)} "
              f"(by {'+'.join(DEDUPE_ON)})")
    print(f"Unique columns found    : {len(discovered)}")
    print(f"Columns exported        : {len(columns)}")
    print(f"API calls made          : {stats['calls']}")
    print(f"Exported to             : {output_path}")
    print("-" * 52)


if __name__ == "__main__":
    main()
