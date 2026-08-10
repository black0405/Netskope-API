#!/usr/bin/env python3
"""
Export Netskope GenAI traffic volume (page events) to CSV.

Endpoint : GET /api/v2/events/datasearch/page
Docs     : https://docs.netskope.com/en/using-the-rest-ap-v2-datasearch-endpoint

Columns exported:
    Application, User, URL, Event Date, Organization Unit,
    Bytes Upload, Bytes Download, Total Bytes

Page events are the ONLY place the traffic byte counts live. Per Netskope's
docs, "Page Events provide information related to the amount of bytes
transferred for a connection" -- application events carry activity and
object detail but no byte fields. The companion script
export_genai_applications.py covers that half.

WHY datasearch AND NOT /events/data/:
  - /api/v2/events/data/*  is the legacy path. Tenants have begun returning
    403 "Legacy Rest API v2 is deprecated ... use the Data Export API", and
    its v1 predecessor was retired on 2024-10-15.
  - /api/v2/events/dataexport/* is an index/iterator feed: forward streaming
    only, in one-hour batches, meant for continuous SIEM ingestion. It is a
    poor fit for "give me one specific past day".
  - /api/v2/events/datasearch/* is the Skope IT Query Language endpoint
    built for ad-hoc queries over an arbitrary time range, and it is the
    only one of the three that accepts `fields`, `query` and server-side
    aggregation. That is exactly this job.

Rows are fetched with offset paging inside time windows and written out
ONE ROW PER EVENT: no grouping, no summing, and byte counts left in raw
bytes exactly as the API returned them. Duplicate events are still dropped
by their `_id`, which removes API paging artefacts only -- it never merges
two genuinely distinct events.

Expect this file to be considerably larger than an aggregated one: it is
one line per page event for the day, not one line per user.

Dependencies: requests, pandas (stdlib: typing, time, csv, json)
"""

import csv
import json
import logging
import os
import shutil
import subprocess
import sys
import time
import zipfile
from logging.handlers import TimedRotatingFileHandler
from typing import Any, Dict, Iterable, List, Optional, Tuple

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
# Everything is configured in this file. There is no config.env, no .env, and
# nothing is read from environment variables -- edit the constants below and
# that is the whole configuration story.
#
# The settings are grouped into two sections further down:
#   LOGGING        -- log folder, level, retention
#   CONFIGURATION  -- tenant, credentials, day to pull, output, SharePoint
#
# SECURITY NOTE: the API token and the SharePoint client secret now live in
# this .py. Anyone who can read the script can read the credentials, and they
# will follow the file into any copy, backup, or repository it lands in.
# Keep the file permissions tight (chmod 600 on Linux) and do not commit it.
# ---------------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def anchored_path(value: str) -> str:
    """
    Resolve a configured folder to an absolute path.

    Relative paths are anchored to the SCRIPT's folder, not the current
    working directory. This matters because cron, systemd and Task
    Scheduler all start the process somewhere unexpected (often the
    filesystem root or the Windows system folder) -- an unanchored
    "./output" would silently write to the wrong place. Absolute paths in
    the config are used as-is.
    """
    if not value:
        return SCRIPT_DIR
    if os.path.isabs(value):
        return value
    return os.path.normpath(os.path.join(SCRIPT_DIR, value))


# ---------------------------------------------------------------------------
# LOGGING
# ---------------------------------------------------------------------------
# One log file per calendar day (netskope_traffic_export.log, rotated at
# midnight into ...log.YYYY-MM-DD). Because the file is per-day and
# opened in append mode, running the job twice a day -- or ten times -- puts
# every run in that day's file, each clearly delimited by a RUN START /
# RUN END banner with its own run id.
#
# Retention: live logs older than LOG_RETENTION_DAYS are swept into a
# monthly zip under logs/archive/, then deleted. Archives older than
# ARCHIVE_RETENTION_DAYS are removed (0 = keep forever).
# ---------------------------------------------------------------------------

LOG_DIR: str = anchored_path("./logs")
LOG_LEVEL: str = "INFO"                    # DEBUG | INFO | WARNING | ERROR
LOG_RETENTION_DAYS: int = 20
ARCHIVE_RETENTION_DAYS: int = 180          # 0 = keep archives forever
# Distinct from the companion script's "netskope_export.log" so the two runs
# don't interleave in one file. They can share LOG_DIR safely.
LOG_BASENAME: str = "netskope_traffic_export.log"

log = logging.getLogger("netskope_export")


class RunContext(logging.Filter):
    """Stamps every record with the current run id."""

    def __init__(self, run_id: str) -> None:
        super().__init__()
        self.run_id = run_id

    def filter(self, record: logging.LogRecord) -> bool:
        record.run_id = self.run_id
        return True


def setup_logging(run_id: str = "-") -> str:
    """
    Configure file + console logging. Returns the active log file path.

    The file handler rotates at midnight so each day is its own file; the
    console handler keeps interactive runs readable. Safe to call twice --
    existing handlers are cleared first.
    """
    os.makedirs(LOG_DIR, exist_ok=True)
    log_path = os.path.join(LOG_DIR, LOG_BASENAME)

    # .upper() matters: getattr(logging, "info") returns the log FUNCTION, not
    # a level int, and setLevel would then fail on a lowercase setting.
    log.setLevel(getattr(logging, LOG_LEVEL.upper(), logging.INFO))
    log.handlers.clear()
    # Filters too, not just handlers -- otherwise a second call stacks another
    # RunContext on the logger and every record is stamped twice over.
    log.filters.clear()
    log.propagate = False

    # Every line carries the run id, so two runs in one day stay readable.
    file_format = logging.Formatter(
        "%(asctime)s  %(levelname)-8s [%(run_id)s]  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    file_handler = TimedRotatingFileHandler(
        log_path, when="midnight", interval=1,
        backupCount=LOG_RETENTION_DAYS + 5,   # archival does the real pruning
        encoding="utf-8",
    )
    file_handler.suffix = "%Y-%m-%d"
    file_handler.setFormatter(file_format)
    log.addHandler(file_handler)

    # Console: no timestamps, they're noise when you're watching it live.
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(logging.Formatter("%(message)s"))
    log.addHandler(console)

    # Attach on the LOGGER, not the handlers: this guarantees run_id exists
    # on every record, including ones emitted by helpers outside main().
    # Without it the formatter raises KeyError: 'run_id'.
    log.addFilter(RunContext(run_id))

    return log_path


def archive_old_logs() -> None:
    """
    Move rotated logs older than LOG_RETENTION_DAYS into a monthly zip in
    logs/archive/, then delete the originals. Old zips are pruned per
    ARCHIVE_RETENTION_DAYS.

    Failures here are logged but never fatal -- housekeeping must not stop
    the export.
    """
    try:
        archive_dir = os.path.join(LOG_DIR, "archive")
        cutoff = time.time() - (LOG_RETENTION_DAYS * ONE_DAY)

        stale: List[str] = []
        for name in os.listdir(LOG_DIR):
            path = os.path.join(LOG_DIR, name)
            # Only rotated files (basename.YYYY-MM-DD), never the live one.
            if not os.path.isfile(path) or not name.startswith(LOG_BASENAME):
                continue
            if name == LOG_BASENAME:
                continue
            if os.path.getmtime(path) < cutoff:
                stale.append(path)

        if stale:
            os.makedirs(archive_dir, exist_ok=True)
            # Group into one zip per month so the folder stays tidy.
            groups: Dict[str, List[str]] = {}
            for path in stale:
                month = time.strftime("%Y-%m", time.localtime(os.path.getmtime(path)))
                groups.setdefault(month, []).append(path)

            for month, paths in groups.items():
                zip_path = os.path.join(archive_dir, f"logs_{month}.zip")
                with zipfile.ZipFile(zip_path, "a", zipfile.ZIP_DEFLATED) as bundle:
                    existing = set(bundle.namelist())
                    for path in paths:
                        name = os.path.basename(path)
                        if name not in existing:
                            bundle.write(path, name)
                for path in paths:
                    os.remove(path)
                log.info("Archived %d log file(s) into %s",
                         len(paths), os.path.basename(zip_path))

        # Prune the archives themselves.
        if ARCHIVE_RETENTION_DAYS > 0 and os.path.isdir(archive_dir):
            zip_cutoff = time.time() - (ARCHIVE_RETENTION_DAYS * ONE_DAY)
            for name in os.listdir(archive_dir):
                path = os.path.join(archive_dir, name)
                if name.endswith(".zip") and os.path.getmtime(path) < zip_cutoff:
                    os.remove(path)
                    log.info("Removed expired archive %s", name)

    except Exception as exc:                      # housekeeping only
        log.warning("Log archival skipped: %s", exc)


# ---------------------------------------------------------------------------
# CONFIGURATION  -- edit the values here; this is the only place settings live
# ---------------------------------------------------------------------------

# Your API token. Netskope issues these under Settings > Tools > REST API v2.
API_TOKEN: str = "PUT_YOUR_API_TOKEN_HERE"

# Tenant base URL, no trailing slash. e.g. "https://myorg.goskope.com"
BASE_URL: str = "https://YOUR_TENANT.goskope.com"

# Endpoint path. Use the *datasearch* endpoints -- they are the ones built
# on the Skope IT Query Language and accept the `query` filter and `fields`
# column selection.
#
# WHICH EVENT TYPE:
#   page         Skope IT > Page Events. Carries the traffic byte counts
#                (Total / Uploaded / Downloaded). This is the one we want.
#   application  Skope IT > Application Events. Activity detail (Login,
#                Upload, Share, object type / object name) and file_size,
#                but NO byte fields.
# Per Netskope's docs: "you can view activities of an app in application
# events and view bytes information in page events."
#
# NOT RELEVANT HERE, but worth knowing it exists: datasearch can aggregate
# server-side. Netskope's documented example is
#
#   /api/v2/events/datasearch/page?starttime=0&endtime=0&groupby=user
#     &fields=server_bytes=sum(server_bytes_total),
#             client_bytes=sum(client_bytes_total),
#             numbytes=sum(numbytes_total)
#
# which returns one pre-aggregated row per user in a single call. This
# script wants the opposite -- untouched per-event rows -- so it pages the
# raw events instead. Keep that URL in mind only if you later decide you
# want totals rather than detail.
ENDPOINT_PATH: str = "/api/v2/events/datasearch/page"

# Authentication style.
#   "bearer"   -> Authorization: Bearer <token>        (requested / generic)
#   "netskope" -> Netskope-Api-Token: <token>          (what Netskope expects)
# If you get 401s with "bearer", flip this to "netskope".
AUTH_MODE: str = "bearer"

# Output file. If DATE_STAMP_FILENAME is True the day being exported is
# appended, e.g. GenerativeAI_Traffic_2026-07-19.csv -- useful when a
# folder of dated files is being picked up downstream.
# Change the extension to .xlsx if you ever want Excel output instead; the
# writer switches on the extension (that path needs the openpyxl package).
OUTPUT_FILE: str = "GenerativeAI_Traffic.csv"

# Folder the CSVs are written into. Created if missing. A relative path is
# anchored to this script's folder, never the current working directory.
OUTPUT_DIR: str = anchored_path("./output")
DATE_STAMP_FILENAME: bool = True

# Worksheet name used only when OUTPUT_FILE has a .xlsx extension.
# Excel limits: max 31 chars, and none of  : \ / ? * [ ]
EXCEL_SHEET_NAME: str = "in"

# --- Which single day to pull ---------------------------------------------
# DAY_MODE:
#   "yesterday" -> the full previous calendar day, 00:00:00 to 23:59:59 local
#   "today"     -> midnight local through right now
#   "date"      -> the specific calendar day named in TARGET_DATE
#   "rolling"   -> the trailing 24 hours from this moment
DAY_MODE: str = "yesterday"
TARGET_DATE: str = "2026-07-19"     # only used when DAY_MODE = "date"

# --- Unattended / scheduled runs -----------------------------------------
# INTERACTIVE False skips the "which day?" prompt entirely and just uses
# DAY_MODE, so the script can run from Task Scheduler / cron with no one at
# the keyboard. Set True if you'd rather be asked each time.
INTERACTIVE: bool = False

# A date can also be passed on the command line, which overrides DAY_MODE:
#     python export_genai_applications.py              -> yesterday
#     python export_genai_applications.py 2026-07-19   -> that specific day
#
# Skip the run (exit 0) if the output file for that day already exists.
# Keeps a re-run or an overlapping schedule from redoing finished work.
# False while testing, so re-running the same day actually re-runs it.
SKIP_IF_EXISTS: bool = False

# --- SharePoint upload ----------------------------------------------------
# False writes the CSV locally and skips the upload entirely -- no Graph
# credentials needed. Set True once you're ready to test the upload leg.
SHAREPOINT_ENABLED: bool = False

# Entra ID app registration (the SPN you had created).
# These three are static values from the SPN your admin created.
#
# Not to be confused with the Graph ACCESS TOKEN, which the script fetches
# at runtime by exchanging these three -- that one is short-lived and is
# never configured anywhere.
#
# SECURITY: CLIENT_SECRET is a live credential sitting in a source file.
# See the note at the top of the script.
TENANT_ID: str = "PUT_TENANT_ID_HERE"
CLIENT_ID: str = "PUT_CLIENT_ID_HERE"
CLIENT_SECRET: str = "PUT_CLIENT_SECRET_HERE"

# Where to put the file. Two ways to identify the site:
#   A) leave SITE_ID blank and fill in the hostname + path, and the script
#      resolves the ID for you (easier, costs one extra API call)
#   B) paste a known SITE_ID and the hostname/path are ignored
SITE_HOSTNAME: str = "yourtenant.sharepoint.com"
SITE_PATH: str = "/sites/YourSiteName"
SITE_ID: str = ""

# Document library ("drive"). Leave DRIVE_ID blank to look the library up by
# name; "Documents" is the default library on most sites (it shows in the UI
# as "Shared Documents").
LIBRARY_NAME: str = "Documents"
DRIVE_ID: str = ""

# Folder inside the library. "" uploads to the library root. Nested paths
# are fine ("Reports/Netskope/GenAI"); missing folders are created.
TARGET_FOLDER: str = "Netskope/GenAI"

# Overwrite a file of the same name if it's already up there.
SHAREPOINT_REPLACE_EXISTING: bool = True

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

# Server-side query timeout in seconds, used only when SEND_QUERY_TIMEOUT is
# True (i.e. on the datasearch endpoints, which document it as required with
# a default of 180). Ignored on /events/data/.
QUERY_TIMEOUT: int = 180

# Print the exact request (URL + params) and a snippet of every raw response
# to the console. Turn this on FIRST if you're getting zero records -- it
# will show you immediately whether the problem is auth, the filter, or the
# response shape not matching what the script expects.
DEBUG: bool = False

# Name of the pagination query parameter -- the number of rows to skip
# before returning results. Both v2 event endpoints document "offset"; the
# older v1 API called it "skip".
#
# TROUBLESHOOTING: if paging misbehaves on this endpoint (page 2 repeats
# page 1, or the run never terminates), the parameter name is being ignored.
# Try "skip" here before assuming anything else is wrong.
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

# Restrict and reorder the exported CSV to exactly these fields. Leave empty
# to keep the default behaviour (every discovered field becomes a column).
# A column listed here that never appears in the data is still exported,
# blank, as required.
#
# The order here IS the column order in the CSV (after the S.No. counter).
# RENAME_COLUMNS below turns each into its friendly header.
DESIRED_COLUMNS: List[str] = [
    "app",                 # Application
    "user",                # User
    "url",                 # URL
    "timestamp",           # Event Date  (epoch -> DATE_FORMAT)
    "organization_unit",   # Organization Unit
    "client_bytes",        # Bytes Upload
    "server_bytes",        # Bytes Download
    "numbytes",            # Total Bytes
]

# Alternative names for the same measure, tried in order. The first one
# present in the returned data wins, and its values are copied onto the
# canonical name used in DESIRED_COLUMNS. This keeps the export working
# across the naming differences between Netskope schemas.
#
# WHICH WAY ROUND THE BYTE FIELDS GO -- per Netskope's field reference:
#   client_bytes  "bytes received from end-user web-client"  -> UPLOAD
#   server_bytes  "bytes received from destination site"     -> DOWNLOAD
#   numbytes      total bytes transferred for the connection
#
# The "_total" variants are the names Netskope's own datasearch example
# uses inside its sum() aggregates, so they are listed first as the most
# likely schema; ingress_* is what some tenants expose instead.
FIELD_ALIASES: Dict[str, List[str]] = {
    "numbytes":     ["numbytes_total", "numbytes", "total_bytes", "bytes"],
    "client_bytes": ["client_bytes_total", "ingress_client_bytes",
                     "client_bytes", "bytes_uploaded", "upload_bytes"],
    "server_bytes": ["server_bytes_total", "ingress_server_bytes",
                     "server_bytes", "bytes_downloaded", "download_bytes"],
}

# Columns in DESIRED_COLUMNS that do NOT exist in the API schema. They are
# still written to the CSV (blank) but are never sent in the `fields`
# parameter, since asking the API for an unknown field can fail the query.
PLACEHOLDER_COLUMNS: set = set()

# Ask the API to return ONLY the fields we actually export, via the `fields`
# parameter (e.g. fields=app,user). This cuts the response from ~100 fields
# per row down to the handful you need. Documented on datasearch as "names
# of the columns whose values want to be retrieved in results".
#
# Because the byte field names vary by schema the request asks for every
# alias in FIELD_ALIASES. If your tenant errors on an unknown name, set this
# to False and the script will pull everything and pick columns locally.
SEND_FIELDS_PARAM: bool = True

# Send the server-side `timeout` query parameter. datasearch documents it
# with a default of 180s, so it is on here.
SEND_QUERY_TIMEOUT: bool = True

# Collapse the export to unique combinations of these fields.
#
# EMPTY = no collapsing. One row per raw event, exactly as Netskope returned
# it: a user who hit twenty URLs on one app gets twenty rows, each with that
# single event's own byte counts. Nothing is merged and nothing is summed,
# so every value in the CSV is traceable back to one event.
#
# Set to ["user", "app"] to get one row per user/application pair instead --
# but only together with SUM_FIELDS below, or the byte columns would show an
# arbitrary single event's figures rather than the pair's total.
DEDUPE_ON: List[str] = []

# Epoch fields (like `timestamp`) are converted to readable dates on export.
# "ctime"                              -> Tue Dec  9 21:00:20 2025
# Month/day/year: "%m/%d/%Y"           -> 07/23/2026
# Day/month/year: "%d/%m/%Y"           -> 23/07/2026
# Year first:     "%Y-%m-%d"           -> 2026-07-23
# With time:      "%m/%d/%Y %H:%M:%S"  -> 07/23/2026 14:05:11
# Times are rendered in the machine's LOCAL timezone, matching how the
# day boundaries in resolve_day_range are calculated.
#
# "ctime" is a special value, not a strftime pattern. It produces the
# day-space-padded form above via time.ctime(). Writing it as a pattern
# would need %e for the space-padded day, and %e does not exist on Windows
# -- so the sentinel is the portable way to get this exact layout.
DATE_FORMAT: str = "ctime"

# Fields holding epoch seconds that should be formatted with DATE_FORMAT.
EPOCH_FIELDS: set = {"timestamp", "_insertion_epoch_timestamp",
                     "_creation_timestamp", "src_time"}

# Numeric fields ADDED UP across rows collapsed by DEDUPE_ON.
#
# EMPTY = nothing is summed. Each row carries its own event's byte counts.
# (With DEDUPE_ON empty there is nothing to sum across anyway, so these two
# settings belong together -- change one and you almost certainly want to
# change the other.)
SUM_FIELDS: set = set()

# Fields holding a byte count that should be converted to MB on export.
#
# EMPTY = values are exported as RAW BYTES, exactly as the API returned
# them. This is why the headers read "Bytes Upload" / "Bytes Download" /
# "Total Bytes" with no unit suffix. To switch to megabytes, put the three
# field names back in here and rename the headers accordingly.
BYTES_FIELDS: set = set()

# Decimal places for the MB conversion. Unused while BYTES_FIELDS is empty.
MB_DECIMALS: int = 2

# Fields written as WHOLE NUMBERS.
#
# This is not cosmetic. pandas types a column float64 the moment one row is
# missing a value, so a single event without a byte count turns every figure
# in that column into "54665.0" instead of "54665". Forcing them back to
# integers on the way out keeps raw byte counts looking like byte counts.
#
# Only applies to fields NOT in BYTES_FIELDS -- an MB value needs decimals.
INT_FIELDS: set = {"numbytes", "server_bytes", "client_bytes"}

# Drop any row where EVERY field listed here is zero (or blank/absent).
#
# A page event with no bytes in either direction carries nothing for this
# report -- typically a connection that was blocked, or one that resolved
# without transferring anything. Listing the two directional fields rather
# than numbytes is deliberate: it is the upload and download legs that
# decide whether real traffic happened.
#
# Set to [] to keep every row including the all-zero ones.
DROP_IF_ALL_ZERO: List[str] = []

# Prepend a leftmost column holding a row counter, so the first data row
# (spreadsheet row 2, since row 1 is the header) is numbered
# ROW_NUMBER_START.
ADD_ROW_NUMBER: bool = False
ROW_NUMBER_START: int = 1

# Header text for that column. Set to "" for a blank heading.
ROW_NUMBER_HEADER: str = "S.No."

# Rename the technical field names to the friendly headers you asked for.
# Only applied to columns actually present in DESIRED_COLUMNS above.
RENAME_COLUMNS: Dict[str, str] = {
    # The eight exported columns, in DESIRED_COLUMNS order.
    "app": "Application",
    "user": "User",
    "url": "URL",
    "timestamp": "Event Date",
    "organization_unit": "Organization Unit",
    "client_bytes": "Bytes Upload",
    "server_bytes": "Bytes Download",
    "numbytes": "Total Bytes",
    # Kept in case they're re-added to DESIRED_COLUMNS. The activity/object
    # fields only exist on application events, not on page events.
    "usergroup": "User Group",
    "action": "Action",
    "alert": "Alert",
    "activity": "Activity",
    "object_type": "Object Type",
    "object": "Object Name",
    "file_size": "Sum - File Size (MB)",
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


def build_query() -> str:
    """
    Build the NQL query. The GenerativeAI application-category filter is
    the only filter applied.
    """
    return build_term(BASE_FILTER_FIELD, BASE_FILTER_OPERATOR, BASE_FILTER_VALUE)


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
            log.warning("Netskope timeout after %ss (attempt %d/%d) -- retry in %.0fs",
                        REQUEST_TIMEOUT, attempt, MAX_RETRIES, wait)
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
            log.warning("Netskope rate limited (429) -- sleeping %.0fs (attempt %d/%d)",
                        wait, attempt, MAX_RETRIES)
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

    # Always ask for _id: it is the per-event unique key the dedupe in
    # fetch_all relies on. Without it, distinct events whose exported
    # columns happen to match (same user/url/second/bytes) collapse into
    # one row.
    wanted: List[str] = ["_id"]
    for column in DESIRED_COLUMNS:
        if column in PLACEHOLDER_COLUMNS:
            continue
        # Ask for every known alias, since the exact name varies by schema.
        # Unknown names are ignored by the API.
        for name in FIELD_ALIASES.get(column, [column]):
            if name not in wanted:
                wanted.append(name)

    return ",".join(wanted) if wanted else None


def resolve_field_aliases(rows: List[Dict[str, Any]]) -> None:
    """
    Normalise field naming, in place.

    Whichever alias the tenant actually returned is copied onto the
    canonical name used in DESIRED_COLUMNS, so the export doesn't care
    whether the schema says `numbytes` or `numbytes_total`. Logs which alias
    won, so an empty byte column is diagnosable at a glance.
    """
    if not rows:
        return

    present: set = set()
    for row in rows[:50]:          # sampling is enough to spot the schema
        present.update(row.keys())

    for canonical, aliases in FIELD_ALIASES.items():
        if canonical in present:
            continue               # already the right name

        match = next((a for a in aliases if a in present), None)
        if match:
            log.info("field      %s -> using '%s'", canonical, match)
            for row in rows:
                if match in row:
                    row[canonical] = row[match]
        else:
            log.warning("field      %s NOT FOUND (tried: %s) -- that column "
                        "will export as 0", canonical, ", ".join(aliases))


def render_progress(stats: Dict[str, int], bar_width: int = 28) -> None:
    """
    Draw a single-line, in-place progress bar (overwrites itself with \r
    rather than scrolling the console with one line per page/window).

    Suppressed when stdout isn't a terminal -- in a scheduled run the output
    is usually redirected to a log file, and carriage-return animation turns
    into thousands of unreadable lines there.
    """
    if not sys.stdout.isatty():
        return

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


def window_params(query: str, starttime: int, endtime: int) -> Dict[str, Any]:
    """
    The query parameters every page request shares.

    `fields` and `timeout` are datasearch-only extras -- on /events/data/
    they are skipped entirely rather than sent and ignored, because an
    unrecognised parameter is a plausible cause of an HTTP 400 here.
    The caller adds the pagination parameter (offset or cursor).
    """
    params: Dict[str, Any] = {
        "query": query,
        "starttime": starttime,
        "endtime": endtime,
        "limit": PAGE_SIZE,
    }
    if SEND_QUERY_TIMEOUT:
        params["timeout"] = QUERY_TIMEOUT
    selected = fields_param()
    if selected:
        params["fields"] = selected
    return params


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
    # Cursor mode is only reached if the tenant returns a next-pointer. If it
    # ever hands back one we've already followed, the loop below would run
    # forever -- remember them and stop instead.
    seen_cursors: set = set()

    while True:
        if next_url:
            # Cursor mode: the API handed us a full URL or token. A bare token
            # still needs the query and the window bounds resent -- dropping
            # them made the follow-up page unfiltered and unbounded.
            if next_url.startswith("http"):
                target, params = next_url, None
            else:
                target = url
                params = window_params(query, starttime, endtime)
                params["cursor"] = next_url
        else:
            target = url
            params = window_params(query, starttime, endtime)
            params[OFFSET_PARAM_NAME] = offset

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
            if cursor in seen_cursors:
                log.warning("Pagination cursor repeated -- stopping this "
                            "window at %d record(s) rather than looping.",
                            len(records))
                break
            seen_cursors.add(cursor)
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


def next_local_midnight(day_start: int) -> int:
    """
    Midnight of the day AFTER the one starting at day_start, in local time.

    Not the same as day_start + 86400: on a DST changeover the local day is
    23 or 25 hours long, so plain addition lands an hour either side of
    midnight and the window silently gains or loses an hour of events.
    Overshooting into the middle of the next day and snapping back is
    correct on every day of the year.
    """
    return local_midnight(day_start + ONE_DAY + (4 * 3600))


def previous_local_midnight(day_start: int) -> int:
    """Midnight of the day BEFORE day_start. DST-safe, same reasoning."""
    return local_midnight(day_start - (4 * 3600))


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
        today = local_midnight(now)
        start = previous_local_midnight(today)
        end = today - BOUNDARY_TRIM

    elif DAY_MODE == "date":
        # strptime gives us the naive date; mktime turns it into local epoch.
        parsed = time.strptime(TARGET_DATE, "%Y-%m-%d")
        start = int(time.mktime(parsed))
        end = next_local_midnight(start) - BOUNDARY_TRIM

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
    Work out which day to export, in this order of precedence:

      1. A YYYY-MM-DD date passed on the command line
      2. The interactive prompt, if INTERACTIVE is True
      3. The DAY_MODE constant (default: yesterday)

    Step 2 is skipped entirely for unattended runs, so a scheduled job never
    blocks waiting on input that will never arrive.
    """
    # 1. Command line wins.
    # First non-flag argument, if any, is treated as the date.
    cli_args = [a for a in sys.argv[1:] if not a.startswith("-")]
    cli_date = cli_args[0].strip() if cli_args else ""
    if cli_date:
        try:
            parsed = time.strptime(cli_date, "%Y-%m-%d")
        except ValueError:
            print(f"'{cli_date}' isn't a valid date (expected YYYY-MM-DD).")
            raise
        start = int(time.mktime(parsed))
        return start, next_local_midnight(start) - BOUNDARY_TRIM, cli_date

    # 2. No prompt on unattended runs.
    if not INTERACTIVE:
        return resolve_day_range()

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
    end = next_local_midnight(start) - BOUNDARY_TRIM
    return start, end, entry


def resolve_output_path(label: str) -> str:
    """
    Append the day being exported to the filename when configured, keeping
    whatever extension OUTPUT_FILE uses (.xlsx, .csv, ...).
    """
    if DATE_STAMP_FILENAME and "." in OUTPUT_FILE:
        stem, _, extension = OUTPUT_FILE.rpartition(".")
        name = f"{stem}_{label}.{extension}"
    elif DATE_STAMP_FILENAME:
        name = f"{OUTPUT_FILE}_{label}"
    else:
        name = OUTPUT_FILE

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    return os.path.join(OUTPUT_DIR, name)


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

    if sys.stdout.isatty():
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
    'User@Example.com' with 'user@example.com' and 'ChatGPT ' with
    'ChatGPT', and a raw string compare would treat those as distinct rows.

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


def drop_zero_rows(rows: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Remove rows where every field in DROP_IF_ALL_ZERO is zero.

    A missing or unparseable value counts as zero, which is the sensible
    reading for "this event moved no data" -- but it also means that if the
    byte fields never resolved to a real column, EVERY row looks like a
    zero row and the export would come back empty. The caller compares the
    before/after counts and says so plainly when that happens, rather than
    letting an empty CSV pass for a quiet day.
    """
    if not DROP_IF_ALL_ZERO:
        return rows
    return [row for row in rows
            if any(to_number(row.get(f)) != 0 for f in DROP_IF_ALL_ZERO)]


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
        if DATE_FORMAT.strip().lower() == "ctime":
            return time.ctime(epoch)   # Tue Dec  9 21:00:20 2025
        return time.strftime(DATE_FORMAT, time.localtime(epoch))
    except (ValueError, OSError, OverflowError):
        return value          # out of range for the platform's localtime


def to_int_text(value: Any) -> Any:
    """
    Render a numeric value as a whole number, so a byte count reads 54665
    rather than 54665.0. Blanks stay blank and anything unparseable is left
    exactly as it arrived.
    """
    if value is None or value == "":
        return ""
    try:
        return int(float(value))
    except (TypeError, ValueError):
        return value


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
    export path (now and anything added later) is guaranteed to
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
        elif column in INT_FIELDS:
            frame[column] = frame[column].apply(to_int_text)

    if RENAME_COLUMNS:
        frame = frame.rename(columns=RENAME_COLUMNS)

    # Unnamed leftmost counter column. Built as an explicit list (not a
    # range) and re-assembled by column order rather than using insert(),
    # so it behaves identically across pandas versions.
    if ADD_ROW_NUMBER:
        numbers = [ROW_NUMBER_START + i for i in range(len(frame))]
        frame = frame.copy()
        frame[ROW_NUMBER_HEADER] = numbers
        ordered = [ROW_NUMBER_HEADER] + [
            c for c in frame.columns if c != ROW_NUMBER_HEADER
        ]
        frame = frame[ordered]

    if path.lower().endswith((".xlsx", ".xlsm")):
        # index=False keeps the unnamed row-number column as the real
        # first column rather than letting pandas write its own index.
        frame.to_excel(
            path,
            sheet_name=EXCEL_SHEET_NAME,
            index=False,
            engine="openpyxl",
        )
    else:
        frame.to_csv(path, index=False, quoting=csv.QUOTE_MINIMAL,
                     encoding="utf-8")


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
    params: Dict[str, Any] = {"starttime": start, "endtime": end,
                              "limit": 5, OFFSET_PARAM_NAME: 0}
    if SEND_QUERY_TIMEOUT:
        params["timeout"] = QUERY_TIMEOUT

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
    # "categ"/"org" cover the filter field; "byte" covers the three measure
    # columns, which are the ones most likely to come back under an
    # unexpected name on a given tenant's schema.
    candidates = [c for c in cols
                  if any(term in c.lower() for term in
                         ("categ", "org", "unit", "ou", "byte"))]
    for field in candidates:
        values = sorted({str(row[field]) for row in flat_sample
                          if row.get(field) not in (None, "")})
        if values:
            print(f"    {field} sample value(s): {', '.join(values[:5])}")


# ---------------------------------------------------------------------------
# SharePoint upload (Microsoft Graph)
# ---------------------------------------------------------------------------
# Uses the OAuth 2.0 client credentials flow -- no signed-in user, which is
# what lets this run unattended. Requires the `msal` package:
#     pip install msal
#
# The app registration needs Microsoft Graph APPLICATION permission
# Sites.Selected, plus a site-level grant on the target site. Sites.Selected
# alone grants nothing until that site grant is added.
# ---------------------------------------------------------------------------

GRAPH_ROOT: str = "https://graph.microsoft.com/v1.0"

# Graph requires upload-session chunks to be a multiple of 320 KiB.
CHUNK_SIZE: int = 320 * 1024 * 10          # 3.2 MB per chunk
SIMPLE_UPLOAD_LIMIT: int = 4 * 1024 * 1024  # 4 MB: above this, use a session


class UploadError(RuntimeError):
    """Raised when the SharePoint upload can't be completed."""


def graph_credentials() -> Tuple[str, str, str]:
    """
    Return the SPN credentials, raising a clear error if they're still the
    placeholder values rather than letting Graph fail with a vague 401.
    """
    tenant, client, secret = TENANT_ID, CLIENT_ID, CLIENT_SECRET

    missing = [
        name for name, value in
        (("tenant id", tenant), ("client id", client), ("client secret", secret))
        if not value or value.startswith("PUT_")
    ]
    if missing:
        raise UploadError(
            f"Missing SharePoint credential(s): {', '.join(missing)}. "
            f"Set TENANT_ID / CLIENT_ID / CLIENT_SECRET in the CONFIGURATION "
            f"section of {os.path.basename(__file__)}."
        )
    return tenant, client, secret


def get_graph_token() -> str:
    """Acquire an app-only Graph token via the client credentials flow."""
    try:
        import msal
    except ImportError:
        raise UploadError(
            "The 'msal' package is required for the SharePoint upload. "
            "Install it with:  pip install msal"
        )

    tenant, client, secret = graph_credentials()
    app = msal.ConfidentialClientApplication(
        client_id=client,
        client_credential=secret,
        authority=f"https://login.microsoftonline.com/{tenant}",
    )

    # .default asks for whatever application permissions were consented.
    result = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    if "access_token" not in result:
        raise UploadError(
            f"Token request failed: {result.get('error')} -- "
            f"{result.get('error_description', '')[:300]}"
        )
    return result["access_token"]


def graph_get(session: requests.Session, url: str) -> Dict[str, Any]:
    """GET a Graph endpoint and return the decoded JSON, or raise."""
    response = session.get(url, timeout=REQUEST_TIMEOUT)

    # Both of these mean the same thing under Sites.Selected: the app
    # authenticated fine but has no permission on THIS site. SharePoint
    # reports it inconsistently -- sometimes a clean 403, sometimes a 401
    # wrapped as "generalException / spException". Treat them together.
    body = response.text or ""
    is_site_grant_issue = (
        response.status_code == 403
        or (response.status_code == 401 and
            ("spException" in body or "generalException" in body))
    )
    if is_site_grant_issue:
        raise UploadError(
            f"Graph returned {response.status_code} for {url}.\n"
            f"  This is the classic Sites.Selected symptom: the app has the "
            f"API permission but has NOT been granted access to this "
            f"specific site.\n"
            f"  The API permission alone grants nothing -- an admin must run "
            f"the site-level grant once.\n"
            f"  Run  python {os.path.basename(sys.argv[0])} --grant-help  "
            f"for the exact command."
        )
    if not response.ok:
        raise UploadError(
            f"Graph GET {url} failed ({response.status_code}): "
            f"{response.text[:300]}"
        )
    return response.json()


def print_grant_help() -> None:
    """
    Print the one-time site-permission grant an admin needs to run, with the
    app's own IDs filled in. This is what fixes the 401/403 spException.
    """
    try:
        _, client, _ = graph_credentials()
    except UploadError:
        client = CLIENT_ID

    path = SITE_PATH if SITE_PATH.startswith("/") else "/" + SITE_PATH
    print("=" * 70)
    print("ONE-TIME SharePoint site grant (run by a SharePoint / Graph admin)")
    print("=" * 70)
    print()
    print("Sites.Selected grants NO access until the app is granted rights on")
    print("the specific site. Do that once with the two calls below.")
    print()
    print("STEP 1 -- get the site id (any admin token with Sites.Read.All):")
    print(f"  GET {GRAPH_ROOT}/sites/{SITE_HOSTNAME}:{path}")
    print("  -> copy the \"id\" from the response")
    print()
    print("STEP 2 -- grant this app WRITE on that site:")
    print(f"  POST {GRAPH_ROOT}/sites/{{site-id}}/permissions")
    print("  Content-Type: application/json")
    print()
    print("  {")
    print('    "roles": ["write"],')
    print('    "grantedToIdentities": [{')
    print('      "application": {')
    print(f'        "id": "{client}",')
    print('        "displayName": "Netskope GenAI Export"')
    print("      }")
    print("    }]")
    print("  }")
    print()
    print("Both steps can be run from Graph Explorer (aka.ms/ge) signed in as")
    print("an admin, or via PowerShell (Get-MgSite / New-MgSitePermission).")
    print("After the grant, re-run the export normally.")
    print("=" * 70)


def resolve_site_id(session: requests.Session) -> str:
    """Look up the site's Graph ID from hostname + path, unless preset."""
    if SITE_ID:
        return SITE_ID
    path = SITE_PATH if SITE_PATH.startswith("/") else "/" + SITE_PATH
    url = f"{GRAPH_ROOT}/sites/{SITE_HOSTNAME}:{path}"
    site = graph_get(session, url)
    if "id" not in site:
        raise UploadError(f"Could not resolve site id from {url}")
    return site["id"]


def resolve_drive_id(session: requests.Session, site_id: str) -> str:
    """Find the document library ('drive') by name, unless preset."""
    if DRIVE_ID:
        return DRIVE_ID
    payload = graph_get(session, f"{GRAPH_ROOT}/sites/{site_id}/drives")
    drives = payload.get("value", [])
    for drive in drives:
        if drive.get("name", "").lower() == LIBRARY_NAME.lower():
            return drive["id"]
    available = ", ".join(d.get("name", "?") for d in drives) or "(none)"
    raise UploadError(
        f"No document library named '{LIBRARY_NAME}' on this site. "
        f"Available: {available}"
    )


def remote_path(filename: str) -> str:
    """Build the destination path inside the library."""
    folder = TARGET_FOLDER.strip("/")
    return f"{folder}/{filename}" if folder else filename


def upload_small(session: requests.Session, drive_id: str,
                 local_path: str, destination: str) -> Dict[str, Any]:
    """Single PUT. Only valid for files under ~4 MB."""
    url = (f"{GRAPH_ROOT}/drives/{drive_id}/root:/{destination}:/content"
           f"?@microsoft.graph.conflictBehavior="
           f"{'replace' if SHAREPOINT_REPLACE_EXISTING else 'fail'}")
    with open(local_path, "rb") as handle:
        response = session.put(url, data=handle, timeout=REQUEST_TIMEOUT)
    if not response.ok:
        raise UploadError(
            f"Upload failed ({response.status_code}): {response.text[:300]}"
        )
    return response.json()


def upload_large(session: requests.Session, drive_id: str,
                 local_path: str, destination: str, size: int) -> Dict[str, Any]:
    """
    Chunked upload for files over the 4 MB simple-upload ceiling. Graph
    hands back an upload URL, then each chunk is PUT with a Content-Range.
    """
    create_url = (f"{GRAPH_ROOT}/drives/{drive_id}/root:/{destination}:"
                  f"/createUploadSession")
    body = {
        "item": {
            "@microsoft.graph.conflictBehavior":
                "replace" if SHAREPOINT_REPLACE_EXISTING else "fail"
        }
    }
    response = session.post(create_url, json=body, timeout=REQUEST_TIMEOUT)
    if not response.ok:
        raise UploadError(
            f"Could not start upload session ({response.status_code}): "
            f"{response.text[:300]}"
        )
    upload_url = response.json().get("uploadUrl")
    if not upload_url:
        raise UploadError("Upload session response had no uploadUrl.")

    sent = 0
    chunk_response: Optional[requests.Response] = None
    with open(local_path, "rb") as handle:
        while sent < size:
            chunk = handle.read(CHUNK_SIZE)
            if not chunk:
                break
            last = sent + len(chunk) - 1
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {sent}-{last}/{size}",
            }
            # NOTE: no auth header here on purpose -- the upload URL is
            # pre-authorized, and sending one can cause a 401.
            chunk_response = requests.put(
                upload_url, data=chunk, headers=headers,
                timeout=REQUEST_TIMEOUT,
            )
            if chunk_response.status_code not in (200, 201, 202):
                raise UploadError(
                    f"Chunk upload failed at byte {sent} "
                    f"({chunk_response.status_code}): "
                    f"{chunk_response.text[:300]}"
                )
            sent += len(chunk)
            if sys.stdout.isatty():
                percent = int(sent / size * 100)
                print(f"\r  uploading... {percent:3d}%", end="", flush=True)

    if sys.stdout.isatty():
        print()

    # chunk_response is None only if the file read empty, in which case no
    # chunk was ever PUT -- referencing it unguarded raised UnboundLocalError.
    if chunk_response is None:
        raise UploadError(f"Nothing was uploaded: {local_path} read as empty.")
    return chunk_response.json() if chunk_response.content else {}


def upload_to_sharepoint(local_path: str) -> str:
    """
    Upload a finished file to SharePoint. Returns the resulting web URL.
    Picks the simple or chunked path based on file size.
    """
    size = 0
    with open(local_path, "rb") as handle:
        handle.seek(0, 2)
        size = handle.tell()

    token = get_graph_token()
    session = requests.Session()
    session.headers["Authorization"] = f"Bearer {token}"
    session.headers["Accept"] = "application/json"

    site_id = resolve_site_id(session)
    drive_id = resolve_drive_id(session, site_id)
    destination = remote_path(os.path.basename(local_path))

    if size <= SIMPLE_UPLOAD_LIMIT:
        item = upload_small(session, drive_id, local_path, destination)
    else:
        item = upload_large(session, drive_id, local_path, destination, size)

    return item.get("webUrl", "(uploaded)")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# TERMINAL MULTIPLEXER (screen / tmux)  -- Linux only
# ---------------------------------------------------------------------------
# On the FIRST, interactive run the script can relaunch itself inside a
# named screen session ("Netskope GenAI") so a long extraction survives the
# SSH connection dropping. You can detach with Ctrl-A D and reattach later
# with:
#     screen -r "Netskope GenAI"
#
# Three guards keep this from causing trouble:
#
#   1. NOT under a scheduler. If stdin/stdout isn't a terminal -- which is
#      how cron and systemd run things -- the wrapper is skipped entirely.
#      Without this you'd accumulate one orphaned session per scheduled
#      run, and cron jobs would appear to "succeed" instantly while the
#      real work happened invisibly elsewhere.
#   2. NOT already inside one. STY/TMUX are set inside a session; the
#      NETSKOPE_IN_SCREEN marker covers the relaunch itself. Together they
#      make infinite re-spawning impossible.
#   3. NOT if the binary is missing. RHEL 9 does not ship `screen` in the
#      base repos (it moved to EPEL), so tmux is used as a fallback and a
#      plain foreground run is the last resort.
# ---------------------------------------------------------------------------

# False while testing -- the relaunch only earns its keep on a long SSH run.
SCREEN_ENABLED: bool = False
SCREEN_NAME: str = "Netskope GenAI Traffic"

# "auto" tries screen then tmux; force one with "screen" or "tmux";
# "none" disables the wrapper the same as SCREEN_ENABLED = False.
SCREEN_TOOL: str = "auto"

# Guard variable set on the relaunched child.
IN_SCREEN_MARKER: str = "NETSKOPE_IN_SCREEN"


def running_interactively() -> bool:
    """
    True only when a human is at the terminal. cron, systemd, Task
    Scheduler and CI all fail this check, which is exactly what we want --
    the screen wrapper must never engage for a scheduled run.
    """
    try:
        return sys.stdin.isatty() and sys.stdout.isatty()
    except (AttributeError, ValueError):
        return False


def already_in_session() -> bool:
    """True if we're inside screen/tmux, or are the relaunched child."""
    return bool(
        os.environ.get(IN_SCREEN_MARKER)
        or os.environ.get("STY")       # set by screen
        or os.environ.get("TMUX")      # set by tmux
    )


def pick_multiplexer() -> Optional[str]:
    """Return the path to screen or tmux, honouring SCREEN_TOOL."""
    choice = SCREEN_TOOL.strip().lower()
    if choice in ("none", "off", "false"):
        return None
    order = {
        "screen": ["screen"],
        "tmux": ["tmux"],
    }.get(choice, ["screen", "tmux"])
    for tool in order:
        path = shutil.which(tool)
        if path:
            return path
    return None


def relaunch_in_session() -> Optional[int]:
    """
    Re-run this script inside a named screen/tmux session and attach to it.

    Returns the child's exit code if the process was handed off (the caller
    should exit with it), or None if we should carry on in the foreground.

    Returning the child's code matters: handing back a bare True made the
    parent exit 0 even when the real run failed, hiding every failure from
    whatever is watching exit codes.
    """
    if not SCREEN_ENABLED or os.name != "posix":
        return None
    if already_in_session() or not running_interactively():
        return None

    tool_path = pick_multiplexer()
    if not tool_path:
        print("Note: neither 'screen' nor 'tmux' found -- running in the "
              "foreground.\n"
              "      On RHEL 9:  sudo dnf install -y tmux\n"
              "      (screen lives in EPEL:  sudo dnf install -y epel-release "
              "screen)\n")
        return None

    tool = os.path.basename(tool_path)
    python = sys.executable or "python3"
    script = os.path.abspath(__file__)
    args = [a for a in sys.argv[1:]]

    child_env = dict(os.environ)
    child_env[IN_SCREEN_MARKER] = "1"

    if tool == "screen":
        # -S names it, -m forces a new session even if one exists.
        command = [tool_path, "-S", SCREEN_NAME, "-m", python, script] + args
        reattach = f'screen -r "{SCREEN_NAME}"'
        detach_keys = "Ctrl-A then D"
    else:
        command = [tool_path, "new-session", "-s", SCREEN_NAME,
                   python, script] + args
        reattach = f'tmux attach -t "{SCREEN_NAME}"'
        detach_keys = "Ctrl-B then D"

    print(f"Starting inside {tool} session: {SCREEN_NAME}")
    print(f"  detach   : {detach_keys}")
    print(f"  reattach : {reattach}")
    print()

    try:
        completed = subprocess.run(command, env=child_env)
        return completed.returncode
    except FileNotFoundError:
        return None
    except Exception as exc:
        print(f"Could not start {tool} ({exc}) -- running in the foreground.")
        return None


# ---------------------------------------------------------------------------
# RUN STATUS (used by --check)
# ---------------------------------------------------------------------------
# Each export writes a small JSON status file recording what happened for
# that day: whether the CSV was written, whether the upload succeeded, and
# the row count. The later --check run reads these instead of re-querying
# Netskope, so verifying costs nothing.
# ---------------------------------------------------------------------------

# Prefix for this script's status files. MUST differ from the companion
# script's, or two exports of the same day sharing LOG_DIR would overwrite
# each other's status and --check would report on the wrong run.
STATUS_PREFIX: str = "traffic"


def status_path(label: str) -> str:
    """Path of the status file for a given export day."""
    folder = os.path.join(LOG_DIR, "status")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, f"run_{STATUS_PREFIX}_{label}.json")


def write_status(label: str, **fields: Any) -> None:
    """
    Record (or update) the outcome for a day. Merges with anything already
    written, so a later upload retry can flip 'uploaded' to True without
    losing the export details.
    """
    path = status_path(label)
    data: Dict[str, Any] = {}
    try:
        with open(path, "r", encoding="utf-8") as handle:
            data = json.load(handle)
    except (FileNotFoundError, ValueError):
        pass

    data.update(fields)
    data["day"] = label
    data["updated"] = time.strftime("%Y-%m-%d %H:%M:%S")

    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(data, handle, indent=2)
    except OSError as exc:
        log.warning("Could not write status file %s: %s", path, exc)


def read_status(label: str) -> Dict[str, Any]:
    """Read a day's status file; empty dict if it doesn't exist."""
    try:
        with open(status_path(label), "r", encoding="utf-8") as handle:
            return json.load(handle)
    except (FileNotFoundError, ValueError):
        return {}


def run_check(run_id: str, started: float) -> int:
    """
    Verification pass -- does NOT re-export.

    Looks at the day that should have been exported and reports whether it
    actually was, and whether it reached SharePoint. If the CSV exists but
    the upload failed earlier, the upload is retried here (that's the cheap
    fix that makes a second daily run worth scheduling).

    Exit codes match the main run so alerting logic stays identical.
    """
    log.info("MODE       check (verify only -- no extraction)")

    try:
        _, _, label = prompt_for_date_range()
    except ValueError as exc:
        log.error("CONFIG ERROR: bad date -- %s", exc)
        return finish(1, started, run_id)

    expected_csv = resolve_output_path(label)
    status = read_status(label)
    csv_exists = os.path.exists(expected_csv)

    log.info("checking   %s", label)
    log.info("expected   %s", expected_csv)

    # --- Did the export happen at all? ------------------------------------
    if not csv_exists:
        log.error("CHECK FAILED: no export file for %s.", label)
        if status:
            log.error("  Last attempt at %s ended: %s",
                      status.get("updated", "?"),
                      status.get("result", "unknown"))
            if status.get("error"):
                log.error("  Reported error: %s", status["error"])
        else:
            log.error("  No run was recorded for this day at all -- the "
                      "scheduled export did not fire, or died before it "
                      "could log anything.")
        log.error("  Backfill with:  python %s %s",
                  os.path.basename(sys.argv[0]), label)
        return finish(2, started, run_id)

    rows = status.get("rows")
    if rows is not None:
        log.info("export     OK (%s rows)", rows)
    else:
        log.info("export     OK")

    # --- Did the upload happen? -------------------------------------------
    if not SHAREPOINT_ENABLED:
        log.info("upload     skipped (SHAREPOINT_ENABLED=false)")
        log.info("CHECK PASSED for %s", label)
        return finish(0, started, run_id)

    if status.get("uploaded"):
        log.info("upload     OK (%s)", status.get("upload_url", "previously uploaded"))
        log.info("CHECK PASSED for %s", label)
        return finish(0, started, run_id)

    # CSV is on disk but never made it up -- retry now.
    log.warning("upload     MISSING for %s -- retrying now.", label)
    if status.get("error"):
        log.warning("  Earlier failure was: %s", status["error"])

    try:
        web_url = upload_to_sharepoint(expected_csv)
    except (UploadError, requests.exceptions.RequestException) as exc:
        log.error("SHAREPOINT RETRY FAILED: %s", exc)
        log.error("The CSV is still safe at %s", expected_csv)
        write_status(label, result="upload_failed", error=str(exc)[:500])
        return finish(4, started, run_id)

    log.info("SHAREPOINT OK (on retry): %s", web_url)
    write_status(label, uploaded=True, upload_url=web_url,
                 result="success_after_retry", error="")
    log.info("CHECK PASSED for %s (upload recovered)", label)
    return finish(0, started, run_id)


def main() -> int:
    """
    Returns a process exit code so a scheduler can detect failures:
        0 = success (or nothing to do)
        1 = configuration / input error
        2 = NETSKOPE api or network failure
        3 = ran fine but the filter matched no records
        4 = data exported locally, but the SHAREPOINT upload failed

    2 vs 4 is the distinction that matters operationally: 2 means no data
    was collected, 4 means the data is on disk and only the upload needs
    retrying.

    Modes:
        (no flag)   full export for the target day
        --check     verify only; no extraction. Confirms the day's export
                    and upload happened, and retries a failed upload.
    """
    stats: Dict[str, int] = {"calls": 0}
    started = time.time()

    # Run id ties every line of a run together, so two runs landing in the
    # same day's log file stay easy to separate.
    run_id = time.strftime("%H%M%S")
    log_path = setup_logging(run_id)

    log.info("=" * 64)
    log.info("RUN START  %s", time.strftime("%Y-%m-%d %H:%M:%S"))
    log.info("config     in-script (%s)", os.path.basename(__file__))
    log.info("log file   %s", log_path)

    archive_old_logs()

    # Verification-only pass: no extraction, just confirm the day's export
    # and upload actually happened (and retry the upload if it didn't).
    if "--check" in sys.argv:
        return run_check(run_id, started)

    if "PUT_YOUR_API_TOKEN" in API_TOKEN or "YOUR_TENANT" in BASE_URL:
        log.error("CONFIG ERROR: API_TOKEN / BASE_URL are still placeholders.")
        log.error("  Set them in the CONFIGURATION section near the top of %s.",
                  os.path.basename(__file__))
        return finish(1, started, run_id)

    try:
        query = build_query()
    except ValueError as exc:
        log.error("CONFIG ERROR: bad filter -- %s", exc)
        return finish(1, started, run_id)

    try:
        start, end, label = prompt_for_date_range()
    except ValueError as exc:
        log.error("CONFIG ERROR: bad date -- %s", exc)
        return finish(1, started, run_id)

    output_path = resolve_output_path(label)

    # Don't redo a day that's already been exported.
    if SKIP_IF_EXISTS and os.path.exists(output_path):
        log.info("SKIPPED: %s already exists -- nothing to do.", output_path)
        return finish(0, started, run_id)

    day_source = "cli" if [a for a in sys.argv[1:] if not a.startswith("-")] \
        else DAY_MODE
    log.info("day        %s (%s)", label, day_source)
    log.info("query      %s", query)
    log.info("endpoint   %s%s", BASE_URL.rstrip("/"), ENDPOINT_PATH)

    session = build_session()

    log.info("Fetching from Netskope...")
    try:
        records = fetch_all(session, query, start, end, stats)
    except ApiError as exc:
        log.error("NETSKOPE FAILED: %s", exc)
        log.error("No data was collected. Nothing written, nothing uploaded.")
        write_status(label, result="netskope_failed", error=str(exc)[:500],
                     uploaded=False)
        return finish(2, started, run_id)
    except requests.exceptions.RequestException as exc:
        log.error("NETSKOPE FAILED (network): %s", exc)
        write_status(label, result="netskope_failed", error=str(exc)[:500],
                     uploaded=False)
        return finish(2, started, run_id)
    except KeyboardInterrupt:
        log.warning("Interrupted by user.")
        return finish(1, started, run_id)

    if not records:
        log.warning("NETSKOPE returned no records for %s (%d API calls).",
                    label, stats["calls"])
        write_status(label, result="no_records", rows=0, uploaded=False)
        try:
            diagnose_zero_results(session, start, end, stats)
        except Exception as exc:
            log.warning("Zero-result diagnostic failed: %s", exc)
        return finish(3, started, run_id)

    flat_rows = [flatten_record(r) for r in records]
    resolve_field_aliases(flat_rows)
    discovered = discover_columns(flat_rows)
    raw_row_count = len(flat_rows)
    flat_rows = dedupe_rows(flat_rows)

    before_zero_filter = len(flat_rows)
    flat_rows = drop_zero_rows(flat_rows)
    dropped_zero = before_zero_filter - len(flat_rows)

    # Everything dropped almost never means "a whole day of zero-byte
    # events". It means the byte fields aren't where we think they are, so
    # every row read as zero. Say that rather than writing a header-only CSV.
    if DROP_IF_ALL_ZERO and before_zero_filter and not flat_rows:
        log.error("Every one of the %d row(s) had zero bytes in %s and was "
                  "dropped.", before_zero_filter, " and ".join(DROP_IF_ALL_ZERO))
        log.error("  That usually means the byte fields resolved to nothing "
                  "-- check the 'field ... NOT FOUND' warnings above and the "
                  "FIELD_ALIASES list, not the traffic itself.")
        log.error("  To export anyway, set DROP_IF_ALL_ZERO = [].")
        write_status(label, result="all_rows_zero", rows=0, uploaded=False)
        return finish(3, started, run_id)

    if DESIRED_COLUMNS:
        missing = [c for c in DESIRED_COLUMNS if c not in discovered]
        if missing:
            log.warning("Requested columns absent from the data "
                        "(exported blank): %s", ", ".join(missing))
        columns = DESIRED_COLUMNS
    else:
        columns = discovered

    try:
        export_csv(flat_rows, columns, output_path)
    except OSError as exc:
        log.error("WRITE FAILED: could not write %s -- %s", output_path, exc)
        return finish(1, started, run_id)

    log.info("-" * 52)
    log.info("Day exported            : %s", label)
    log.info("Total records retrieved : %s", raw_row_count)
    if dropped_zero:
        log.info("Zero-traffic rows dropped: %s", dropped_zero)
    if DEDUPE_ON:
        log.info("Unique rows exported    : %s (by %s)",
                 len(flat_rows), "+".join(DEDUPE_ON))
    else:
        log.info("Rows exported           : %s (one per event, not grouped)",
                 len(flat_rows))
    log.info("Columns exported        : %s", len(columns))
    log.info("API calls made          : %s", stats["calls"])
    log.info("Exported to             : %s", output_path)
    log.info("-" * 52)

    write_status(label, result="exported", rows=len(flat_rows),
                 csv=output_path, uploaded=False, error="")

    # Upload happens after the local file is safely written, so a failure
    # here never costs you the extracted data -- you can re-upload later.
    if SHAREPOINT_ENABLED:
        log.info("Uploading to SharePoint...")
        try:
            web_url = upload_to_sharepoint(output_path)
            log.info("SHAREPOINT OK: %s", web_url)
            write_status(label, result="success", uploaded=True,
                         upload_url=web_url, error="")
        except UploadError as exc:
            log.error("SHAREPOINT FAILED: %s", exc)
            log.error("The CSV is safe locally at %s -- the next --check run "
                      "will retry the upload automatically.", output_path)
            write_status(label, result="upload_failed", error=str(exc)[:500],
                         uploaded=False)
            return finish(4, started, run_id)
        except requests.exceptions.RequestException as exc:
            log.error("SHAREPOINT FAILED (network): %s", exc)
            log.error("The CSV is safe locally at %s", output_path)
            write_status(label, result="upload_failed", error=str(exc)[:500],
                         uploaded=False)
            return finish(4, started, run_id)
    else:
        log.info("SharePoint upload disabled (SHAREPOINT_ENABLED=false).")
        write_status(label, result="exported_no_upload")

    return finish(0, started, run_id)


EXIT_MEANING: Dict[int, str] = {
    0: "SUCCESS",
    1: "CONFIG/INPUT ERROR",
    2: "NETSKOPE FAILURE",
    3: "NO RECORDS",
    4: "SHAREPOINT UPLOAD FAILURE",
}


def finish(code: int, started: float, run_id: str) -> int:
    """Write the closing banner. Every run ends with one line you can grep."""
    elapsed = time.time() - started
    level = logging.INFO if code == 0 else logging.ERROR
    if code == 3:
        level = logging.WARNING
    log.log(level, "RUN END    status=%s exit=%d elapsed=%.0fs",
            EXIT_MEANING.get(code, "UNKNOWN"), code, elapsed)
    log.info("=" * 64)
    logging.shutdown()
    return code


if __name__ == "__main__":
    if "--grant-help" in sys.argv:
        print_grant_help()
        sys.exit(0)

    # Interactive first run: hand off into a named screen/tmux session so
    # the job survives a dropped SSH connection. Silently skipped under
    # cron/systemd, and when --no-screen is passed. The child's exit code is
    # passed straight through so a failure inside the session still surfaces.
    if "--no-screen" not in sys.argv:
        child_code = relaunch_in_session()
        if child_code is not None:
            sys.exit(child_code)

    # Exit code lets Task Scheduler / cron flag a failed run.
    sys.exit(main())
