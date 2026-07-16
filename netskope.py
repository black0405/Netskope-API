#!/usr/bin/env python3
"""
Daily Netskope GenAI application-events export for Power BI.

Runs once a day (via Task Scheduler / cron), pulls ONE full calendar day of
application events (by default the previous day, 00:00 -> next 00:00) filtered
by category (default: generativeai) plus any extra filters, and writes one CSV
per day into an output folder that Power BI's Folder connector points at.
Each file is dated by the DAY it represents, so days never overlap.

Designed for unattended scheduling:
  * Stable schema  -> pin --fields so every day's CSV has identical columns
                      (Power BI "Combine & Transform" on a folder needs this).
  * Atomic write   -> writes to a .tmp file then renames, so the gateway
                      never reads a half-written file.
  * Log file       -> <output-dir>/logs/ plus stderr, for scheduler visibility.
  * Exit codes     -> non-zero on failure so the scheduler can alert.

Auth / connection (env vars or CLI):
    NETSKOPE_TENANT     e.g. acme -> https://acme.goskope.com  (--tenant)
    NETSKOPE_API_TOKEN  REST API v2 token with the "events" scope (--token)
    NETSKOPE_OUTPUT_DIR folder Power BI reads from (--output-dir)

Example (what the scheduler runs):
    python netskope_genai_daily.py \
        --output-dir "D:/PowerBI/netskope_genai" \
        --fields "timestamp,user,app,activity,category,ccl,url,action"
"""

import argparse
import csv
import json
import logging
import os
import sys
import time
from datetime import datetime, timezone, timedelta
from urllib.parse import urlparse

import requests

DEFAULT_LIMIT = 5000
MAX_PAGES = 10_000
MAX_RETRIES = 5
BACKOFF_BASE = 2.0

logger = logging.getLogger("netskope_genai_daily")


# --------------------------------------------------------------------------- #
# Connection / query helpers (shared with the ad-hoc extractor)
# --------------------------------------------------------------------------- #
def build_base_url(tenant: str) -> str:
    tenant = tenant.strip().rstrip("/")
    if tenant.startswith(("http://", "https://")):
        host = urlparse(tenant).netloc
    elif "." in tenant:
        host = tenant
    else:
        host = f"{tenant}.goskope.com"
    return f"https://{host}"


def quote_value(value: str) -> str:
    value = value.strip()
    if value and " " not in value and "'" not in value and value == value.strip('"'):
        return value
    inner = value.strip('"').replace('"', '\\"')
    return f'"{inner}"'


def build_query(category: str | None, extra_filters: list[str]) -> str | None:
    clauses: list[str] = []
    if category:
        clauses.append(f"category eq {quote_value(category)}")
    for f in extra_filters:
        f = f.strip()
        if not f:
            continue
        parts = f.split(None, 2)
        if len(parts) == 3:
            field, op, value = parts
            clauses.append(f"{field} {op} {quote_value(value)}")
        else:
            clauses.append(f)
    return " and ".join(clauses) if clauses else None


def request_page(session, url, headers, params):
    for attempt in range(1, MAX_RETRIES + 1):
        resp = session.get(url, headers=headers, params=params, timeout=60)
        if resp.status_code == 200:
            return resp.json()
        if resp.status_code == 429 or resp.status_code >= 500:
            retry_after = resp.headers.get("Retry-After")
            wait = float(retry_after) if retry_after else BACKOFF_BASE ** attempt
            logger.warning("HTTP %s, retry %d/%d in %.1fs",
                           resp.status_code, attempt, MAX_RETRIES, wait)
            time.sleep(wait)
            continue
        raise RuntimeError(f"HTTP {resp.status_code} from Netskope: {resp.text[:500]}")
    raise RuntimeError(f"Gave up after {MAX_RETRIES} retries.")


def extract_records(payload) -> list[dict]:
    if isinstance(payload, dict):
        for key in ("result", "data", "events"):
            if isinstance(payload.get(key), list):
                return payload[key]
    return []


def fetch_all(base_url, token, query, starttime, endtime, limit):
    url = f"{base_url}/api/v2/events/data/application"
    headers = {"Netskope-Api-Token": token, "Accept": "application/json"}
    session = requests.Session()
    all_records: list[dict] = []
    skip = 0

    for page in range(MAX_PAGES):
        params = {"starttime": starttime, "endtime": endtime,
                  "limit": limit, "skip": skip}
        if query:
            params["query"] = query
        payload = request_page(session, url, headers, params)
        records = extract_records(payload)
        all_records.extend(records)
        logger.info("page %d: %d records (running total %d)",
                    page + 1, len(records), len(all_records))
        if len(records) < limit:
            break
        skip += limit
    return all_records


# --------------------------------------------------------------------------- #
# Output: schema-stable, atomic CSV for Power BI
# --------------------------------------------------------------------------- #
def resolve_columns(records, pinned_fields):
    """Pinned fields give an identical schema every day (best for Power BI).
    Without them, fall back to a deterministic sorted union of all keys."""
    if pinned_fields:
        return pinned_fields
    keys = set()
    for r in records:
        keys.update(r.keys())
    return sorted(keys)


def write_csv_atomic(records, out_path, columns):
    tmp_path = out_path + ".tmp"
    with open(tmp_path, "w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        writer.writeheader()
        for r in records:
            writer.writerow(
                {c: (json.dumps(r[c]) if isinstance(r.get(c), (dict, list))
                     else r.get(c, "")) for c in columns}
            )
    os.replace(tmp_path, out_path)  # atomic on the same filesystem


def setup_logging(output_dir):
    log_dir = os.path.join(output_dir, "logs")
    os.makedirs(log_dir, exist_ok=True)
    log_file = os.path.join(
        log_dir, f"export_{datetime.now(timezone.utc):%Y-%m-%d}.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
        handlers=[logging.FileHandler(log_file, encoding="utf-8"),
                  logging.StreamHandler(sys.stderr)],
    )


def parse_args():
    p = argparse.ArgumentParser(
        description="Daily Netskope GenAI app-events export for Power BI.")
    p.add_argument("--tenant", default=os.environ.get("NETSKOPE_TENANT"))
    p.add_argument("--token", default=os.environ.get("NETSKOPE_API_TOKEN"))
    p.add_argument("--output-dir",
                   default=os.environ.get("NETSKOPE_OUTPUT_DIR"),
                   help="Folder Power BI reads from.")
    p.add_argument("--category", default="generativeai",
                   help="Category filter (default: generativeai). '' disables.")
    p.add_argument("--filter", dest="filters", action="append", default=[],
                   help='Extra clause, e.g. "activity eq Upload". Repeatable.')
    p.add_argument("--fields", default=os.environ.get("NETSKOPE_FIELDS"),
                   help="Comma-separated columns to pin (recommended for Power BI).")
    p.add_argument("--file-prefix", default="netskope_genai_events",
                   help="Dated output filename prefix.")
    p.add_argument("--date",
                   help="Export this specific calendar day (YYYY-MM-DD) instead "
                        "of yesterday. Handy for backfilling a missed run.")
    p.add_argument("--utc", action="store_true",
                   help="Align day boundaries to UTC midnight instead of the "
                        "machine's local midnight.")
    p.add_argument("--limit", type=int, default=DEFAULT_LIMIT)
    return p.parse_args()


def compute_day_window(date_str: str | None, use_utc: bool):
    """
    Return (start_epoch, end_epoch, day_date) for one calendar day.

    Default is yesterday: [yesterday 00:00, today 00:00).
    Boundaries are local-midnight unless use_utc is set. day_date is the date
    the window represents, used to name the output file.
    """
    tz = timezone.utc if use_utc else None  # None -> system local time
    now = datetime.now(tz)

    if date_str:
        day_date = datetime.strptime(date_str, "%Y-%m-%d").date()
    else:
        day_date = (now - timedelta(days=1)).date()

    start_dt = datetime(day_date.year, day_date.month, day_date.day, tzinfo=tz)
    end_dt = start_dt + timedelta(days=1)
    return int(start_dt.timestamp()), int(end_dt.timestamp()), day_date


def main():
    args = parse_args()
    if not args.tenant or not args.token or not args.output_dir:
        sys.stderr.write("[error] Need --tenant, --token, and --output-dir "
                         "(or the matching env vars).\n")
        return 2

    os.makedirs(args.output_dir, exist_ok=True)
    setup_logging(args.output_dir)

    # One full calendar day (yesterday by default), in epoch seconds.
    starttime, endtime, day_date = compute_day_window(args.date, args.utc)

    base_url = build_base_url(args.tenant)
    query = build_query(args.category or None, args.filters)
    pinned = [c.strip() for c in args.fields.split(",")] if args.fields else None

    tz_label = "UTC" if args.utc else "local"
    logger.info("exporting calendar day %s (%s midnight boundaries); "
                "epoch %d -> %d", day_date.isoformat(), tz_label,
                starttime, endtime)
    logger.info("query: %s", query or "(none)")

    try:
        records = fetch_all(base_url, args.token, query,
                            starttime, endtime, args.limit)
    except Exception as exc:  # noqa: BLE001 - top-level guard for the scheduler
        logger.error("export failed: %s", exc)
        return 1

    columns = resolve_columns(records, pinned)
    out_name = f"{args.file_prefix}_{day_date:%Y-%m-%d}.csv"
    out_path = os.path.join(args.output_dir, out_name)
    write_csv_atomic(records, out_path, columns)

    logger.info("wrote %d events -> %s (%d columns)",
                len(records), out_path, len(columns))
    if not pinned:
        logger.warning("No --fields pinned: column set may vary between days, "
                       "which can break the Power BI folder combine. Pin --fields.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
