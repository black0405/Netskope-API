"""
Migrate old hand-named GenAI activity CSVs to the format produced by
export_genai_applications.py.

    python migrate_old_exports.py [--dry-run]

Run it from inside the folder containing the old files.

Handles filenames like:
    GENAI apps activity 10 march.csv
    GENAI apps activity 14-19jan.csv
    GENAI_apps_activity_2026-03-05T0601.csv

Each becomes GenerativeAI_Applications_YYYY-MM-DD.csv (first day of a range;
year defaults to 2026 when the name has none), and the columns are reshaped
to the export layout:
    blank first column      -> S.No.  (values renumbered 1..n; added if absent)
    Event Data              -> Event Date
    Sum - File Size(MB)     -> Sum - File Size (MB)
    extra columns (e.g. department) are dropped
A file is skipped only if one of the export columns is missing entirely.

.xlsx files are handled too (first worksheet); their sheet is renamed to
EXCEL_SHEET_NAME to match the export script's output.
"""

import csv
import re
import sys
from pathlib import Path

DEFAULT_YEAR = 2026

# Same worksheet name export_genai_applications.py uses for .xlsx output.
EXCEL_SHEET_NAME = "in"

TARGET_HEADERS = ["S.No.", "User", "Application", "URL", "Activity",
                  "Object Type", "Object Name", "Event Date",
                  "Organization Unit", "Sum - File Size (MB)"]

MONTHS = {"jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
          "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12}


def parse_date(name: str) -> str:
    """Pull a YYYY-MM-DD label out of an old filename, or raise ValueError."""
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", name)
    if m:
        return m.group(0)
    # "10 march" / "14-19jan" -> first day + month word, default year
    m = re.search(r"(\d{1,2})(?:\s*-\s*\d{1,2})?[\s_-]*([a-zA-Z]+)", name)
    if m:
        month = MONTHS.get(m.group(2).lower()[:3])
        if month:
            return f"{DEFAULT_YEAR}-{month:02d}-{int(m.group(1)):02d}"
    raise ValueError(f"no date found in {name!r}")


def fix_header(old: list) -> list:
    """Map old header variants onto the canonical export headers."""
    canon = {re.sub(r"[^a-z]", "", h.lower()): h for h in TARGET_HEADERS}
    canon[re.sub(r"[^a-z]", "", "Event Data".lower())] = "Event Date"
    out = []
    for i, h in enumerate(old):
        if i == 0 and not h.strip():
            out.append("S.No.")
        else:
            out.append(canon.get(re.sub(r"[^a-z]", "", h.lower()), h))
    return out


def read_rows(path: Path) -> list:
    if path.suffix.lower() == ".xlsx":
        from openpyxl import load_workbook
        ws = load_workbook(path).worksheets[0]
        return [["" if c is None else c for c in row]
                for row in ws.iter_rows(values_only=True)]
    with open(path, newline="", encoding="utf-8-sig") as f:
        return list(csv.reader(f))


def write_rows(path: Path, rows: list) -> None:
    if path.suffix.lower() == ".xlsx":
        from openpyxl import Workbook
        wb = Workbook()
        wb.active.title = EXCEL_SHEET_NAME
        for row in rows:
            wb.active.append(row)
        wb.save(path)
        return
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)


def migrate(folder: Path, dry_run: bool) -> None:
    old_files = [p for ext in ("*.csv", "*.xlsx") for p in folder.glob(ext)
                 if re.match(r"GENAI[ _]apps[ _]activity", p.name, re.I)]
    if not old_files:
        print(f"No 'GENAI apps activity' files found in {folder}")
        return
    for path in sorted(old_files):
        try:
            label = parse_date(path.stem)
        except ValueError as exc:
            print(f"SKIP {path.name}: {exc}")
            continue
        rows = read_rows(path)
        fixed = fix_header([str(h) for h in rows[0]]) if rows else []
        missing = [h for h in TARGET_HEADERS[1:] if h not in fixed]
        if missing:
            print(f"SKIP {path.name}: missing columns {missing}")
            continue
        ext = path.suffix.lower()
        dest = folder / f"GenerativeAI_Applications_{label}{ext}"
        n = 2
        while dest.exists():                     # never clobber
            dest = folder / f"GenerativeAI_Applications_{label}_{n}{ext}"
            n += 1
        print(f"{path.name}  ->  {dest.name}")
        if dry_run:
            continue
        # Reshape every row to the export layout: fresh S.No. counter, then
        # the nine data columns in order; anything else (department...) drops.
        idx = [fixed.index(h) for h in TARGET_HEADERS[1:]]
        out = [TARGET_HEADERS] + [
            [i] + [row[j] if j < len(row) else "" for j in idx]
            for i, row in enumerate(rows[1:], 1) if row
        ]
        write_rows(dest, out)
        path.unlink()


def selftest() -> None:
    assert parse_date("GENAI apps activity 10 march") == "2026-03-10"
    assert parse_date("GENAI apps activity 14-19jan") == "2026-01-14"
    assert parse_date("GENAI_apps_activity_2026-03-05T0601") == "2026-03-05"
    assert fix_header(["", "User", "Event Data", "Sum - File Size(MB)"]) == \
        ["S.No.", "User", "Event Date", "Sum - File Size (MB)"]
    print("selftest OK")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if a != "--dry-run"]
    if "--selftest" in args:
        selftest()
        sys.exit(0)
    migrate(Path(args[0]) if args else Path.cwd(), dry_run="--dry-run" in sys.argv)
