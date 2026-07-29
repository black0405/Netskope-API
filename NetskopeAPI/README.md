# Netskope GenAI Application Export

Pulls the previous day's Generative AI application events from the Netskope
REST API v2, writes them to CSV, and uploads the file to a SharePoint
document library. Built to run unattended on a schedule.

## What it does

- Queries `/api/v2/events/datasearch/application` filtered to
  `appcategory in ['Generative AI']`
- Pulls a full calendar day (00:00 → 00:00 local), subdividing into time
  windows so pagination stays correct on high-volume days
- Deduplicates to one row per user + application, summing file sizes
  across the collapsed events
- Exports a fixed column set with friendly headers
- Uploads to SharePoint via Microsoft Graph (client credentials flow)
- Logs every run, rotates daily, archives after 20 days

## Output columns

| Column | Source field |
|---|---|
| *(row number)* | generated |
| User | `user` |
| Application | `app` |
| URL | `url` |
| Activity | `activity` |
| Object Type | `object_type` |
| Object Name | `object` |
| Event Date | `timestamp` (formatted) |
| Organization Unit | `organization_unit` |
| Sum - File Size (MB) | `file_size` (summed, bytes → MB) |

## Requirements

- Python 3.8+
- `pip install requests pandas msal`

`msal` is only needed if SharePoint upload is enabled.

## Setup

```bash
git clone https://github.com/black0405/Netskope-API.git
cd Netskope-API

python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt

cp config.env.example config.env
# edit config.env with your tenant URL, API token, and SPN details
chmod 600 config.env
```

## Usage

```bash
# Export yesterday (the default, and what the scheduler runs)
python export_genai_applications.py

# Export a specific day
python export_genai_applications.py 2026-07-19

# Verify a day's export and upload happened; retries a failed upload
python export_genai_applications.py --check

# Print the one-time SharePoint site-grant command for an admin
python export_genai_applications.py --grant-help
```

## Exit codes

| Code | Meaning |
|---|---|
| 0 | Success, or nothing to do |
| 1 | Config / input error |
| 2 | Netskope API failure — no data collected |
| 3 | Ran fine, but zero records matched |
| 4 | Exported locally, but SharePoint upload failed |

The 2 vs 4 split is the useful one for alerting: **2 means no data**,
**4 means the data is on disk and only the upload needs retrying**.

## Scheduling

Export in the morning, verify in the afternoon:

```cron
30  2  * * *  /path/to/venv/bin/python /path/to/export_genai_applications.py
30 14  * * *  /path/to/venv/bin/python /path/to/export_genai_applications.py --check
```

Relative paths in `config.env` anchor to the script's own folder, so no
`cd` is needed. See [UAT_DEPLOYMENT.md](UAT_DEPLOYMENT.md) for Task
Scheduler, systemd, alerting, and backfill instructions.

## Configuration

All settings live in `config.env` — see `config.env.example` for the full
annotated list. Real environment variables override the file, so a
scheduler or vault can inject secrets without editing anything.

Point at a different config with:

```bash
NETSKOPE_CONFIG_FILE=/etc/netskope/prod.env python export_genai_applications.py
```

## Security

- **`config.env` is gitignored and must stay that way** — it holds the
  Netskope API token and the Entra client secret
- Prefer environment variables or a secret store over the file on
  production hosts
- The SPN needs Microsoft Graph **`Sites.Selected`** (application
  permission) *plus* a site-level grant on the target site. The API
  permission alone grants nothing — run `--grant-help` for the command
- Client secrets expire; when they do the export keeps working and only
  the upload fails (exit 4). Worth a calendar reminder

## Notes

- Exported CSVs contain user activity data. `output/` and `*.csv` are
  gitignored deliberately — don't commit extracts
- Day boundaries use the host's local timezone
- Run the export at 02:00–03:00, not midnight; Netskope events arrive with
  ingestion lag and a midnight run can under-count
