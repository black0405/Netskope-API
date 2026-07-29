# Netskope GenAI Export — UAT Deployment

## 1. Files to deploy

```
netskope-export/
├── export_genai_applications.py
├── config.env           <- created from config.env.example, NOT in source control
├── output/              <- created automatically
└── logs/
    ├── archive/         <- created automatically
    └── status/          <- one JSON per exported day, used by --check
```

## 2. Install

```bash
python -m venv venv
# Linux
source venv/bin/activate
# Windows
venv\Scripts\activate

pip install requests pandas msal
```

A venv matters more than usual here: it pins the versions so a later
system-wide `pip install` on the UAT box can't silently change behaviour.

## 3. Configure

```bash
cp config.env.example config.env
```

The script looks for `config.env` next to itself. It also accepts
`api_tokens.env`, `config.ini`, or `.env` if you prefer one of those, and
`NETSKOPE_CONFIG_FILE=/path/to/file` overrides the lot — which is how you
point the same script at UAT and PROD configs.

Fill in, at minimum:

| Key | Notes |
|---|---|
| `NETSKOPE_BASE_URL` | UAT tenant URL |
| `NETSKOPE_API_TOKEN` | REST API v2 token |
| `SPN_TENANT_ID` / `SPN_CLIENT_ID` / `SPN_CLIENT_SECRET` | the SPN |
| `SHAREPOINT_SITE_HOSTNAME` / `SHAREPOINT_SITE_PATH` | target site |
| `OUTPUT_DIR` / `LOG_DIR` | **use absolute paths** for scheduled runs |

**Paths.** Relative paths (`./output`, `./logs`) are anchored to the
*script's own folder*, not the working directory, so they behave correctly
under cron even without a `cd`. Absolute paths are used as-is if you'd
rather point the output somewhere specific like `/var/data/netskope`.

Lock the file down:

```bash
chmod 600 config.env              # Linux
icacls config.env /inheritance:r /grant:r "DOMAIN\svc-netskope:(R)"   # Windows
```

## 4. Smoke test before scheduling

Run these in order — each catches a different failure class.

```bash
# 1. Config loads, one specific day, upload off
#    (set SHAREPOINT_ENABLED=false in config.env first)
python export_genai_applications.py 2026-07-20
```
Expect `RUN END status=SUCCESS exit=0` and a CSV in `output/`.

```bash
# 2. Turn SHAREPOINT_ENABLED=true, run again for a different day
python export_genai_applications.py 2026-07-21
```
If this fails with a `Sites.Selected` message, the SPN's site grant hasn't
been applied. Get the exact admin command with:

```bash
python export_genai_applications.py --grant-help
```

```bash
# 3. Default path (yesterday), exactly as the scheduler will call it
python export_genai_applications.py

# 4. Verification pass -- should report the day above as already done
python export_genai_applications.py --check
```

Check the exit code explicitly — this is what the scheduler keys on:

```bash
echo $?          # Linux
echo %ERRORLEVEL%   # Windows cmd
$LASTEXITCODE    # PowerShell
```

## 5. Schedule it — export once, check once

Two scheduled entries, doing different jobs:

| Time | Command | What it does |
|---|---|---|
| 02:30 | `python export_genai_applications.py` | Full export + upload |
| 14:30 | `python export_genai_applications.py --check` | Verify only — no extraction |

`--check` does **not** re-query Netskope. It reads the day's status file and
confirms two things: the CSV exists, and it reached SharePoint. If the CSV
is on disk but the upload failed that morning, **it retries the upload** —
which is the whole reason the second run earns its slot. If the export never
happened at all, it logs the original error and the exact backfill command.

Do not run the export at 00:05. Netskope events arrive with ingestion lag,
so a midnight run can quietly under-count yesterday.

### Windows — Task Scheduler

Two tasks, same program, different arguments.

```cmd
schtasks /create /tn "Netskope GenAI Export" ^
  /tr "C:\netskope-export\venv\Scripts\python.exe C:\netskope-export\export_genai_applications.py" ^
  /sc daily /st 02:30 /ru "DOMAIN\svc-netskope" /rp

schtasks /create /tn "Netskope GenAI Check" ^
  /tr "C:\netskope-export\venv\Scripts\python.exe C:\netskope-export\export_genai_applications.py --check" ^
  /sc daily /st 14:30 /ru "DOMAIN\svc-netskope" /rp
```

In the GUI, set **Start in** to `C:\netskope-export` on both.

### Linux — cron

Edit with `crontab -e` as the service account (**not** root, unless the
files are owned by root):

```cron
# m  h  dom mon dow  command
30  2  *   *   *    /opt/netskope-export/venv/bin/python /opt/netskope-export/export_genai_applications.py >> /opt/netskope-export/logs/cron.out 2>&1
30 14  *   *   *    /opt/netskope-export/venv/bin/python /opt/netskope-export/export_genai_applications.py --check >> /opt/netskope-export/logs/cron.out 2>&1
```

Note there is no `cd` — the script anchors its own paths, so it works from
whatever directory cron happens to start in.

**Cron gotchas worth knowing:**

- **Use the full interpreter path.** Cron's `PATH` is minimal (often just
  `/usr/bin:/bin`), so a bare `python` may not resolve and the venv
  definitely won't be active. Point straight at
  `/opt/netskope-export/venv/bin/python`.
- **`%` must be escaped** in crontab (`\%`). Not an issue for these two
  lines, but it bites the moment you add a `date +%F` to the command.
- **Cron mails output to the local user** by default. The `>>` redirect
  above prevents a mailbox filling up on a box nobody reads.
- **Timezone** is the system's. Check with `timedatectl` — if the server is
  UTC but you're reasoning in IST, 02:30 isn't the hour you think it is,
  and "yesterday" shifts with it.

Verify the entries were accepted:

```bash
crontab -l
grep CRON /var/log/syslog | tail      # Debian/Ubuntu
journalctl -u cron --since today      # systemd hosts
```

The script writes its own structured log; `cron.out` only catches things
that die before logging starts (a syntax error, a missing interpreter).

### Linux — systemd timers (preferred if available)

Two service + timer pairs. The export:

`/etc/systemd/system/netskope-export.service`
```ini
[Unit]
Description=Netskope GenAI Export

[Service]
Type=oneshot
User=svc-netskope
WorkingDirectory=/opt/netskope-export
ExecStart=/opt/netskope-export/venv/bin/python export_genai_applications.py
```

`/etc/systemd/system/netskope-export.timer`
```ini
[Unit]
Description=Netskope GenAI Export (daily)

[Timer]
OnCalendar=*-*-* 02:30:00
Persistent=true

[Install]
WantedBy=timers.target
```

The check is the same with `--check` appended to `ExecStart` and
`OnCalendar=*-*-* 14:30:00`.

```bash
sudo systemctl enable --now netskope-export.timer netskope-check.timer
systemctl list-timers 'netskope-*'
```

`Persistent=true` is the advantage here: if the box was down at 02:30, the
run fires as soon as it comes back. Cron and Task Scheduler just skip it.

## 6. Exit codes — what to alert on

| Code | Meaning | Action |
|---|---|---|
| 0 | Success, or already done | none |
| 1 | Config / input error | check `config.env` |
| 2 | **Netskope failed** | no data collected — check token, tenant, network |
| 3 | Ran fine, zero records | usually genuine; check filter if it repeats |
| 4 | **SharePoint upload failed** | CSV is safe locally; fix perms and re-run |

The 2 vs 4 split is the useful one: **2 means you have no data**, **4 means
you have the data and only the upload needs retrying**.

`--check` returns the same codes, so one alerting rule covers both runs:

| Check result | Code | Meaning |
|---|---|---|
| Export + upload both fine | 0 | nothing to do |
| Upload had failed, retry worked | 0 | self-healed |
| Upload had failed, retry failed again | 4 | needs attention |
| No export file for the day | 2 | morning run never produced data |

Status files live in `logs/status/run_YYYY-MM-DD.json` — one per day,
recording rows exported, upload state, and the last error.

## 7. Reading the logs

One file per day, `logs/netskope_export.log`, rotated at midnight. Both
daily runs land in the same file, each tagged with a run id like `[023015]`.

```bash
# Did anything fail today?
grep "RUN END" logs/netskope_export.log

# Only failures, across all retained logs
grep -h "FAILED\|status=NETSKOPE\|status=SHAREPOINT" logs/netskope_export.log*

# Follow a single run end to end
grep "\[023015\]" logs/netskope_export.log
```

Retention is automatic: logs older than `LOG_RETENTION_DAYS` (20) get zipped
into `logs/archive/logs_YYYY-MM.zip` and removed. Archives older than
`ARCHIVE_RETENTION_DAYS` (180) are deleted. Set that to `0` to keep archives
forever.

## 8. Simple failure alerting

The script doesn't email — one line of wrapper does it, and keeps SMTP
config out of the tool.

**Windows** — attach a second Task Scheduler action, or wrap in PowerShell:

```powershell
& "C:\netskope-export\venv\Scripts\python.exe" "C:\netskope-export\export_genai_applications.py"
if ($LASTEXITCODE -ne 0) {
    Send-MailMessage -To "you@example.com" -From "netskope-uat@example.com" `
      -Subject "Netskope export FAILED (exit $LASTEXITCODE)" `
      -Body (Get-Content "C:\netskope-export\logs\netskope_export.log" -Tail 40 | Out-String) `
      -SmtpServer "smtp.example.com"
}
```

**Linux:**

```bash
#!/usr/bin/env bash
cd /opt/netskope-export || exit 1
./venv/bin/python export_genai_applications.py
code=$?
if [ $code -ne 0 ]; then
    tail -40 logs/netskope_export.log | \
      mail -s "Netskope export FAILED (exit $code)" you@example.com
fi
exit $code
```

Consider treating exit 3 (no records) as a warning rather than a page — a
genuinely quiet day is possible, though several in a row is worth a look.

## 9. Backfilling a missed day

A missed run is never auto-filled. To catch up:

```bash
python export_genai_applications.py 2026-07-19
```

`SKIP_IF_EXISTS=true` makes it safe to loop over a range — days already
done are skipped instantly.

```bash
for d in 2026-07-15 2026-07-16 2026-07-17; do
    ./venv/bin/python export_genai_applications.py "$d"
done
```

## 10. Promoting UAT → production

Because all configuration is in `config.env`, the promotion is: copy the
same `.py`, swap the `config.env`. Nothing in the code changes.

Worth verifying at that point:
- production Netskope token (a UAT token usually won't work)
- production SharePoint site, and its **own** `Sites.Selected` site grant —
  the UAT grant does not carry over
- absolute paths updated for the prod host
- client secret expiry date recorded somewhere with a reminder; when it
  lapses the export keeps working and only the upload starts failing
  (exit 4) — easy to miss without alerting
