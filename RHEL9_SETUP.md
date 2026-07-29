# RHEL 9.6 Setup

Tested against RHEL 9.x defaults: Python 3.9, systemd, SELinux enforcing.

## 1. Python

RHEL 9 ships Python 3.9 as `/usr/bin/python3`. The script is written to run
on 3.8+, so the system Python is fine — no need for a newer build.

```bash
python3 --version          # expect 3.9.x
sudo dnf install -y python3-pip
```

If your estate standardises on a newer Python, RHEL 9 also carries 3.11 and
3.12 as parallel installs:

```bash
sudo dnf install -y python3.12 python3.12-pip
```

Use whichever you pick consistently in the venv and the systemd unit.

## 2. screen or tmux

**`screen` is not in the RHEL 9 base repositories** — it moved to EPEL.
`tmux` *is* in AppStream, so it's the lower-friction choice:

```bash
# tmux (recommended on RHEL 9 -- base repo, no EPEL needed)
sudo dnf install -y tmux
```

If you specifically want `screen`:

```bash
sudo dnf install -y https://dl.fedoraproject.org/pub/epel/epel-release-latest-9.noarch.rpm
sudo dnf install -y screen
```

The script auto-detects: it prefers `screen`, falls back to `tmux`, and runs
in the foreground if neither exists. Force one with `SCREEN_TOOL=tmux` in
`config.env`.

## 3. Service account and directories

```bash
sudo useradd -r -m -d /opt/netskope-export -s /sbin/nologin svc-netskope
sudo mkdir -p /opt/netskope-export/{output,logs}
```

Deploy the files, then:

```bash
cd /opt/netskope-export
sudo python3 -m venv venv
sudo ./venv/bin/pip install -r requirements.txt

sudo cp config.env.example config.env
sudo vi config.env                       # fill in tokens and site details

sudo chown -R svc-netskope:svc-netskope /opt/netskope-export
sudo chmod 600 /opt/netskope-export/config.env
sudo chmod 750 /opt/netskope-export
```

`chmod 600` on `config.env` matters — it holds the Netskope token and the
SPN client secret.

## 4. SELinux

SELinux is enforcing by default on RHEL 9. Two things can bite:

**Writing outside the home directory.** If you put `OUTPUT_DIR` somewhere
like `/var/data`, label it:

```bash
sudo semanage fcontext -a -t var_t "/var/data/netskope(/.*)?"
sudo restorecon -Rv /var/data/netskope
```

Keeping output under `/opt/netskope-export/output` avoids this entirely.

**Outbound HTTPS from a systemd service** is allowed by default. If your
build has a tighter policy, check for denials after the first run:

```bash
sudo ausearch -m avc -ts recent
sudo sealert -a /var/log/audit/audit.log      # needs setroubleshoot-server
```

Do not blanket-disable SELinux to fix this. If a denial appears, generate a
targeted policy:

```bash
sudo ausearch -c 'python3' --raw | audit2allow -M netskope-export
sudo semodule -i netskope-export.pp
```

## 5. Firewall / proxy

The script needs outbound 443 to two places:

| Host | Purpose |
|---|---|
| `<tenant>.goskope.com` | Netskope API |
| `graph.microsoft.com` and `login.microsoftonline.com` | SharePoint upload |

`firewalld` doesn't restrict outbound by default, so usually nothing to do.
Behind a corporate proxy, set it in the systemd unit rather than globally:

```ini
Environment=HTTPS_PROXY=http://proxy.corp.example.com:8080
Environment=NO_PROXY=localhost,127.0.0.1,.corp.example.com
```

## 6. First run — interactive

Run it by hand once. Because you're at a terminal, it relaunches itself
inside a session named **Netskope GenAI**:

```bash
cd /opt/netskope-export
sudo -u svc-netskope ./venv/bin/python export_genai_applications.py
```

```
Starting inside tmux session: Netskope GenAI
  detach   : Ctrl-B then D
  reattach : tmux attach -t "Netskope GenAI"
```

Detach and reattach freely — useful when the first extraction is large and
you don't want an SSH drop to kill it.

Skip the wrapper with `--no-screen`, or set `USE_SCREEN=false` in
`config.env`.

## 7. Scheduling — systemd (preferred on RHEL)

systemd is a better fit than cron here: `Persistent=true` catches up a run
that was missed while the box was down, and `journalctl` gives you the
run history alongside the script's own logs.

The screen wrapper **automatically disables itself** under systemd — there's
no terminal, so it detects that and runs directly.

`/etc/systemd/system/netskope-export.service`
```ini
[Unit]
Description=Netskope GenAI Export
After=network-online.target
Wants=network-online.target

[Service]
Type=oneshot
User=svc-netskope
Group=svc-netskope
WorkingDirectory=/opt/netskope-export
ExecStart=/opt/netskope-export/venv/bin/python /opt/netskope-export/export_genai_applications.py
StandardOutput=journal
StandardError=journal

# Hardening -- output and logs stay writable, nothing else
ProtectSystem=strict
ProtectHome=yes
PrivateTmp=yes
NoNewPrivileges=yes
ReadWritePaths=/opt/netskope-export/output /opt/netskope-export/logs

[Install]
WantedBy=multi-user.target
```

`/etc/systemd/system/netskope-export.timer`
```ini
[Unit]
Description=Netskope GenAI Export (daily 02:30)

[Timer]
OnCalendar=*-*-* 02:30:00
Persistent=true
RandomizedDelaySec=300

[Install]
WantedBy=timers.target
```

The verification pass — same pattern, `--check` appended:

`/etc/systemd/system/netskope-check.service`
```ini
[Unit]
Description=Netskope GenAI Export -- verification pass
After=network-online.target
Wants=network-online.target

[Service]
Type=oneshot
User=svc-netskope
Group=svc-netskope
WorkingDirectory=/opt/netskope-export
ExecStart=/opt/netskope-export/venv/bin/python /opt/netskope-export/export_genai_applications.py --check
ProtectSystem=strict
ProtectHome=yes
PrivateTmp=yes
NoNewPrivileges=yes
ReadWritePaths=/opt/netskope-export/output /opt/netskope-export/logs

[Install]
WantedBy=multi-user.target
```

`/etc/systemd/system/netskope-check.timer`
```ini
[Unit]
Description=Netskope GenAI verification (daily 14:30)

[Timer]
OnCalendar=*-*-* 14:30:00
Persistent=true

[Install]
WantedBy=timers.target
```

Enable both:

```bash
sudo systemctl daemon-reload
sudo systemctl enable --now netskope-export.timer netskope-check.timer
systemctl list-timers 'netskope-*'
```

Trigger a run immediately to confirm it works under systemd:

```bash
sudo systemctl start netskope-export.service
sudo journalctl -u netskope-export.service -n 50 --no-pager
```

## 8. Scheduling — cron (if systemd timers aren't an option)

```bash
sudo -u svc-netskope crontab -e
```

```cron
30  2  * * *  /opt/netskope-export/venv/bin/python /opt/netskope-export/export_genai_applications.py
30 14  * * *  /opt/netskope-export/venv/bin/python /opt/netskope-export/export_genai_applications.py --check
```

No `cd` needed — the script anchors its own paths. Use the full venv
interpreter path; cron's `PATH` is minimal.

## 9. Timezone

Day boundaries follow the host clock:

```bash
timedatectl
sudo timedatectl set-timezone Asia/Kolkata     # if it should be IST
```

Worth settling before go-live — "yesterday" shifts with it, and changing it
later makes one day's file overlap or gap against the previous ones.

## 10. Verify

```bash
# by hand, a specific day
sudo -u svc-netskope /opt/netskope-export/venv/bin/python \
  /opt/netskope-export/export_genai_applications.py 2026-07-20 --no-screen
echo "exit=$?"

# what the timers will do
systemctl list-timers 'netskope-*'

# logs
tail -f /opt/netskope-export/logs/netskope_export.log
grep "RUN END" /opt/netskope-export/logs/netskope_export.log
```

Exit codes: `0` success, `1` config, `2` Netskope failed, `3` no records,
`4` SharePoint upload failed.
