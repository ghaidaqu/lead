# lead project

This folder is the canonical copy of the lead project.

## Structure
- `scripts/build_report.py` - builds the cleaned Excel report and dashboard workbook
- `scripts/auto_update.py` - watches the raw workbook, rebuilds outputs, commits, and pushes
- `scripts/update_statement.py` - rebuilds the bank-statement sheet
- `scripts/theme_workbook.py` - applies the workbook theme
- `web/generate_site.py` - builds the local HTML dashboard
- `scripts/sync_from_lead.py` - one-shot scraper that updates the raw workbook and regenerates outputs
- `scripts/lead_worker.py` - long-running worker that schedules the sync script
- `launchd/` - macOS LaunchAgent plist for automatic updates
- `lead6_rules.md` - project rules and calculation logic
- `lead6_playbook.md` - operating notes for future updates

## Data files
The raw workbook and generated outputs are intentionally kept out of version control:
- `source/lead6.xlsx`
- `output/lead6_report.xlsx`
- `lead6_host/`

This keeps the repository code-only and avoids sharing private source data.

## Rebuild
1. Run `python3 scripts/build_report.py`
2. Run `python3 web/generate_site.py`
3. Run `python3 app.py` to serve the dashboard locally on `http://localhost:8000`

## Local sync
- Run `python3 scripts/sync_from_lead.py` for a single sync pass
- Run `python3 scripts/lead_worker.py` to wait for the 01:30 start window and then sync every hour
- Optional overrides:
  - `LEAD_WORKER_START_HOUR=1`
  - `LEAD_WORKER_START_MINUTE=30`
  - `LEAD_WORKER_INTERVAL_SECONDS=3600`

## Automatic updates
- The watcher script is `python3 scripts/auto_update.py --watch`
- By default it watches the synced Google Drive copy of `lead6.xlsx` at `~/Library/CloudStorage/GoogleDrive-gf.smartas@gmail.com/My Drive/lead6.xlsx`
- When the source changes, it copies the workbook into `source/lead6.xlsx`, rebuilds the report and dashboard, syncs `lead6_host/`, commits the tracked outputs, and pushes to GitHub
- The LaunchAgent plist in `launchd/` is intended to keep the watcher running after login on macOS
- For hosted deployments, run a separate Railway worker service with `python3 scripts/lead_worker.py` so sync runs outside the laptop session.

## Railway
- `web: gunicorn app:app --bind 0.0.0.0:$PORT` serves the dashboard only
- `worker: python3 scripts/lead_worker.py` runs the cloud sync loop
- Set these Railway environment variables: `LEAD_USERNAME`, `LEAD_PASSWORD`, `LEAD_BASE_URL`
- Optional worker schedule overrides: `LEAD_WORKER_START_HOUR`, `LEAD_WORKER_START_MINUTE`, `LEAD_WORKER_INTERVAL_SECONDS`

## Runtime
- `app.py` serves the dashboard from `web/index.html` when the generated file exists.
- `GET /report.xlsx` downloads the current workbook when it exists.
- `requirements.txt` lists the runtime packages for local runs and deployment.
