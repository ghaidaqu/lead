# lead project

This folder is the canonical copy of the lead project.

## Structure
- `scripts/build_report.py` - builds the cleaned Excel report and dashboard workbook
- `scripts/update_statement.py` - rebuilds the bank-statement sheet
- `scripts/theme_workbook.py` - applies the workbook theme
- `web/generate_site.py` - builds the local HTML dashboard
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

## Runtime
- `app.py` serves the dashboard from `web/index.html` when the generated file exists.
- `GET /report.xlsx` downloads the current workbook when it exists.
- `requirements.txt` lists the runtime packages for local runs and deployment.
