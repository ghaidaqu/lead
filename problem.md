# Current Problem Summary

This document captures the current state of the Lead automation pipeline and the main issue still being worked through.

## What is working

- The worker successfully logs into the Lead website in the cloud environment.
- The worker can open the dashboard, shipments, wallet, and COD pages.
- Data is being scraped from the website itself, not from Excel.
- The worker can collect:
  - shipments
  - wallet transactions
  - payments
  - COD collections
- PostgreSQL connectivity is working.
- The price sheet snapshot is now being stored in PostgreSQL as a persistent reference.
- The worker runs on a periodic schedule and performs an immediate sync on startup.

## What was broken before

Earlier runs failed for several different reasons:

1. `subprocess` was missing from `scripts/sync_from_lead.py`.
2. `build_report.py` could not find a required price sheet reference.
3. `web/generate_site.py` could not import `scripts.db_store` on Railway because the package was not initialized correctly.
4. The worker originally depended on `lead6.xlsx` too aggressively, which was not what we wanted for the cloud pipeline.

Those issues have been addressed one by one.

## Current problem

The remaining problem is no longer the website scrape itself. The scrape is working.

The current blocker is the PostgreSQL handoff inside the worker:

- `DATABASE_URL` is present in the worker environment.
- The worker now logs `postgres_enabled=true`.
- But the database connection step is still failing before any upsert happens.
- The latest failure path reported:
  - `postgres_connected=false`
  - `sync_run_id=null`
  - `error: "'NoneType' object is not callable"`

That tells us the scrape is healthy, but the DB helper path inside `scripts/sync_from_lead.py` is still not resolving the PostgreSQL connector correctly in the worker runtime.

## Latest issue encountered

The latest worker runs now show a different and more specific issue:

1. `build_report.py` is no longer the blocker.
   - The latest runs show `build_report_error=null`.

2. `web/generate_site.py` is also no longer the blocker.
   - The latest runs show `site_error=null`.

3. The worker can still open the Lead pages successfully.
   - dashboard
   - shipments
   - wallet
   - COD

4. The remaining failure is PostgreSQL connection wiring inside the worker.
   - The worker reports `postgres_enabled=true`
   - but `postgres_connected=false`
   - and `sync_run_id=null`
   - with the error:
     - `'NoneType' object is not callable`

That points to a code-path problem rather than a data problem.

## Important rule note

`لكل كيلو زيادة` is not just a sheet label or a loose column name. It is a real business rule used in the profit calculation for overweight shipments.

That means:

- the system must preserve the rule itself,
- the lookup can be flexible about the exact text in Excel,
- but the calculation logic must continue to treat the value as part of the official pricing rules.

So if the source text changes slightly, the code should still resolve it correctly without changing the actual math.

In short:

- Lead website access: working
- Scraping: working
- PostgreSQL storage: working
- Price sheet persistence: now stored in PostgreSQL
- Final report and dashboard generation: needs a clean verification pass after the latest changes

## Why this matters

The business rules and financial calculations must remain unchanged.

The system is intended to work like this:

1. Scrape the Lead website.
2. Store the data in PostgreSQL.
3. Keep the price sheet snapshot as a stable reference.
4. Generate the Excel export.
5. Build the dashboard from the validated data.

If the report generation cannot reliably find the price reference, the downstream dashboard and Excel output may fail even though the scrape itself is healthy.

## Current hypothesis

The most likely remaining issue is one of these:

- `scripts/sync_from_lead.py` is still importing the DB helpers in a way that can leave `get_conn` or a related helper unset in the worker runtime.
- The worker should not rely on a fragile top-level import for the PostgreSQL path.
- The DB import should be resolved lazily and explicitly inside the sync run so the worker either connects cleanly or reports a real import error.

## Recommended next step

The next step should be a focused verification pass that checks:

- the worker imports the PostgreSQL helper module directly,
- `get_conn()` is reached from the real module and not a fallback `None`,
- a `sync_run_id` is created,
- rows are upserted into PostgreSQL,
- and the one-line `SYNC_SUMMARY` shows `postgres_connected=true`.

The report/dashboard lookup logic still needs to remain flexible, but at the moment it is not the active blocker.

## Latest worker run update

The newest worker run confirmed that website scraping still works:

- dashboard opened successfully
- shipments opened successfully
- wallet opened successfully
- COD opened successfully

But the same two lookup failures are still blocking the reporting layer:

1. `build_report.py` still fails on the extra-kilo price header lookup:
   - `KeyError: 'السعر لكل كيلو  زيادة'`
   - The report builder is still too strict about the exact header text.

2. `web/generate_site.py` still fails when the detail sheet is not found:
   - `KeyError: 'Worksheet تفاصيل شهر 6 does not exist.'`
   - The dashboard fallback path is still expecting a hardcoded sheet name.

Current state:

- scrape: working
- login: working
- website pages: working
- PostgreSQL connectivity: working
- report generation: still blocked by price header lookup
- dashboard generation: still blocked by detail-sheet lookup

## Bottom line

The ingestion side is now in good shape.
The current work is to make the worker’s PostgreSQL connection path reliable end to end so the scrape results are actually persisted and logged in `sync_runs`.

## Current housekeeping

- Duplicate backup artifacts and stray system files have been removed from the repo workspace.
- The project now stays focused on the live pipeline files instead of generated duplicates.
- The remaining tracked problem is still the report/dashboard lookup behavior, not the scraping itself.
