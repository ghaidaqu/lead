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

The remaining problem is not the website scrape itself. The scrape is working.

The current risk area is keeping the reporting pipeline fully stable and aligned with the cloud architecture:

- `build_report.py` still depends on a valid price reference.
- The dashboard generation path still needs to be confirmed end-to-end after the latest PostgreSQL price snapshot changes.
- We need to verify that the price sheet stored in PostgreSQL becomes the stable source used by the reporting flow when the workbook does not contain the expected sheet.

## Latest issue encountered

The latest worker runs exposed two concrete lookup failures:

1. `build_report.py` raised a `KeyError` because it could not find the expected price column name:
   - `السعر لكل كيلو  زيادة`
   - This meant the reporting code was too strict about the exact header text.

2. `web/generate_site.py` raised a `KeyError` because the workbook did not contain the sheet:
   - `تفاصيل شهر 6`
   - The code was assuming a single hardcoded sheet name instead of accepting other valid detail-sheet names.

These are not scrape failures. They are report/dashboard lookup failures caused by rigid workbook expectations.

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

- the report builder is still expecting the price sheet in the workbook at runtime,
- the PostgreSQL price snapshot needs to be wired into the report builder as the fallback reference,
- or the dashboard generator needs a small adjustment so it can rely on the stored price rules consistently.

## Recommended next step

The next step should be a verification pass that checks:

- the latest worker run after the PostgreSQL price snapshot change,
- whether `build_report.py` succeeds,
- whether `web/generate_site.py` succeeds,
- whether the dashboard reflects the same calculations as before,
- and whether the price rules are being read from PostgreSQL when needed.

The lookup logic should also stay flexible enough to accept:

- multiple spellings of the extra-kilo price header,
- multiple acceptable names for the detail sheet,
- and PostgreSQL-backed price rules as the fallback reference when the workbook is incomplete.

## Bottom line

The ingestion side is now in good shape.  
The remaining work is to make the reporting layer fully dependable with PostgreSQL-backed price rules so the whole pipeline runs cleanly end to end.
