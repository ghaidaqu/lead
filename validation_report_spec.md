# Validation Report Specification

This document defines how PostgreSQL validation must be checked before the database becomes the authoritative source.

## Scope
Compare the legacy pipeline and the PostgreSQL-backed pipeline for the same input data.

## Required metrics
- Number of shipments
- Revenue
- Platform cost
- Overweight profit
- COD profit
- Total profit
- Top merchants
- Top cities
- Wallets
- Payments

## Required columns
- Current value
- New value
- Difference
- Difference percentage

## Acceptance rule
- If any metric differs by more than `0.01%`, PostgreSQL is not accepted as authoritative.
- The report must show the reason for each mismatch.

## Sample comparison set
- 20 random shipments must be compared row by row.
- For each shipment, include:
  - old calculation
  - new calculation
  - comparison result

## Suggested report output
- `metric_name`
- `current_value`
- `new_value`
- `difference`
- `difference_pct`
- `reason`

## Validation sources
- Old pipeline:
  - `scripts/build_report.py`
  - `web/generate_site.py`
- New pipeline:
  - PostgreSQL tables populated by `scripts/sync_from_lead.py`

## Notes
- PostgreSQL must not be considered final until the report shows full agreement within tolerance.
- Excel remains the fallback output until validation passes.

