# lead project

Scrapes lead-sa.com, computes per-shipment profit, and serves an Arabic RTL
analytics dashboard. Flask + PostgreSQL, deployed on Railway.

The pipeline is fully database-backed — **there are no Excel files**. Pricing
lives in PostgreSQL and the worker computes profit inline at sync time.

## Architecture

```
scrape (pure parsers) → pricing snapshot (DB) → pure profit engine → upsert (PostgreSQL)
                                                                          ↓
                       read-only JSON API (app.py) → dashboard + pricing admin page
```

## Code
- `scripts/sync_from_lead.py` — scrapes shipments/wallet/COD + carrier prices, computes profit, upserts to PostgreSQL
- `scripts/pricing.py` — pure, stateless profit engine (no IO); the single source of the profit math
- `scripts/db_store.py` — PostgreSQL access, schema, and the pricing model helpers
- `scripts/export_report.py` — builds the `/report.xlsx` download in memory from the DB (the only place openpyxl is used)
- `scripts/lead_worker.py` — long-running worker that runs the sync on a schedule
- `web/generate_site.py` — renders the dashboard HTML from DB data (`build_html` / `load_data_from_db`)
- `web/admin.html` — pricing admin UI (COD fee, extra-kilo, per-merchant overrides)
- `app.py` — Flask app: dashboard, auth, pricing API, on-demand xlsx

## Pricing model (PostgreSQL)
- `carriers` — per-carrier prices, **auto-synced** from `/admin/shipping-companies.php` each sync
- `merchant_overrides` — per-merchant negotiated prices (admin-editable)
- `pricing_settings` — COD fee, extra-kilo pricing, VAT (admin-editable)

The COD fee, extra-kilo pricing, and merchant overrides are not on the Lead
site, so they are edited in the web admin page (`/admin`, login required) and
persisted in the DB. Carrier prices are read-only there (auto-synced).

## Railway
- `web: gunicorn app:app --bind 0.0.0.0:$PORT` — dashboard + API
- `worker: python3 scripts/lead_worker.py` — cloud sync loop
- Required env vars: `LEAD_USERNAME`, `LEAD_PASSWORD`, `LEAD_BASE_URL`, `DATABASE_URL`
- Auth (required to log in / edit pricing): `LEAD_AUTH_USER`, `LEAD_AUTH_PASS`, `LEAD_SESSION_SECRET`
- Optional worker schedule: `LEAD_WORKER_INTERVAL_SECONDS`

## Local
- `DATABASE_URL=... python3 scripts/sync_from_lead.py` — one sync pass
- `DATABASE_URL=... python3 app.py` — serve on `http://localhost:8000`

## Runtime
- `GET /` → dashboard (rendered live from PostgreSQL)
- `GET /admin` → pricing admin (auth required)
- `GET /api/pricing` → pricing model JSON; `PUT /api/pricing/settings`, `*/api/pricing/merchants` → edits (auth required)
- `GET /report.xlsx` → Excel export generated on demand from the DB (no file stored)
- `GET /health` → data-source / Postgres status
