# Lead Dashboard — Architecture

A financial dashboard for the **owner of lead-sa.com** (a shipping aggregator).
It scrapes the owner's own Lead admin panel, stores the data in PostgreSQL, and
serves an Arabic RTL single-page dashboard with revenue / cost / profit, carrier
and merchant breakdowns, COD tracking, and wallet/finance views.

> **The defining design decision:** profit/cost numbers come from **Lead's own
> financial reports** (`reports.php`), *not* from a maintained price sheet. Lead
> already computes the real carrier cost and net profit per period; we read those
> as the source of truth. See [§5 Pricing](#5-pricing-actuals-not-a-price-sheet).

---

## 1. Topology

```
                       ┌─────────────────────────────────────────────┐
                       │                 Railway                      │
                       │                                              │
  lead-sa.com  ◀───────┤  lead-worker (scraper loop)                  │
  (admin panel)        │     scripts/lead_worker.py                   │
        ▲              │       └─ scripts/sync_from_lead.py           │
        │ scrape       │                  │ upsert                    │
        │              │                  ▼                           │
        │              │            Postgres ──────────┐              │
        │              │                  ▲             │ read        │
        │              │           read   │             ▼             │
        └──────────────┤  lead (web)  app.py ── /api/* ── web/dashboard│
           (live           │     gunicorn        JSON      static SPA  │
            reports        └─────────────────────────────────────────┘
            via worker)
```

Three Railway services:

| Service       | Process                         | Role                                              |
|---------------|---------------------------------|---------------------------------------------------|
| `lead`        | `gunicorn app:app`              | Web: read-only JSON API + serves the static SPA   |
| `lead-worker` | `python3 scripts/lead_worker.py`| Scrapes Lead on a loop, upserts to Postgres       |
| `Postgres`    | managed                         | System of record                                  |

`Procfile` defines both process types.

### Deployment note (important)

Railway **auto-deploys from the GitHub repo** (`ghaidaqu/lead.git`). Manual
`railway up` deploys are **ephemeral** — any GitHub push *or env-var change*
triggers a rebuild from GitHub `HEAD` and reverts an ephemeral deploy. To make a
change permanent, commit it to GitHub. (During this rework, changes were deployed
ephemerally via `railway up` and **not committed**.)

---

## 2. Data flow

1. **Scrape** (`sync_from_lead.py`) — log in to Lead, fetch each admin page. List
   pages (`shipments.php`, `wallet.php`) are server-capped at 100 rows, so they're
   **date-window-sliced** (month → week → day, dedup by id) to capture *every* row
   (`window_scrape`). Pages are parsed by **header-matched** column remapping
   (`remap_table` + `*_SPEC`), so the parser survives column reordering.
2. **Pure parsers** turn HTML into typed rows (`extract_shipments`, `extract_wallet`,
   `extract_cod`, `extract_shipping_companies`, `extract_reports`, …). No IO inside
   parsers — they take HTML, return data.
3. **Idempotent upsert** into Postgres on natural keys (`db_store.upsert_rows`), so
   re-running a sync never duplicates.
4. **Actuals capture** — the worker also fetches Lead's `reports.php` for a set of
   date ranges and stores the parsed result in the `lead_reports` setting (see §5).
5. **Serve** — `app.py` exposes read-only `/api/*` endpoints that aggregate from
   Postgres (+ the captured actuals) and a static SPA that renders them.

---

## 3. Components

### `app.py` — web service
Flask app: read-only JSON API + static file serving. Key routes:

| Route | Purpose |
|---|---|
| `GET /api/dashboard?from=&to=&preset=` | The whole dashboard payload for a range |
| `POST /api/login` · `POST /api/logout` · `GET /api/session` | Dashboard auth (separate from Lead creds) |
| `GET /api/pricing` · `PUT /api/pricing/settings` · `POST/PUT/DELETE /api/pricing/merchants` | Legacy price-sheet admin (being retired) |
| `GET /report.xlsx` | On-demand in-memory Excel export (no file on disk) |
| `GET /dashboard/…`, `GET /admin` | Serve the SPA |
| `GET /health` | Liveness + data-source status |

- **Auth is fail-closed**: if `LEAD_AUTH_USER`/`LEAD_AUTH_PASS` are unset, all
  logins are denied (no "any credentials" hole). All write endpoints are gated
  behind `session["lead_authenticated"]` via `require_auth`.
- `get_lead_reports(date_from, date_to, preset)` reads the worker-captured
  `lead_reports` setting and returns Lead's actuals for the selected range.

### `scripts/sync_from_lead.py` — scraper / ETL (the core)
Stateless parsers + the sync `main()`. Notable pieces:
- `window_scrape(...)` — complete scrape past the 100-row cap.
- `*_SPEC` + `remap_table` / `_header_key` — header-based column mapping.
- `extract_reports(html)` — parses Lead's financial report into
  `{totals, carriers, merchants}` (**Lead's actual cost/profit**).
- `extract_cod_fee(html)` — the COD fee from `cod-settings.php`.
- `fetch_reports(opener, base, start, end)` + `_report_ranges(today)` — capture
  Lead's actuals per labelled range (see §5).
- `make_logged_in_opener(env)` — reusable authenticated session.
- `extract_site_kpis` — captures Lead's headline KPI cards (`site_kpis` setting).
- HTTP layer (`http_get_meta`, retries/backoff, `safe_http_get`).

### `scripts/lead_worker.py` — the loop
Runs `sync_from_lead.py` immediately on boot, then every
`LEAD_WORKER_INTERVAL_SECONDS` (default 3600s).

### `scripts/db_store.py` — persistence
`SCHEMA_SQL` + helpers. Tables:
`shipments`, `wallet_transactions`, `payments`, `cod_collections`, `sync_runs`,
`settings` (key/value text — holds `site_kpis`, `lead_reports`, `min_sync_date`),
plus the **legacy price sheet**: `carriers`, `merchant_overrides`,
`pricing_settings` (and the unused `price_rules`).
Generic `upsert_rows`, `get_setting`/`set_setting`, `load_pricing_snapshot`,
`prune_missing_shipments`.

### `scripts/pricing.py` — legacy profit engine
Pure functions (`compute_profit`, `PricingSnapshot`) that derived per-shipment
profit from the price sheet. **Being retired** in favour of Lead's actuals; still
used to populate the per-shipment profit columns that some sub-views read.

### `web/dashboard/index.html` — the SPA
Single self-contained file: Arabic RTL, Chart.js, dark mode, client-side routing
(dashboard ↔ pricing admin, no reload). Date-range presets (`الكل` / `هذا الشهر`
/ `آخر 30 يوم`) + custom range, COD show-more, per-panel methodology "i" buttons,
and a strip that surfaces Lead's actual profit vs the old sheet figure.

### `web/generate_site.py` — DB → dashboard payload
`load_data_from_db(date_from, date_to)` aggregates Postgres rows into the
dashboard shape (totals, top carriers/merchants/cities, daily series, finance
buckets, COD list, statement, missing-sequence numbers).

### `scripts/validate_totals.py`
Logs into Lead and diffs our DB aggregates against the live KPIs (counts must
match exactly; amounts allow rounding). Used to prove parity.

---

## 4. Authentication (two separate logins)

1. **Dashboard login** — `LEAD_AUTH_USER`/`LEAD_AUTH_PASS`, gates the web UI.
2. **Lead site login** — `LEAD_USERNAME`/`LEAD_PASSWORD`/`LEAD_BASE_URL`, used by
   the worker to scrape. Login flow: `GET /login.php` for the form + CSRF, then
   POST credentials, keep the `PHPSESSID` cookie (cached in `.auth/lead_state.json`).

> **Environment gotcha:** Lead's login succeeds from Railway's network but is
> rejected (`error=unauthorized`) from some other IPs, and a plain `urllib` login
> can fail where a real browser succeeds. Because of this, **only the worker holds
> a reliable Lead session** — the web service does *not* log into Lead. That's why
> the actuals are captured by the worker and read from the DB by the web app
> (§5), rather than fetched live in the web request path.

---

## 5. Pricing: actuals, not a price sheet

### Why
The original project computed profit from a maintained `الاسعار` price sheet
(carrier rates, special-merchant overrides, COD/extra-weight fees, a `0.85`
gross→net fallback, a VAT rate). This **drifted from reality** (it inflated profit
~60%) and required manual upkeep.

Lead's own **`reports.php`** already publishes, for any date range
(`?period=custom&start_date=…&end_date=…`):
- totals: `إجمالي الإيرادات`, `الإيرادات المحققة`, `التكلفة الأساسية`, `صافي الأرباح`
  (and `net_profit = realized_revenue − base_cost`, exactly);
- a **per-carrier** table (revenue, **actual cost**, profit, margin);
- a **per-merchant** table (revenue, profit) — special-merchant pricing already
  baked in.

So the price sheet is redundant. Every sheet number maps to an actual:

| Sheet input | Replaced by |
|---|---|
| carrier customer price | the actual `التكلفة` charge on each shipment row |
| carrier platform cost | `reports.php` per-carrier **cost** |
| merchant overrides + `is_special` | `reports.php` per-merchant **profit** |
| COD fee (`5`) | scraped from `cod-settings.php` |
| COD cost (`3`), extra-weight (`2.5/1.5`, `15kg`) | already inside Lead's reported cost/profit |
| `net_fallback_factor=0.85`, `vat_rate=0.15` | obsolete (no gross→net estimation needed) |

### How it's wired (worker captures, web reads)
Because only the worker can log into Lead reliably (§4):

1. **Worker** (`sync_from_lead.main()`): for each labelled range from
   `_report_ranges(today)` — `all`, `month`, `30d`, and the last 6 closed months
   (`YYYY-MM`) — fetch `reports.php` and parse with `extract_reports`. Store all of
   them, plus the scraped `cod_fee`, into the **`lead_reports`** setting:
   ```json
   { "ranges": { "all": {...}, "month": {...}, "2026-05": {...}, ... },
     "cod_fee": 5.0, "captured_at": "…" }
   ```
   Labels are **semantic** (not raw dates) so a date rollover between capture and
   page-load can't cause a miss.
2. **Web** (`get_lead_reports`): maps the selected range → a label
   (`preset` param for tabs; month-boundary detection for custom ranges) and
   returns that snapshot. Attached to `/api/dashboard` as `payload["lead_reports"]`.
3. **SPA**: headline revenue/cost/profit use `lead_reports.totals`; a strip shows
   the old sheet figure for transparency during the cutover.

### Frozen history
`reports.php` returns a closed month's actuals unchanged regardless of current
rates (carrier costs are real, booked payments), so any past range is exact and
re-derivable. A **price change therefore only affects future shipments** — it
never retroactively rewrites history (unlike the sheet, which re-priced all rows).

---

## 6. Database (Postgres)

Key tables (schema in `db_store.py`):
- `shipments` — one row per shipment; raw scraped columns + (legacy) computed
  profit columns; `raw_payload` JSON mirror.
- `wallet_transactions` / `payments` — wallet ledger (deposits, شحن debits,
  deductions, refunds). Keyed on the transaction id.
- `cod_collections` — COD orders with collection/transfer dates.
- `settings` — key/value text: `site_kpis`, **`lead_reports`**, `min_sync_date`.
- `carriers` / `merchant_overrides` / `pricing_settings` — the **legacy** price
  sheet (still present; being retired).
- `sync_runs` — sync audit trail.

---

## 7. Environment

| Var | Used by | Purpose |
|---|---|---|
| `DATABASE_URL` / `DATABASE_PUBLIC_URL` | both | Postgres connection |
| `LEAD_USERNAME` / `LEAD_PASSWORD` / `LEAD_BASE_URL` | worker | Lead site login |
| `LEAD_AUTH_USER` / `LEAD_AUTH_PASS` | web | Dashboard login |
| `LEAD_SESSION_SECRET` | web | Flask session secret |
| `LEAD_WORKER_INTERVAL_SECONDS` | worker | Sync cadence (default 3600) |

`requirements.txt`: Flask, gunicorn, openpyxl (xlsx export only), psycopg[binary].

### Run locally
```bash
pip install -r requirements.txt
# web:
DATABASE_URL=… LEAD_AUTH_USER=… LEAD_AUTH_PASS=… gunicorn app:app
# one sync:
DATABASE_URL=… LEAD_USERNAME=… LEAD_PASSWORD=… LEAD_BASE_URL=https://lead-sa.com \
  python3 scripts/sync_from_lead.py
```

---

## 8. Status & remaining work

**Live and reconciling to Lead's books:** headline revenue/cost/profit and the
carrier/merchant breakdowns, for every preset and full-month custom range
(e.g. June actual net profit **1,076** vs the old sheet's 1,801).

**Remaining to fully retire the sheet:**
1. **Per-shipment cost allocation** — daily chart and per-city profit still read
   sheet-derived per-shipment values. Allocate each carrier's actual cost across
   its shipments by weight, balanced to also hit each merchant's actual total
   (proportional fitting across the merchant×carrier grid), so sub-views also tie
   out to Lead.
2. **Delete the sheet** — once nothing reads them, drop `pricing.py`'s role,
   the `carriers`/`merchant_overrides`/`pricing_settings` tables, and the pricing
   admin UI/endpoints.

**Known constraints:** `reports.php` totals are whole-riyal rounded; the COD fee
is read as the *current* setting (fine going forward; a historical fee change
before capture isn't recoverable). Excel has been fully removed from the data path
(export is in-memory only).
