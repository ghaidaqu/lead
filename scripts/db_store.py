from __future__ import annotations

import datetime as dt
import json
import os
from contextlib import contextmanager
from pathlib import Path
from typing import Any

try:
    import psycopg
    from psycopg.rows import dict_row
    from psycopg.types.json import Jsonb
except Exception:  # pragma: no cover
    psycopg = None
    dict_row = None
    Jsonb = None


PROJECT_DIR = Path(__file__).resolve().parents[1]
def db_url() -> str | None:
    for key in ("DATABASE_URL", "POSTGRES_URL", "RAILWAY_DATABASE_URL"):
        value = os.environ.get(key, "").strip()
        if value:
            return value
    return None


def db_enabled() -> bool:
    return bool(db_url() and psycopg is not None)


def _json_default(value: Any):
    if isinstance(value, (dt.datetime, dt.date)):
        return value.isoformat()
    return str(value)


def _db_value(value: Any):
    if isinstance(value, (dt.datetime, dt.date)):
        return value
    if isinstance(value, (dict, list, tuple)):
        if Jsonb is not None:
            return Jsonb(value, dumps=lambda obj: json.dumps(obj, ensure_ascii=False, default=_json_default))
        return json.dumps(value, ensure_ascii=False, default=_json_default)
    return value


@contextmanager
def get_conn():
    url = db_url()
    if not url or psycopg is None:
        raise RuntimeError("DATABASE_URL is not configured or psycopg is unavailable.")
    conn = psycopg.connect(url, row_factory=dict_row)
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS shipments (
    id BIGSERIAL PRIMARY KEY,
    order_id TEXT NOT NULL UNIQUE,
    tracking_number TEXT,
    merchant_name TEXT,
    store_name TEXT,
    customer_name TEXT,
    city TEXT,
    carrier TEXT,
    payment_type TEXT,
    status TEXT,
    shipment_date DATE,
    delivery_date DATE,
    weight NUMERIC(10,2),
    cod_amount NUMERIC(12,2),
    shipping_charge NUMERIC(12,2),
    customer_price_gross NUMERIC(12,2),
    customer_price_net NUMERIC(12,2),
    platform_cost_gross NUMERIC(12,2),
    platform_cost_net NUMERIC(12,2),
    base_profit NUMERIC(12,2),
    extra_kg NUMERIC(10,2),
    extra_profit NUMERIC(12,2),
    cod_profit NUMERIC(12,2),
    total_profit NUMERIC(12,2),
    actual_base_cost NUMERIC(12,2),
    actual_extra_cost NUMERIC(12,2),
    actual_revenue NUMERIC(12,2),
    actual_profit NUMERIC(12,2),
    cost_source TEXT,
    included_in_profit BOOLEAN NOT NULL DEFAULT TRUE,
    source_row INTEGER,
    source_hash TEXT,
    raw_payload JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS wallet_transactions (
    id BIGSERIAL PRIMARY KEY,
    transaction_key TEXT NOT NULL UNIQUE,
    transaction_date TIMESTAMPTZ,
    user_name TEXT,
    description TEXT,
    amount NUMERIC(12,2),
    transaction_type TEXT,
    balance_before NUMERIC(12,2),
    balance_after NUMERIC(12,2),
    source_page TEXT,
    raw_payload JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS payments (
    id BIGSERIAL PRIMARY KEY,
    payment_key TEXT NOT NULL UNIQUE,
    payment_date TIMESTAMPTZ,
    customer_name TEXT,
    amount NUMERIC(12,2),
    method TEXT,
    status TEXT,
    source_page TEXT,
    raw_payload JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS cod_collections (
    id BIGSERIAL PRIMARY KEY,
    order_id TEXT NOT NULL UNIQUE,
    tracking_number TEXT,
    merchant_name TEXT,
    customer_name TEXT,
    cod_amount NUMERIC(12,2),
    collection_date DATE,
    transfer_date DATE,
    settlement_status TEXT,
    shipment_status TEXT,
    raw_payload JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS sync_runs (
    id BIGSERIAL PRIMARY KEY,
    started_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    finished_at TIMESTAMPTZ,
    status TEXT NOT NULL DEFAULT 'running',
    source TEXT,
    rows_inserted INTEGER NOT NULL DEFAULT 0,
    rows_updated INTEGER NOT NULL DEFAULT 0,
    rows_skipped INTEGER NOT NULL DEFAULT 0,
    error_message TEXT,
    log_path TEXT,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS price_rules (
    id BIGSERIAL PRIMARY KEY,
    merchant_name TEXT,
    carrier_name TEXT,
    customer_gross NUMERIC(12,2),
    customer_net NUMERIC(12,2),
    platform_gross NUMERIC(12,2),
    platform_net NUMERIC(12,2),
    tax_mode TEXT,
    extra_kilo_mode TEXT,
    cod_mode TEXT,
    raw_payload JSONB,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS carriers (
    id BIGSERIAL PRIMARY KEY,
    carrier_name TEXT NOT NULL UNIQUE,
    customer_gross NUMERIC(12,4),
    customer_net NUMERIC(12,4),
    platform_gross NUMERIC(12,4),
    platform_net NUMERIC(12,4),
    source TEXT NOT NULL DEFAULT 'scrape',
    active BOOLEAN NOT NULL DEFAULT TRUE,
    created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE TABLE IF NOT EXISTS pricing_settings (
    key TEXT PRIMARY KEY,
    value NUMERIC(12,4),
    updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW()
);

CREATE INDEX IF NOT EXISTS idx_shipments_status_date ON shipments(status, shipment_date);
CREATE INDEX IF NOT EXISTS idx_wallet_transactions_date ON wallet_transactions(transaction_date);
CREATE INDEX IF NOT EXISTS idx_payments_date ON payments(payment_date);
CREATE INDEX IF NOT EXISTS idx_cod_collections_dates ON cod_collections(collection_date, transfer_date);
"""


def ensure_schema(conn) -> None:
    with conn.cursor() as cur:
        cur.execute(SCHEMA_SQL)
        # Lightweight migrations for columns added to an already-created table.
        for ddl in (
            "ALTER TABLE shipments ADD COLUMN IF NOT EXISTS actual_base_cost NUMERIC(12,2)",
            "ALTER TABLE shipments ADD COLUMN IF NOT EXISTS actual_revenue NUMERIC(12,2)",
            "ALTER TABLE shipments ADD COLUMN IF NOT EXISTS actual_profit NUMERIC(12,2)",
            "ALTER TABLE shipments ADD COLUMN IF NOT EXISTS cost_source TEXT",
            "ALTER TABLE shipments ADD COLUMN IF NOT EXISTS actual_extra_cost NUMERIC(12,2)",
            # The manual price sheet is retired — profit comes from Lead's actuals.
            "DROP TABLE IF EXISTS merchant_overrides",
        ):
            cur.execute(ddl)


def upsert_rows(conn, table: str, rows: list[dict[str, Any]], key_columns: list[str], update_columns: list[str]) -> tuple[int, int]:
    if not rows:
        return 0, 0
    inserted = updated = 0
    cols = list(rows[0].keys())
    insert_cols = ", ".join(cols)
    placeholders = ", ".join(f"%({c})s" for c in cols)
    updates = ", ".join(f"{col}=EXCLUDED.{col}" for col in update_columns)
    conflict = ", ".join(key_columns)
    sql = f"INSERT INTO {table} ({insert_cols}) VALUES ({placeholders}) ON CONFLICT ({conflict}) DO UPDATE SET {updates}, updated_at=NOW()"
    with conn.cursor() as cur:
        for row in rows:
            cur.execute(sql, {key: _db_value(value) for key, value in row.items()})
            if cur.rowcount == 1:
                inserted += 1
            else:
                updated += 1
    return inserted, updated


def compare_counts(conn) -> dict[str, Any]:
    with conn.cursor() as cur:
        cur.execute("SELECT COUNT(*) AS count FROM shipments")
        shipments = int(cur.fetchone()["count"])
        cur.execute("SELECT COUNT(*) AS count FROM wallet_transactions")
        wallet = int(cur.fetchone()["count"])
        cur.execute("SELECT COUNT(*) AS count FROM payments")
        payments = int(cur.fetchone()["count"])
        cur.execute("SELECT COUNT(*) AS count FROM cod_collections")
        cod = int(cur.fetchone()["count"])
    return {"shipments": shipments, "wallet_transactions": wallet, "payments": payments, "cod_collections": cod}


def fetch_dashboard_rows(conn) -> dict[str, Any]:
    """Fetch dashboard-ready records from PostgreSQL.

    The dashboard logic still expects workbook-like rows for shipments and
    finance entries, so we keep the raw payloads as the primary bridge.
    """
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT
                order_id,
                tracking_number,
                merchant_name,
                store_name,
                customer_name,
                city,
                carrier,
                payment_type,
                status,
                shipment_date,
                delivery_date,
                weight,
                cod_amount,
                shipping_charge,
                customer_price_gross,
                customer_price_net,
                platform_cost_gross,
                platform_cost_net,
                base_profit,
                extra_kg,
                extra_profit,
                cod_profit,
                total_profit,
                actual_revenue,
                actual_base_cost,
                actual_extra_cost,
                actual_profit,
                included_in_profit,
                source_row,
                raw_payload
            FROM shipments
            ORDER BY shipment_date NULLS LAST, order_id
            """
        )
        shipments = [dict(row) for row in cur.fetchall()]

        cur.execute(
            """
            SELECT transaction_key, transaction_date, user_name, description, amount,
                   transaction_type, balance_before, balance_after, source_page, raw_payload
            FROM wallet_transactions
            ORDER BY transaction_date NULLS LAST, transaction_key
            """
        )
        wallet = [dict(row) for row in cur.fetchall()]

        cur.execute(
            """
            SELECT payment_key, payment_date, customer_name, amount, method, status, source_page, raw_payload
            FROM payments
            ORDER BY payment_date NULLS LAST, payment_key
            """
        )
        payments = [dict(row) for row in cur.fetchall()]

        cur.execute(
            """
            SELECT order_id, tracking_number, merchant_name, customer_name, cod_amount,
                   collection_date, transfer_date, settlement_status, shipment_status, raw_payload
            FROM cod_collections
            ORDER BY collection_date NULLS LAST, order_id
            """
        )
        cod = [dict(row) for row in cur.fetchall()]

    return {
        "shipments": shipments,
        "wallet": wallet,
        "payments": payments,
        "cod": cod,
    }


def prune_missing_shipments(conn, kept_order_ids: list[str]) -> int:
    """Delete shipments whose order_id is no longer returned by the scrape
    (drafts/orders removed on the site). Returns the number deleted. Caller must
    only invoke this after a healthy COMPLETE scrape, so a partial/failed scrape
    can't delete valid rows."""
    ids = [str(x).strip() for x in kept_order_ids if str(x).strip()]
    if not ids:
        return 0
    with conn.cursor() as cur:
        cur.execute("DELETE FROM shipments WHERE order_id <> ALL(%s)", (ids,))
        return cur.rowcount


def make_sync_run(conn, source: str) -> int:
    with conn.cursor() as cur:
        cur.execute("INSERT INTO sync_runs (source) VALUES (%s) RETURNING id", (source,))
        return int(cur.fetchone()["id"])


def finish_sync_run(conn, run_id: int, status: str, inserted: int = 0, updated: int = 0, skipped: int = 0, error_message: str | None = None, log_path: str | None = None) -> None:
    with conn.cursor() as cur:
        cur.execute(
            """
            UPDATE sync_runs
            SET finished_at = NOW(),
                status = %s,
                rows_inserted = %s,
                rows_updated = %s,
                rows_skipped = %s,
                error_message = %s,
                log_path = %s,
                updated_at = NOW()
            WHERE id = %s
            """,
            (status, inserted, updated, skipped, error_message, log_path, run_id),
        )


# ---------------------------------------------------------------------------
# Pricing model: carriers, per-merchant overrides, and global pricing settings.
# These three tables are the DB home for what used to live only in the Excel
# `الاسعار` sheet. Carrier prices are auto-synced from the Lead site; the COD
# fee, extra-kilo pricing, and merchant overrides are admin-editable.
#
# Price columns are intentionally NULLABLE: the profit engine's net fallback
# (`net = sheet_net or gross * 0.85`) only fires when a net is NULL, so NULLs
# must be preserved end-to-end — never coerce them to 0.
# ---------------------------------------------------------------------------

def get_setting(conn, key: str, default: str = "") -> str:
    """Read a free-text setting from the `settings` table (key TEXT, value TEXT)."""
    with conn.cursor() as cur:
        cur.execute("SELECT value FROM settings WHERE key = %s", (key,))
        row = cur.fetchone()
    if not row or row["value"] is None:
        return default
    return str(row["value"])


def set_setting(conn, key: str, value: str | None) -> None:
    with conn.cursor() as cur:
        cur.execute(
            """
            INSERT INTO settings (key, value) VALUES (%s, %s)
            ON CONFLICT (key) DO UPDATE SET value = EXCLUDED.value, updated_at = NOW()
            """,
            (key, value),
        )

_PRICE_COLUMNS = ("customer_gross", "customer_net", "platform_gross", "platform_net")


def upsert_carriers(conn, rows: list[dict[str, Any]]) -> tuple[int, int]:
    """Upsert carrier price rows keyed on carrier_name. Each row needs
    carrier_name plus any of the price columns (+ optional source)."""
    payload = []
    for row in rows:
        name = (row.get("carrier_name") or "").strip()
        if not name:
            continue
        payload.append({
            "carrier_name": name,
            "customer_gross": row.get("customer_gross"),
            "customer_net": row.get("customer_net"),
            "platform_gross": row.get("platform_gross"),
            "platform_net": row.get("platform_net"),
            "source": row.get("source", "scrape"),
        })
    if not payload:
        return 0, 0
    return upsert_rows(
        conn, "carriers", payload, ["carrier_name"],
        ["customer_gross", "customer_net", "platform_gross", "platform_net", "source"],
    )


def _floats(row: dict[str, Any], keys: tuple[str, ...]) -> dict[str, float | None]:
    out = {}
    for key in keys:
        val = row.get(key)
        out[key] = float(val) if val is not None else None
    return out


def load_pricing_snapshot(conn) -> dict[str, Any]:
    """Read the full pricing model in one shot. Returns a plain dict so this
    module stays IO-only; `scripts.pricing.PricingSnapshot.from_db` turns it
    into the typed snapshot the pure engine consumes."""
    with conn.cursor() as cur:
        cur.execute("SELECT carrier_name, customer_gross, customer_net, platform_gross, platform_net FROM carriers WHERE active")
        carriers = {r["carrier_name"]: _floats(r, _PRICE_COLUMNS) for r in cur.fetchall()}
        cur.execute("SELECT key, value FROM pricing_settings")
        settings = {r["key"]: (float(r["value"]) if r["value"] is not None else None) for r in cur.fetchall()}
    return {"carriers": carriers, "settings": settings}
