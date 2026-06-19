from __future__ import annotations

import datetime as dt
import hashlib
import json
import os
from contextlib import contextmanager
from pathlib import Path
from typing import Any, Iterable

try:
    import psycopg
    from psycopg.rows import dict_row
except Exception:  # pragma: no cover
    psycopg = None
    dict_row = None


PROJECT_DIR = Path(__file__).resolve().parents[1]
SCHEMA_VERSION = 1


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


def normalize_key(values: Iterable[Any]) -> str:
    packed = "|".join("" if v is None else str(v).strip() for v in values)
    return hashlib.sha256(packed.encode("utf-8")).hexdigest()


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

CREATE INDEX IF NOT EXISTS idx_shipments_status_date ON shipments(status, shipment_date);
CREATE INDEX IF NOT EXISTS idx_wallet_transactions_date ON wallet_transactions(transaction_date);
CREATE INDEX IF NOT EXISTS idx_payments_date ON payments(payment_date);
CREATE INDEX IF NOT EXISTS idx_cod_collections_dates ON cod_collections(collection_date, transfer_date);
"""


def ensure_schema(conn) -> None:
    with conn.cursor() as cur:
        cur.execute(SCHEMA_SQL)


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
            cur.execute(sql, row)
            if cur.rowcount == 1:
                inserted += 1
            else:
                updated += 1
    return inserted, updated


def replace_price_rules(conn, rows: list[list[Any]]) -> int:
    if not rows:
        return 0
    max_cols = max((len(row) for row in rows), default=0)
    inserted = 0
    with conn.cursor() as cur:
        cur.execute("DELETE FROM price_rules")
        for idx in range(1, max_cols):
            col_name = str(rows[0][idx]).strip() if len(rows[0]) > idx and rows[0][idx] is not None else f"col_{idx}"
            values = [rows[r][idx] if len(rows[r]) > idx else None for r in range(min(len(rows), 6))]
            if not any(cell not in (None, "") for cell in values):
                continue
            mapping = {
                "merchant_name": col_name,
                "carrier_name": col_name,
                "customer_gross": values[1] if len(values) > 1 else None,
                "customer_net": values[2] if len(values) > 2 else None,
                "platform_gross": values[3] if len(values) > 3 else None,
                "platform_net": values[4] if len(values) > 4 else None,
                "tax_mode": None,
                "extra_kilo_mode": None,
                "cod_mode": None,
                "raw_payload": values,
            }
            cur.execute(
                """
                INSERT INTO price_rules (
                    merchant_name, carrier_name, customer_gross, customer_net,
                    platform_gross, platform_net, tax_mode, extra_kilo_mode, cod_mode, raw_payload
                ) VALUES (%(merchant_name)s, %(carrier_name)s, %(customer_gross)s, %(customer_net)s,
                          %(platform_gross)s, %(platform_net)s, %(tax_mode)s, %(extra_kilo_mode)s, %(cod_mode)s, %(raw_payload)s)
                """,
                mapping,
            )
            inserted += 1
    return inserted


def fetch_price_rules_rows(conn) -> list[list[Any]]:
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT merchant_name, carrier_name, customer_gross, customer_net,
                   platform_gross, platform_net, tax_mode, extra_kilo_mode, cod_mode, raw_payload
            FROM price_rules
            ORDER BY id
            """
        )
        rows = []
        for row in cur.fetchall():
            payload = row["raw_payload"]
            rows.append([
                row["merchant_name"],
                row["carrier_name"],
                row["customer_gross"],
                row["customer_net"],
                row["platform_gross"],
                row["platform_net"],
                row["tax_mode"],
                row["extra_kilo_mode"],
                row["cod_mode"],
                payload,
            ])
        return rows


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


def aggregate_shipments(conn) -> dict[str, Any]:
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT
                COUNT(*) AS total_count,
                SUM(CASE WHEN included_in_profit THEN 1 ELSE 0 END) AS included_count,
                SUM(CASE WHEN NOT included_in_profit THEN 1 ELSE 0 END) AS excluded_count,
                COALESCE(SUM(base_profit), 0) AS base_total,
                COALESCE(SUM(extra_profit), 0) AS extra_total,
                COALESCE(SUM(cod_profit), 0) AS cod_total,
                COALESCE(SUM(total_profit), 0) AS total_profit,
                COALESCE(SUM(cod_amount), 0) AS cod_amount,
                COALESCE(SUM(CASE WHEN payment_type = '💵 COD' AND included_in_profit THEN 1 ELSE 0 END), 0) AS cod_count,
                COALESCE(SUM(CASE WHEN extra_kg > 0 AND included_in_profit THEN 1 ELSE 0 END), 0) AS overweight_count
            FROM shipments
            """
        )
        row = dict(cur.fetchone())
        for key in list(row.keys()):
            if key.endswith("count"):
                row[key] = int(row[key] or 0)
            else:
                row[key] = float(row[key] or 0)
        return row


def aggregate_top_entities(conn, table_name: str, group_column: str, total_column: str = "total_profit", limit: int = 5) -> list[dict[str, Any]]:
    with conn.cursor() as cur:
        cur.execute(
            f"""
            SELECT {group_column} AS label,
                   COUNT(*) AS count,
                   COALESCE(SUM(base_profit), 0) AS base_total,
                   COALESCE(SUM(extra_profit), 0) AS extra_total,
                   COALESCE(SUM(cod_profit), 0) AS cod_total,
                   COALESCE(SUM({total_column}), 0) AS total
            FROM {table_name}
            WHERE included_in_profit
            GROUP BY {group_column}
            ORDER BY total DESC NULLS LAST
            LIMIT %s
            """,
            (limit,),
        )
        out = []
        for row in cur.fetchall():
            item = dict(row)
            item["count"] = int(item["count"] or 0)
            for key in ("base_total", "extra_total", "cod_total", "total"):
                item[key] = float(item[key] or 0)
            out.append(item)
        return out


def shipment_samples(conn, order_ids: list[str]) -> dict[str, dict[str, Any]]:
    if not order_ids:
        return {}
    with conn.cursor() as cur:
        cur.execute(
            """
            SELECT order_id, tracking_number, merchant_name, city, carrier, status, shipment_date,
                   weight, cod_amount, shipping_charge, customer_price_net, platform_cost_net,
                   base_profit, extra_kg, extra_profit, cod_profit, total_profit, included_in_profit
            FROM shipments
            WHERE order_id = ANY(%s)
            """,
            (order_ids,),
        )
        out = {}
        for row in cur.fetchall():
            item = dict(row)
            for key in ("weight", "cod_amount", "shipping_charge", "customer_price_net", "platform_cost_net", "base_profit", "extra_kg", "extra_profit", "cod_profit", "total_profit"):
                item[key] = float(item[key] or 0)
            out[str(row["order_id"])] = item
        return out


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
