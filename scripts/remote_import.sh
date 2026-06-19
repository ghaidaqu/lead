set -euo pipefail

cd /app
LEAD_USERNAME=lead LEAD_PASSWORD=12345 LEAD_BASE_URL=https://lead-sa.com /app/.venv/bin/python - <<'PY'
import os, json, datetime as dt
from openpyxl import load_workbook
import psycopg
from psycopg.rows import dict_row

wb = load_workbook("/app/output/lead6_report.xlsx", data_only=True)
details = wb["تفاصيل شهر 6"]
ops = wb["العمليات المالية"]
conn = psycopg.connect(os.environ["DATABASE_URL"], row_factory=dict_row)
cur = conn.cursor()
cur.execute("""
CREATE TABLE IF NOT EXISTS shipments (
 id BIGSERIAL PRIMARY KEY, order_id TEXT NOT NULL UNIQUE, tracking_number TEXT, merchant_name TEXT, store_name TEXT, customer_name TEXT, city TEXT, carrier TEXT, payment_type TEXT, status TEXT, shipment_date DATE, delivery_date DATE, weight NUMERIC(10,2), cod_amount NUMERIC(12,2), shipping_charge NUMERIC(12,2), customer_price_gross NUMERIC(12,2), customer_price_net NUMERIC(12,2), platform_cost_gross NUMERIC(12,2), platform_cost_net NUMERIC(12,2), base_profit NUMERIC(12,2), extra_kg NUMERIC(10,2), extra_profit NUMERIC(12,2), cod_profit NUMERIC(12,2), total_profit NUMERIC(12,2), included_in_profit BOOLEAN NOT NULL DEFAULT TRUE, source_row INTEGER, source_hash TEXT, raw_payload JSONB, created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW());
CREATE TABLE IF NOT EXISTS wallet_transactions (
 id BIGSERIAL PRIMARY KEY, transaction_key TEXT NOT NULL UNIQUE, transaction_date TIMESTAMPTZ, user_name TEXT, description TEXT, amount NUMERIC(12,2), transaction_type TEXT, balance_before NUMERIC(12,2), balance_after NUMERIC(12,2), source_page TEXT, raw_payload JSONB, created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW());
CREATE TABLE IF NOT EXISTS payments (
 id BIGSERIAL PRIMARY KEY, payment_key TEXT NOT NULL UNIQUE, payment_date TIMESTAMPTZ, customer_name TEXT, amount NUMERIC(12,2), method TEXT, status TEXT, source_page TEXT, raw_payload JSONB, created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW());
CREATE TABLE IF NOT EXISTS cod_collections (
 id BIGSERIAL PRIMARY KEY, order_id TEXT NOT NULL UNIQUE, tracking_number TEXT, merchant_name TEXT, customer_name TEXT, cod_amount NUMERIC(12,2), collection_date DATE, transfer_date DATE, settlement_status TEXT, shipment_status TEXT, raw_payload JSONB, created_at TIMESTAMPTZ NOT NULL DEFAULT NOW(), updated_at TIMESTAMPTZ NOT NULL DEFAULT NOW());
""")

def n(v): return "" if v is None else str(v).strip()
def f(v):
    try: return float(v)
    except: return 0.0
def d(v):
    return v.date() if isinstance(v, dt.datetime) else v if isinstance(v, dt.date) else None

ship_rows=[]
for row in details.iter_rows(min_row=2, values_only=True):
    if not row or row[0] in (None, "", "الإجمالي"):
        continue
    ship_rows.append({
        "order_id": n(row[0]), "tracking_number": n(row[1]), "merchant_name": n(row[2]), "store_name": n(row[3]),
        "customer_name": n(row[4]), "city": n(row[5]), "carrier": n(row[10]), "payment_type": n(row[9]),
        "status": n(row[12]), "shipment_date": d(row[13]), "delivery_date": None, "weight": f(row[11]),
        "cod_amount": f(row[7]), "shipping_charge": f(row[6]), "customer_price_gross": f(row[16]),
        "customer_price_net": f(row[16]), "platform_cost_gross": f(row[17]), "platform_cost_net": f(row[17]),
        "base_profit": f(row[19]), "extra_kg": f(row[18]), "extra_profit": f(row[20]), "cod_profit": f(row[21]),
        "total_profit": f(row[22]), "included_in_profit": n(row[14]) == "نعم",
        "source_row": int(row[23]) if str(row[23]).isdigit() else None, "source_hash": n(row[23]),
        "raw_payload": json.dumps(row, default=str),
    })

ops_rows=[]
for row in ops.iter_rows(min_row=12, values_only=True):
    if not row or row[0] in (None, ""):
        continue
    ops_rows.append(row)

wallet=[]; payments=[]
for i, row in enumerate(ops_rows, 1):
    key=f"{row[0]}|{row[1]}|{row[2]}|{row[3]}|{i}"
    amt=f(row[2]) if len(row) > 2 else 0.0
    wallet.append({"transaction_key": key, "transaction_date": None, "user_name": n(row[0]), "description": n(row[0]), "amount": amt, "transaction_type": "finance", "balance_before": None, "balance_after": None, "source_page": "العمليات المالية", "raw_payload": json.dumps(row, default=str)})
    payments.append({"payment_key": key, "payment_date": None, "customer_name": n(row[0]), "amount": amt, "method": n(row[1]), "status": n(row[3]), "source_page": "العمليات المالية", "raw_payload": json.dumps(row, default=str)})

cod=[]
for row in ship_rows:
    cod.append({"order_id": row["order_id"], "tracking_number": row["tracking_number"], "merchant_name": row["merchant_name"], "customer_name": row["customer_name"], "cod_amount": row["cod_amount"], "collection_date": row["shipment_date"], "transfer_date": None, "settlement_status": row["status"], "shipment_status": row["status"], "raw_payload": json.dumps(row, default=str)})

for table, rows, keys, upd in [
    ("shipments", ship_rows, ["order_id"], ["tracking_number","merchant_name","store_name","customer_name","city","carrier","payment_type","status","shipment_date","delivery_date","weight","cod_amount","shipping_charge","customer_price_gross","customer_price_net","platform_cost_gross","platform_cost_net","base_profit","extra_kg","extra_profit","cod_profit","total_profit","included_in_profit","source_row","source_hash","raw_payload"]),
    ("wallet_transactions", wallet, ["transaction_key"], ["transaction_date","user_name","description","amount","transaction_type","balance_before","balance_after","source_page","raw_payload"]),
    ("payments", payments, ["payment_key"], ["payment_date","customer_name","amount","method","status","source_page","raw_payload"]),
    ("cod_collections", cod, ["order_id"], ["tracking_number","merchant_name","customer_name","cod_amount","collection_date","transfer_date","settlement_status","shipment_status","raw_payload"]),
]:
    if not rows:
        continue
    cols=list(rows[0].keys())
    q="INSERT INTO %s (%s) VALUES (%s) ON CONFLICT (%s) DO UPDATE SET %s, updated_at=NOW()" % (
        table,
        ", ".join(cols),
        ", ".join("%%(%s)s" % c for c in cols),
        ", ".join(keys),
        ", ".join("%s=EXCLUDED.%s" % (c, c) for c in upd),
    )
    for r in rows:
        cur.execute(q, r)
conn.commit()
conn.close()
print(json.dumps({"shipments": len(ship_rows), "wallet": len(wallet), "payments": len(payments), "cod": len(cod)}, ensure_ascii=False))
PY
