set -euo pipefail

cd /app
/app/.venv/bin/python - <<'PY'
import json, random
from openpyxl import load_workbook
import psycopg, os
from psycopg.rows import dict_row

wb = load_workbook("/app/output/lead6_report.xlsx", data_only=True)
details = wb["تفاصيل شهر 6"]
rows = [list(r) for r in details.iter_rows(min_row=2, values_only=True) if r and r[0] not in (None, "", "الإجمالي")]
shipments = len(rows)
revenue = sum(float(r[6] or 0) for r in rows)
platform = sum(float(r[17] or 0) for r in rows)
overweight = sum(float(r[20] or 0) for r in rows)
cod_profit = sum(float(r[21] or 0) for r in rows)
total = sum(float(r[22] or 0) for r in rows)

conn = psycopg.connect(os.environ["DATABASE_URL"], row_factory=dict_row)
cur = conn.cursor()
cur.execute("select count(*) c, coalesce(sum(total_profit),0) t, coalesce(sum(extra_profit),0) e, coalesce(sum(cod_profit),0) c2, coalesce(sum(cod_amount),0) a, coalesce(sum(shipping_charge),0) s from shipments")
pg = cur.fetchone()
cur.execute("select order_id, base_profit, extra_profit, cod_profit, total_profit from shipments order by random() limit 20")
samples = []
for r in cur.fetchall():
    d = dict(r)
    for k in ("base_profit", "extra_profit", "cod_profit", "total_profit"):
        d[k] = float(d[k] or 0)
    samples.append(d)
conn.close()

report = {
    "status": "ok",
    "metrics": [
        {"metric_name":"عدد الشحنات","current_value":shipments,"new_value":int(pg["c"]),"difference":int(pg["c"])-shipments,"difference_pct":0 if shipments==0 else abs(int(pg["c"])-shipments)/shipments*100,"reason":"", "ok": int(pg["c"])==shipments},
        {"metric_name":"الإيرادات","current_value":revenue,"new_value":float(pg["s"] if "s" in pg.keys() else pg["a"]),"difference":float(pg["s"] if "s" in pg.keys() else pg["a"])-revenue,"difference_pct":0 if revenue==0 else abs(float(pg["s"] if "s" in pg.keys() else pg["a"])-revenue)/revenue*100,"reason":"", "ok": abs(float(pg["s"] if "s" in pg.keys() else pg["a"])-revenue)<=0.01},
        {"metric_name":"ربح الوزن الزائد","current_value":overweight,"new_value":float(pg["e"]),"difference":float(pg["e"])-overweight,"difference_pct":0 if overweight==0 else abs(float(pg["e"])-overweight)/overweight*100,"reason":"", "ok": abs(float(pg["e"])-overweight)<=0.01},
        {"metric_name":"ربح COD","current_value":cod_profit,"new_value":float(pg["c2"]),"difference":float(pg["c2"])-cod_profit,"difference_pct":0 if cod_profit==0 else abs(float(pg["c2"])-cod_profit)/cod_profit*100,"reason":"", "ok": abs(float(pg["c2"])-cod_profit)<=0.01},
        {"metric_name":"إجمالي الربح","current_value":total,"new_value":float(pg["t"]),"difference":float(pg["t"])-total,"difference_pct":0 if total==0 else abs(float(pg["t"])-total)/total*100,"reason":"", "ok": abs(float(pg["t"])-total)<=0.01},
    ],
    "sample_size": len(samples),
    "samples": samples,
}
print(json.dumps(report, ensure_ascii=False, indent=2))
PY
