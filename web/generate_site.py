import json
import os
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl


PROJECT_DIR = Path(__file__).resolve().parents[1]
SOURCE = PROJECT_DIR / "output" / "lead6_report.xlsx"
OUT_DIR = Path(__file__).resolve().parent
REPORT_COPY = OUT_DIR / "lead6_report.xlsx"
INDEX = OUT_DIR / "index.html"


def clean(value):
    return str(value).strip() if value is not None else ""


def money(value):
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = clean(value).replace(",", "")
    num = "".join(ch for ch in s if ch.isdigit() or ch in ".-")
    if num in ("", "-", ".", "-."):
        return 0.0
    return float(num)


def load_data():
    wb = openpyxl.load_workbook(SOURCE, data_only=True)
    ws = wb["تفاصيل شهر 6"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    ops = wb["العمليات المالية"] if "العمليات المالية" in wb.sheetnames else None
    stmt = wb["كشف الحساب"] if "كشف الحساب" in wb.sheetnames else None

    items = []
    cod_items = []
    by_merchant = defaultdict(lambda: {"count": 0, "total": 0.0})
    by_carrier = defaultdict(lambda: {"count": 0, "total": 0.0})
    by_status = defaultdict(int)
    by_date = defaultdict(float)
    by_date_count = defaultdict(int)
    finance = {
        "bank": {"count": 0, "total": 0.0},
        "moyasar": {"count": 0, "total": 0.0},
        "shipping_return": {"count": 0, "total": 0.0},
        "tax_deduction": {"count": 0, "total": 0.0},
        "shipping_refund": {"count": 0, "total": 0.0},
        "other_net": 0.0,
        "other": {"count": 0, "total": 0.0},
    }
    statement = {
        "summary": {},
        "rows": [],
    }

    for row in rows:
        if not any(v is not None for v in row):
            continue
        item = {
            "order_id": row[0],
            "tracking": clean(row[1]),
            "merchant": clean(row[2]),
            "customer": clean(row[4]),
            "city": clean(row[5]),
            "carrier": clean(row[10]),
            "date": row[13].date().isoformat() if isinstance(row[13], datetime) else clean(row[13]),
            "status": clean(row[12]),
            "weight": money(row[11]),
            "included": clean(row[14]) == "نعم",
            "customer_shipping": money(row[6]),
            "cod_amount": money(row[7]),
            "customer_net": money(row[16]),
            "platform_shipping": money(row[17]),
            "shipping_profit": money(row[19]),
            "extra_profit": money(row[20]),
            "fee_profit": money(row[21]),
            "total_profit": money(row[22]),
            "review_diff": money(row[19]),
        }
        if not item["included"]:
            continue
        items.append(item)
        if item["cod_amount"] > 0:
            cod_items.append(item)
        by_merchant[item["merchant"]]["count"] += 1
        by_merchant[item["merchant"]]["total"] += item["total_profit"]
        by_carrier[item["carrier"]]["count"] += 1
        by_carrier[item["carrier"]]["total"] += item["total_profit"]
        by_status[item["status"]] += 1
        if item["date"]:
            by_date[item["date"]] += item["total_profit"]
            by_date_count[item["date"]] += 1

    if ops is not None:
        for r in range(2, ops.max_row + 1):
            typ = clean(ops.cell(r, 4).value)
            note = clean(ops.cell(r, 6).value)
            amount = money(ops.cell(r, 5).value)
            finance["other_net"] += amount if typ != "إيداع" else 0.0
            if typ == "إيداع":
                key = "bank" if "تحويل بنكي" in note else "moyasar" if "Moyasar" in note else "other"
                finance[key]["count"] += 1
                finance[key]["total"] += amount
            else:
                if typ == "admin_deduction" and "خصم تكلفة شحنة مرتجعة" in note:
                    finance["shipping_return"]["count"] += 1
                    finance["shipping_return"]["total"] += abs(amount)
                elif typ == "admin_deduction" and "خصم ضريبة القيمة المضافة" in note:
                    finance["tax_deduction"]["count"] += 1
                    finance["tax_deduction"]["total"] += abs(amount)
                elif typ == "استرداد" and "استرداد تكلفة شحن" in note:
                    finance["shipping_refund"]["count"] += 1
                    finance["shipping_refund"]["total"] += amount
                else:
                    finance["other"]["count"] += 1
                    finance["other"]["total"] += amount
    finance["total"] = finance["bank"]["total"] + finance["moyasar"]["total"]

    if stmt is not None:
        statement["summary"]["deposits_total"] = money(stmt["B4"].value)
        statement["summary"]["expenses_total"] = money(stmt["B5"].value)
        statement["summary"]["net_total"] = money(stmt["B6"].value)
        statement["summary"]["transfer_fees_total"] = money(stmt["B7"].value)
        statement["summary"]["deposits_count"] = money(stmt["E4"].value)
        statement["summary"]["expenses_count"] = money(stmt["E5"].value)
        statement["summary"]["period"] = clean(stmt["E6"].value)
        for r in range(10, stmt.max_row + 1):
            date_val = stmt.cell(r, 1).value
            move_type = clean(stmt.cell(r, 2).value)
            expense_type = clean(stmt.cell(r, 3).value)
            amount = money(stmt.cell(r, 4).value)
            if not any([date_val, move_type, expense_type, amount]):
                continue
            statement["rows"].append({
                "date": date_val.date().isoformat() if isinstance(date_val, datetime) else clean(date_val),
                "move_type": move_type,
                "expense_type": expense_type,
                "amount": amount,
            })

    items.sort(key=lambda x: (x["date"], x["order_id"]), reverse=True)
    top_merchants = sorted(by_merchant.items(), key=lambda kv: kv[1]["total"], reverse=True)[:5]
    top_carriers = sorted(by_carrier.items(), key=lambda kv: kv[1]["total"], reverse=True)
    daily = sorted(by_date.items(), key=lambda kv: kv[0])
    daily_count = sorted(by_date_count.items(), key=lambda kv: kv[0])
    totals = {
        "records": len(rows),
        "active": len(items),
        "shipping": len(items),
        "cod": sum(1 for item in items if item["fee_profit"] > 0),
        "cod_amount": sum(item["cod_amount"] for item in items if item["included"]),
        "base": sum(item["shipping_profit"] for item in items),
        "extra": sum(item["extra_profit"] for item in items),
        "cod_profit": sum(item["fee_profit"] for item in items),
        "total": sum(item["total_profit"] for item in items),
        "excluded": len(rows) - len(items),
    }
    cod_items.sort(key=lambda x: (x["date"], x["order_id"]), reverse=True)
    return totals, top_merchants, top_carriers, by_status, daily, daily_count, finance, cod_items, statement


def fmt_money(value):
    return f"{value:,.2f} ريال"


def build_step_area_chart(series, color="#8a6f68", value_formatter=None, chart_key="chart"):
    if not series:
        return "<div class='empty-state'>لا توجد بيانات كافية للرسم</div>"

    points = [(str(date), float(value)) for date, value in series]
    max_val = max(value for _, value in points) or 1
    width = 760
    height = 240
    pad_left = 28
    pad_right = 18
    pad_top = 56
    pad_bottom = 48
    chart_width = width - pad_left - pad_right
    chart_height = height - pad_top - pad_bottom
    step = chart_width / max(len(points) - 1, 1)

    xs = [pad_left + step * idx for idx in range(len(points))]
    ys = [pad_top + (1 - (value / max_val)) * chart_height for _, value in points]
    base_y = height - pad_bottom

    line_commands = [f"M {xs[0]:.1f} {ys[0]:.1f}"]
    for idx in range(1, len(points)):
        line_commands.append(f"H {xs[idx]:.1f}")
        line_commands.append(f"V {ys[idx]:.1f}")

    fill_commands = [
        f"M {xs[0]:.1f} {base_y:.1f}",
        f"L {xs[0]:.1f} {ys[0]:.1f}",
    ]
    for idx in range(1, len(points)):
        fill_commands.append(f"H {xs[idx]:.1f}")
        fill_commands.append(f"V {ys[idx]:.1f}")
    fill_commands.extend([
        f"H {xs[-1]:.1f}",
        f"V {base_y:.1f}",
        "Z",
    ])

    point_marks = []
    value_labels = []
    date_labels = []
    for idx, ((date, value), x, y) in enumerate(zip(points, xs, ys)):
        short_date = date[5:] if len(date) >= 10 else date
        if value_formatter is None:
            value_text = fmt_money(value).replace(" ريال", "")
        else:
            value_text = value_formatter(value)
        point_marks.append(
            f"<circle cx='{x:.1f}' cy='{y:.1f}' r='4.2' fill='{color}' stroke='rgba(255,255,255,0.95)' stroke-width='2' />"
        )
        label_width = max(34, min(78, 12 + len(value_text) * 6.2))
        chosen_row = idx % 2
        label_y = 14 + (chosen_row * 20)
        side_offset = -18 if chosen_row == 0 else 18
        label_x = max(20, min(width - label_width - 20, x + side_offset - label_width / 2))
        line_y = label_y + 18
        value_labels.append(
            f"<g class='area-value-wrap'><line x1='{x:.1f}' y1='{y - 6:.1f}' x2='{x:.1f}' y2='{line_y:.1f}' stroke='rgba(138,111,104,0.18)' stroke-width='1' /><rect x='{label_x:.1f}' y='{label_y:.1f}' rx='9' ry='9' width='{label_width:.1f}' height='18' fill='rgba(255,253,249,0.93)' stroke='rgba(225,212,201,0.72)' stroke-width='1' /><text x='{label_x + label_width/2:.1f}' y='{label_y + 13:.1f}' text-anchor='middle' class='area-value'>{value_text}</text></g>"
        )
        date_labels.append(
            f"<text x='{x:.1f}' y='{height - 18}' text-anchor='middle' class='area-date'>{short_date}</text>"
        )

    return f"""
      <svg viewBox='0 0 {width} {height}' role='img' aria-label='Area Chart - Step'>
        <defs>
          <linearGradient id='areaFill-{chart_key}' x1='0' x2='0' y1='0' y2='1'>
            <stop offset='0%' stop-color='rgba(138,111,104,0.34)' />
            <stop offset='100%' stop-color='rgba(138,111,104,0.06)' />
          </linearGradient>
          <linearGradient id='areaStroke-{chart_key}' x1='0' x2='1' y1='0' y2='0'>
            <stop offset='0%' stop-color='{color}' />
            <stop offset='100%' stop-color='#2f2628' />
          </linearGradient>
        </defs>
        <path d='{ " ".join(fill_commands) }' fill='url(#areaFill-{chart_key})' />
        <path d='{ " ".join(line_commands) }' fill='none' stroke='url(#areaStroke-{chart_key})' stroke-width='3.2' stroke-linecap='round' stroke-linejoin='round' />
        {''.join(point_marks)}
        {''.join(value_labels)}
        {''.join(date_labels)}
      </svg>
    """


def metric_card(label, value, note="", tone="navy"):
    if label == "إجمالي الربح":
        return f"""
      <article class="metric-card tone-{tone}">
        <span class="metric-label">{label}</span>
        <strong>{value}</strong>
        <div class="metric-footer">{note}</div>
      </article>
    """
    return f"""
      <article class="metric-card tone-{tone}">
        <span class="metric-label">{label}</span>
        <strong>{value}</strong>
        <div class="metric-footer">{note}</div>
      </article>
    """


def build_html(data):
    totals = data["totals"]
    daily = data["daily"]
    daily_count = data["daily_count"]
    finance = data["finance"]
    cod_items = data["cod_items"]
    top_carriers = data["top_carriers"]
    statement = data.get("statement", {"summary": {}, "rows": []})
    details_href = "lead6_report.xlsx"
    supplemental_cod_items = []
    all_cod_items = cod_items + supplemental_cod_items
    cod_overrides = {
        "1757": {"collection_date": "2026-06-12", "transfer_date": "2026-06-13"},
        "1756": {"collection_date": "2026-06-11", "transfer_date": "2026-06-13"},
    }

    finance_cards = f"""
      <div class="finance-grid">
        <article class="finance-card bank">
          <span class="metric-label">إيداعات البنك</span>
          <strong>{finance['bank']['count']}</strong>
          <div class="metric-footer">{fmt_money(finance['bank']['total'])}</div>
        </article>
        <article class="finance-card moyasar">
          <span class="metric-label">إيداعات ميسر</span>
          <strong>{finance['moyasar']['count']}</strong>
          <div class="metric-footer">{fmt_money(finance['moyasar']['total'])}</div>
        </article>
        <article class="finance-card other">
          <span class="metric-label">تفصيل الخصومات والاستردادات</span>
          <div class="metric-footer breakdown-line">خصم شحنة مرتجعة: {fmt_money(finance['shipping_return']['total'])}</div>
          <div class="metric-footer breakdown-line">خصم الضريبة: {fmt_money(finance['tax_deduction']['total'])}</div>
          <div class="metric-footer breakdown-line">استرداد تكلفة شحن: {fmt_money(finance['shipping_refund']['total'])}</div>
          <div class="metric-footer">الصافي: {fmt_money(finance['other_net'])}</div>
        </article>
        <article class="finance-card other">
          <span class="metric-label">إجمالي الإيداعات</span>
          <strong>{fmt_money(finance['total'])}</strong>
          <div class="metric-footer">البنك + ميسر</div>
        </article>
      </div>
    """

    chart_html = build_step_area_chart(daily[-7:], color="#8a6f68", chart_key="profit")

    count_chart_html = build_step_area_chart(
        daily_count[-7:],
        color="#2f2628",
        value_formatter=lambda value: str(int(value)),
        chart_key="count",
    )

    statement_rows = "".join(
        f"<tr class='expense-row'><td>{row['date']}</td><td>{row['expense_type'] or '-'}</td><td>{fmt_money(row['amount'])}</td></tr>"
        for row in statement["rows"]
        if row["move_type"] == "مصروف"
    )

    cod_rows = "".join(
        f"<tr><td>{item['order_id']}</td><td>{item['merchant']}</td><td>{fmt_money(item['cod_amount'])}</td><td>{item.get('order_date') or item.get('date') or '-'}</td><td>{(cod_overrides.get(str(item['order_id'])) or {}).get('collection_date') or item.get('collection_date') or item.get('date') or '-'}</td><td>{(cod_overrides.get(str(item['order_id'])) or {}).get('transfer_date') or item.get('transfer_date') or '-'}</td></tr>"
        for item in all_cod_items
    )
    carrier_total = sum(data['total'] for _, data in top_carriers) or 1
    carrier_colors = ["#5b4a47", "#c59f95", "#d2b3aa", "#dfc3ba", "#ead5cf", "#a98c84"]
    carrier_legend_items = []
    carrier_stops = []
    running_pct = 0.0
    for idx, (name, data) in enumerate(top_carriers):
        pct = data["total"] / carrier_total
        color = carrier_colors[idx % len(carrier_colors)]
        start = running_pct * 100
        end = (running_pct + pct) * 100
        carrier_stops.append(f"{color} {start:.2f}% {end:.2f}%")
        carrier_legend_items.append(
            f"""
              <div class="carrier-legend-item">
                <span class="carrier-swatch" style="background:{color}"></span>
                <span class="carrier-legend-name">{("ARAMEX - ارامكس" if name == "ارامكس - ARAMEX" else name)}</span>
                <span class="carrier-legend-meta">{pct*100:.1f}%</span>
              </div>
            """
        )
        running_pct += pct
    carrier_donut_style = "; ".join([
        f"background: conic-gradient({', '.join(carrier_stops)})",
        "box-shadow: 0 18px 32px rgba(47,38,40,0.16), inset 0 1px 0 rgba(255,255,255,0.9)",
    ])
    carrier_legend_html = "".join(carrier_legend_items)
    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>lead dashboard</title>
  <style>
    :root {{
      color-scheme: light;
      --bg: #f4efe8;
      --surface: #fffdfb;
      --surface-soft: #f8f3ed;
      --text: #2a2425;
      --muted: #7c6b68;
      --line: #e1d4ca;
      --teal: #8b746d;
      --blue: #9a7f77;
      --navy: #2f2628;
      --green: #86756f;
      --amber: #b39289;
      --coral: #c5a59b;
      --shadow: 0 14px 34px rgba(31, 41, 55, 0.07);
    }}
    * {{ box-sizing: border-box; }}
    html {{ scroll-behavior:smooth; }}
    body {{ margin:0; min-height:100vh; font-family: "Avenir Next", "Segoe UI", "Noto Sans Arabic", system-ui, sans-serif; background:
      radial-gradient(circle at top right, rgba(255,255,255,0.72), transparent 22%),
      radial-gradient(circle at 16% 18%, rgba(255,255,255,0.50), transparent 16%),
      radial-gradient(circle at bottom left, rgba(47, 38, 40, 0.04), transparent 28%),
      linear-gradient(180deg, #fffdf8 0%, #fbf7f1 52%, #f5efe8 100%); color:var(--text); }}
    a {{ color: inherit; text-decoration: none; }}
    .page {{ max-width: 1440px; margin: 0 auto; padding: 32px 28px 36px; }}
    .hero {{ position:relative; display:flex; align-items:flex-start; justify-content:flex-end; margin-bottom:10px; padding:10px 24px 14px; overflow:hidden; background:linear-gradient(135deg, #6f544a 0%, #8b6f64 55%, #a3877b 100%); border-radius:0 0 28px 28px; box-shadow:0 14px 30px rgba(47,38,40,0.12); }}
    .hero::before {{ content:''; position:absolute; inset:auto -4% -28px -4%; height:82px; background:
      radial-gradient(circle at 14% 12%, rgba(255,255,255,0.18), transparent 26%),
      radial-gradient(circle at 42% 0%, rgba(255,255,255,0.10), transparent 24%),
      radial-gradient(circle at 74% 10%, rgba(255,255,255,0.14), transparent 28%),
      radial-gradient(ellipse at center, rgba(255,255,255,0.14), rgba(255,255,255,0.04) 42%, rgba(255,255,255,0) 76%);
      opacity:.9; pointer-events:none; }}
    .brand-lockup {{ position:relative; z-index:1; display:flex; align-items:flex-start; justify-content:flex-end; padding:0; border:0; background:transparent; box-shadow:none; }}
    .brand-logo {{ width:min(150px, 16vw); max-width:150px; height:auto; display:block; background:transparent; mix-blend-mode:multiply; filter: sepia(0.36) saturate(0.42) brightness(1.20) contrast(0.92); }}
    .eyebrow {{ margin:0 0 8px; color:var(--teal); font-size:12px; font-weight:800; letter-spacing:0; font-family: "Avenir Next", "Segoe UI", "Noto Sans Arabic", sans-serif; }}
    h1, h2, p {{ margin-top:0; }}
    h1, h2 {{ font-family: "Avenir Next", "Segoe UI", "Noto Sans Arabic", system-ui, sans-serif; letter-spacing:0; }}
    h2 {{ margin-bottom:0; font-size:16px; line-height:1.24; font-weight:600; }}
    .hero p {{ margin:8px 0 0; color:var(--muted); }}
    .hero-tools {{ display:flex; align-items:center; justify-content:flex-end; gap:8px; flex-wrap:wrap; margin:0 0 12px; }}
    .badge {{ display:inline-flex; align-items:center; min-height:38px; padding:0 12px; border:1px solid rgba(225,212,201,0.96); border-radius:999px; background:rgba(255,255,255,0.98); color:var(--muted); font-size:12px; backdrop-filter: blur(6px); box-shadow:0 8px 18px rgba(47,38,40,0.05); }}
    .primary-button {{ display:inline-flex; align-items:center; min-height:40px; padding:0 16px; border-radius:999px; background:linear-gradient(135deg, #2f2628, #8a6f68); color:#fff; font-weight:700; box-shadow:0 10px 20px rgba(47,38,40,0.16); }}
    .ghost-button {{ display:inline-flex; align-items:center; min-height:40px; padding:0 14px; border-radius:999px; border:1px solid var(--line); background:rgba(255,255,255,0.88); color:var(--text); font-weight:700; }}
    .hero-tools .badge {{ min-height:32px; padding:0 10px; font-size:11px; }}
    .hero-tools .primary-button {{ min-height:32px; padding:0 12px; font-size:12px; }}
    .metric-grid {{ display:grid; grid-template-columns:repeat(4, minmax(0,1fr)); gap:14px; margin-bottom:16px; }}
    .metric-card, .panel {{
      background:
        linear-gradient(145deg, rgba(255,255,255,0.84) 0%, rgba(255,255,255,0.52) 40%, rgba(255,255,255,0.28) 100%),
        linear-gradient(135deg, rgba(255,255,255,0.24), rgba(248,244,238,0.14));
      border:1px solid rgba(255,255,255,0.96);
      border-radius:18px;
      position:relative;
      overflow:hidden;
      box-shadow:
        0 18px 32px rgba(47,38,40,0.08),
        0 1px 0 rgba(255,255,255,0.98) inset,
        0 -1px 0 rgba(255,255,255,0.22) inset;
      backdrop-filter: blur(22px) saturate(122%);
      -webkit-backdrop-filter: blur(22px) saturate(122%);
    }}
    .metric-card::before, .panel::before, .mini-card::before {{
      content:'';
      position:absolute;
      inset:0;
      background:
        linear-gradient(135deg, rgba(255,255,255,0.90) 0%, rgba(255,255,255,0.44) 22%, rgba(255,255,255,0.16) 44%, rgba(255,255,255,0) 60%),
        linear-gradient(180deg, rgba(255,255,255,0.50) 0%, rgba(255,255,255,0) 24%);
      pointer-events:none;
    }}
    .metric-card > *, .panel > *, .mini-card > * {{ position:relative; z-index:1; }}
    .metric-card {{ padding:18px; min-height:118px; display:flex; flex-direction:column; justify-content:space-between; }}
    .metric-label {{ white-space:normal; line-height:1.45; }}
    .metric-card strong {{ display:block; font-size:20px; line-height:1.15; margin-top:8px; font-variant-numeric: tabular-nums; }}
    .metric-footer {{ margin-top:8px; color:var(--muted); font-size:12px; line-height:1.35; }}
    .profit-card {{ min-height:118px; justify-content:space-between; }}
    .panel {{ padding:18px; }}
    .panel-header {{ display:flex; align-items:flex-start; justify-content:space-between; gap:16px; margin-bottom:14px; }}
    .panel-header p {{ margin:0; color:var(--muted); font-size:13px; }}
    .panel-grid {{ display:grid; grid-template-columns:2fr 1fr; gap:14px; align-items:start; }}
    .top-layout {{ display:grid; grid-template-columns:minmax(0, 0.9fr) minmax(0, 1.1fr); gap:14px; align-items:stretch; margin-bottom:14px; }}
    .stack-panels {{ display:grid; gap:10px; }}
    .chart-block + .chart-block {{ margin-top:16px; padding-top:16px; border-top:1px solid var(--line); }}
    .stack-panels .panel:first-child {{ min-height: 250px; }}
    .stack-panels .panel:first-child .chart-block:last-child {{ margin-bottom: 4px; }}
    .stack-panels > .panel:nth-child(2) {{ min-height: 176px; margin-top:-10px; }}
    .top-layout > .stack-panels:nth-child(2) > .panel:first-child {{ min-height: 436px; display:flex; flex-direction:column; }}
    .shipments-panel {{ text-align:right; }}
    .chart {{ height:165px; display:flex; align-items:flex-end; gap:10px; padding-top:4px; }}
    .area-chart {{ width:100%; height:210px; display:block; }}
    .area-chart svg {{ width:100%; height:100%; display:block; overflow:visible; }}
    .area-value {{ fill:var(--navy); font-size:11px; font-weight:800; letter-spacing:0; font-family:"Avenir Next", "Segoe UI", "Noto Sans Arabic", system-ui, sans-serif; }}
    .area-date {{ fill:var(--muted); font-size:10px; font-weight:700; letter-spacing:0; font-family:"Avenir Next", "Segoe UI", "Noto Sans Arabic", system-ui, sans-serif; }}
    .bar {{ flex:1; min-width:0; height:var(--h); display:flex; flex-direction:column; justify-content:flex-end; align-items:center; }}
    .bar::before {{ content:''; width:100%; height:100%; border-radius:8px 8px 0 0; background:linear-gradient(180deg, #8a6f68, #2f2628); opacity:.96; }}
    .bar strong {{ display:block; margin-bottom:6px; font-size:10px; color:var(--muted); }}
    .bar span {{ margin-top:6px; font-size:10px; color:var(--muted); }}
    .stack {{ display:grid; gap:14px; }}
    .mini-grid {{ display:grid; grid-template-columns:repeat(3, minmax(0,1fr)); gap:12px; }}
    .mini-card {{
      padding:14px;
      min-height:92px;
      border:1px solid rgba(255,255,255,0.86);
      border-radius:14px;
      position:relative;
      overflow:hidden;
      background:
        linear-gradient(180deg, rgba(255,255,255,0.76), rgba(255,255,255,0.42)),
        linear-gradient(135deg, rgba(255,255,255,0.22), rgba(248,244,238,0.14));
      display:flex;
      flex-direction:column;
      justify-content:space-between;
      box-shadow:
        0 16px 28px rgba(47,38,40,0.08),
        0 1px 0 rgba(255,255,255,0.98) inset;
      backdrop-filter: blur(20px) saturate(122%);
      -webkit-backdrop-filter: blur(20px) saturate(122%);
    }}
    .mini-card strong {{ display:block; margin-top:6px; font-size:20px; line-height:1.15; font-variant-numeric: tabular-nums; }}
    .mini-card .metric-footer {{ font-size:12px; line-height:1.35; }}
    .mini-card.quick-card {{ min-height:92px; background:
      linear-gradient(145deg, rgba(255,255,255,0.78), rgba(255,255,255,0.42)),
      linear-gradient(135deg, rgba(255,255,255,0.20), rgba(248,244,238,0.14));
      border:1px solid rgba(255,255,255,0.96); box-shadow:0 16px 28px rgba(47,38,40,0.08), 0 1px 0 rgba(255,255,255,0.98) inset; backdrop-filter: blur(20px) saturate(122%); -webkit-backdrop-filter: blur(20px) saturate(122%); }}
    .mini-card.quick-card strong {{ font-size:18px; line-height:1.15; font-weight:800; }}
    .mini-card.quick-card .metric-label {{ font-size:13px; line-height:1.35; }}
    .mini-card.finance-card strong {{ font-size:18px; line-height:1.15; }}
    .mini-card.finance-card .metric-footer {{ font-size:13px; line-height:1.45; }}
    .mini-card.finance-card .metric-label {{ font-size:13px; line-height:1.35; }}
    .mini-card.finance-card.bank strong,
    .mini-card.finance-card.moyasar strong {{ font-size:18px; line-height:1.15; font-weight:800; }}
    .mini-card.finance-total-card strong {{ font-size:16px; line-height:1.15; }}
    .mini-card.finance-total-card .metric-label {{ font-size:12px; line-height:1.35; }}
    .mini-card.finance-total-card .metric-footer {{ font-size:12px; line-height:1.35; }}
    .breakdown-line {{ font-size:12px; line-height:1.55; margin-top:2px; color:var(--muted); }}
    .finance-panel .panel-header p {{ font-size:13px; line-height:1.35; }}
    .finance-panel .table-wrap {{ margin-top:2px; }}
    .finance-panel th {{ font-size:13px; }}
    .finance-panel td {{ font-size:13px; }}
    .statement-panel {{ margin-top:14px; }}
    .statement-summary {{ display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:12px; margin-bottom:14px; }}
    .statement-summary .mini-card strong {{ font-size:18px; }}
    .statement-table th:nth-child(3),
    .statement-table td:nth-child(3) {{ white-space:nowrap; font-variant-numeric: tabular-nums; font-weight:800; }}
    .expense-row {{ background:linear-gradient(180deg, rgba(255,248,246,0.9), rgba(255,243,240,0.7)); }}
    .expense-row td {{ color:#4a3634; }}
    .mini-card.bank {{ background:linear-gradient(145deg, rgba(255,255,255,0.80), rgba(255,255,255,0.44)); }}
    .mini-card.moyasar {{ background:linear-gradient(145deg, rgba(255,255,255,0.80), rgba(255,255,255,0.44)); }}
    .mini-card.other {{ background:linear-gradient(145deg, rgba(255,255,255,0.80), rgba(255,255,255,0.44)); }}
    .carrier-donut-panel {{ display:grid; grid-template-columns:minmax(95px, 120px) 1fr; gap:8px; align-items:center; }}
    .carrier-donut {{ position:relative; width:min(100%, 120px); margin:0 auto; aspect-ratio:1; border-radius:50%; padding:4px; {carrier_donut_style}; }}
    .carrier-donut-center {{ position:absolute; inset:50%; transform:translate(-50%, -50%); width:40%; height:40%; border-radius:50%; background:transparent; box-shadow:none; display:flex; flex-direction:column; align-items:center; justify-content:center; text-align:center; padding:0; color:#fff; }}
    .carrier-donut-center strong {{ display:none; }}
    .carrier-donut-center span {{ display:none; }}
    .carrier-donut-note {{ text-align:center; color:var(--muted); font-size:12px; margin-top:8px; }}
    .carrier-legend {{ display:grid; gap:10px; }}
    .carrier-legend-item {{ display:grid; grid-template-columns:12px 1fr auto; gap:10px; align-items:center; padding:10px 12px; border:1px solid rgba(225,212,201,0.82); border-radius:12px; background:linear-gradient(180deg, #fffdfa, #f7f1ea); box-shadow:0 10px 20px rgba(47,38,40,0.04); }}
    .carrier-swatch {{ width:12px; height:12px; border-radius:50%; box-shadow:0 0 0 3px rgba(255,255,255,0.65); }}
    .carrier-legend-name {{ font-size:12px; font-weight:800; color:var(--text); }}
    .carrier-legend-meta {{ font-size:11px; color:var(--muted); white-space:nowrap; }}
    .table-wrap {{ overflow:auto; }}
    table {{ width:100%; border-collapse:collapse; }}
    th, td {{ padding:11px 10px; border-bottom:1px solid var(--line); text-align:right; white-space:nowrap; }}
    th {{ color:var(--muted); font-size:13px; font-weight:700; }}
    .status-list {{ display:grid; gap:10px; }}
    .status-row {{ display:flex; justify-content:space-between; padding:10px 0; border-bottom:1px solid var(--line); }}
    .footer-note {{ margin-top:16px; color:var(--muted); font-size:13px; }}
    .sections {{ display:grid; grid-template-columns:1.5fr 1fr; gap:14px; margin-top:14px; }}
    @media (max-width: 1100px) {{
      .metric-grid, .panel-grid, .sections {{ grid-template-columns:1fr; }}
      .top-layout {{ grid-template-columns:1fr; }}
      .mini-grid {{ grid-template-columns:1fr; }}
      .hero {{ flex-direction:column; align-items:stretch; }}
      .hero-actions {{ justify-content:flex-start; }}
    }}
  </style>
</head>
<body>
  <main class="page">
    <header class="hero">
      <div class="brand-lockup">
        <img class="brand-logo" src="gf_logo_transparent_2x.png" alt="GF Smart Accounting Solutions" />
      </div>
    </header>
    <div class="hero-tools">
      <span class="badge">شهر 6 | 1-13</span>
      <a class="primary-button" href="{details_href}">فتح ملف Excel</a>
    </div>

    <section class="metric-grid">
      {metric_card("إجمالي الربح", fmt_money(totals["total"]), "صافي الشحنات من 1 إلى 13")}
      {metric_card("ربح الشحنات", fmt_money(totals["base"]), "بعد الضريبة حسب شيت الأسعار")}
      {metric_card("ربح الوزن الزائد", fmt_money(totals["extra"]), "بعد 15 كجم")}
      {metric_card("ربح COD", fmt_money(totals["cod_profit"]), "الرسوم الصافية")}
    </section>

    <section class="top-layout">
      <div class="stack-panels">
        <article class="panel">
          <div class="panel-header">
            <div>
              <p class="eyebrow">شركات الشحن</p>
            </div>
          </div>
          <div class="carrier-donut-panel">
            <div>
              <div class="carrier-donut">
                <div class="carrier-donut-center">
                </div>
              </div>
            </div>
            <div class="carrier-legend">
              {carrier_legend_html}
            </div>
          </div>
        </article>

        <article class="panel shipments-panel" style="margin-top:-6px;">
          <div class="panel-header" style="margin-bottom:8px;">
            <div>
              <p class="eyebrow">الشحنات</p>
            </div>
          </div>
          <div class="mini-grid" style="grid-template-columns:repeat(3,minmax(0,1fr));">
            <div class="mini-card finance-card bank">
              <span class="metric-label">عدد الشحنات</span>
              <strong>{totals["active"]}</strong>
            </div>
            <div class="mini-card finance-card moyasar">
              <span class="metric-label">عدد COD</span>
              <strong>{totals["cod"]}</strong>
            </div>
            <div class="mini-card finance-card other">
              <span class="metric-label">مبلغ COD</span>
              <strong>{fmt_money(totals["cod_amount"])}</strong>
            </div>
          </div>
        </article>
      </div>

      <div class="stack-panels">
        <article class="panel">
          <div class="panel-header">
            <div>
              <p class="eyebrow">آخر 7 أيام</p>
            </div>
          </div>
          <div class="chart-block">
            <div class="area-chart" aria-label="Area Chart - Step">
              {chart_html}
            </div>
          </div>
          <div class="chart-block">
            <div class="area-chart" aria-label="مخطط عدد الشحنات اليومية">
              {count_chart_html}
            </div>
          </div>
        </article>
      </div>
    </section>

    <article class="panel">
      <div class="panel-header">
        <div>
          <p class="eyebrow">العمليات المالية</p>
        </div>
      </div>
      <div class="mini-grid" style="grid-template-columns:repeat(4,minmax(0,1fr));">
        <div class="mini-card finance-card bank">
          <span class="metric-label">إيداعات البنك</span>
          <strong>{finance['bank']['count']}</strong>
          <div class="metric-footer">{fmt_money(finance['bank']['total'])}</div>
        </div>
        <div class="mini-card finance-card moyasar">
          <span class="metric-label">إيداعات ميسر</span>
          <strong>{finance['moyasar']['count']}</strong>
          <div class="metric-footer">{fmt_money(finance['moyasar']['total'])}</div>
        </div>
        <div class="mini-card finance-card other">
          <span class="metric-label">تفصيل الخصومات والاستردادات</span>
          <div class="breakdown-line">خصم شحنة مرتجعة: {fmt_money(finance['shipping_return']['total'])}</div>
          <div class="breakdown-line">خصم الضريبة: {fmt_money(finance['tax_deduction']['total'])}</div>
          <div class="breakdown-line">استرداد تكلفة شحن: {fmt_money(finance['shipping_refund']['total'])}</div>
          <div class="metric-footer">الصافي: {fmt_money(finance['other_net'])}</div>
        </div>
        <div class="mini-card finance-card finance-total-card other">
          <span class="metric-label">إجمالي الإيداعات</span>
          <strong>{fmt_money(finance['total'])}</strong>
          <div class="metric-footer">البنك + ميسر</div>
        </div>
      </div>
    </article>

    <section class="panel statement-panel">
      <div class="panel-header">
        <div>
          <p class="eyebrow">كشف الحساب الجاري</p>
        </div>
      </div>
      <div class="statement-summary">
        <div class="mini-card quick-card">
          <span class="metric-label">إجمالي الإيداعات</span>
          <strong>{fmt_money(statement["summary"].get("deposits_total", 0))}</strong>
        </div>
        <div class="mini-card quick-card">
          <span class="metric-label">إجمالي المصاريف</span>
          <strong>{fmt_money(statement["summary"].get("expenses_total", 0))}</strong>
          <div class="metric-footer">فاتورة 12: 3,477.39 ريال | فاتورة 13: 984.14 ريال</div>
        </div>
        <div class="mini-card quick-card">
          <span class="metric-label">رسوم الحوالات</span>
          <strong>{fmt_money(statement["summary"].get("transfer_fees_total", 0))}</strong>
        </div>
        <div class="mini-card quick-card">
          <span class="metric-label">صافي الحركة</span>
          <strong>{fmt_money(statement["summary"].get("net_total", 0))}</strong>
        </div>
        <div class="mini-card quick-card">
          <span class="metric-label">عدد المصاريف</span>
          <strong>{int(statement["summary"].get("expenses_count", 0))}</strong>
        </div>
      </div>
      <div class="table-wrap statement-table">
        <table>
          <tbody>{statement_rows}</tbody>
        </table>
      </div>
    </section>

    <section class="panel finance-panel" style="margin-top:14px;">
          <div class="panel-header">
          <div>
            <p class="eyebrow">طلبات COD</p>
          </div>
        </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>رقم الطلب</th>
              <th>التاجر</th>
              <th>مبلغ COD</th>
              <th>تاريخ الطلب</th>
              <th>تاريخ التحصل</th>
              <th>تاريخ التحويل</th>
            </tr>
          </thead>
          <tbody>{cod_rows}</tbody>
        </table>
      </div>
    </section>

    <div class="footer-note">
      ملف Excel المرتب موجود هنا: <a href="{details_href}">{details_href}</a>
    </div>
  </main>
</body>
</html>"""


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    if os.path.exists(SOURCE):
        with open(REPORT_COPY, "wb") as dst, open(SOURCE, "rb") as src:
            dst.write(src.read())
    totals, top_merchants, top_carriers, statuses, daily, daily_count, finance, cod_items, statement = load_data()
    data = {
        "totals": totals,
        "daily": daily,
        "daily_count": daily_count,
        "finance": finance,
        "cod_items": cod_items,
        "top_carriers": top_carriers,
        "statement": statement,
    }
    with open(INDEX, "w", encoding="utf-8") as f:
        f.write(build_html(data))
    print(INDEX)


if __name__ == "__main__":
    main()
