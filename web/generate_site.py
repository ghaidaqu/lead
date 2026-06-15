import html
import json
import os
from collections import defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl


PROJECT_DIR = Path(__file__).resolve().parents[1]
SOURCE = PROJECT_DIR / "output" / "lead6_report.xlsx"
OUT_DIR = Path(__file__).resolve().parent
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


def june_period_label(dates):
    days = sorted({
        datetime.fromisoformat(str(date)).day
        for date in dates
        if date
    })
    if not days:
        return "يونيو"
    if days[-1] == 1:
        return "1 يونيو"
    return f"1 - {days[-1]} يونيو"


def load_data():
    wb = openpyxl.load_workbook(SOURCE, data_only=True)
    ws = wb["تفاصيل شهر 6"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    ops = wb["العمليات المالية"] if "العمليات المالية" in wb.sheetnames else None
    stmt = wb["كشف الحساب"] if "كشف الحساب" in wb.sheetnames else None

    items = []
    cod_items = []
    by_merchant = defaultdict(lambda: {"count": 0, "total": 0.0})
    by_city = defaultdict(lambda: {"count": 0, "total": 0.0})
    by_carrier = defaultdict(lambda: {"count": 0, "total": 0.0})
    by_status = defaultdict(int)
    by_date = defaultdict(float)
    by_date_count = defaultdict(int)
    all_shipment_dates = set()
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
        if item["date"]:
            all_shipment_dates.add(item["date"])
        if not item["included"]:
            continue
        items.append(item)
        if item["cod_amount"] > 0:
            cod_items.append(item)
        by_merchant[item["merchant"]]["count"] += 1
        by_merchant[item["merchant"]]["total"] += item["total_profit"]
        by_city[item["city"]]["count"] += 1
        by_city[item["city"]]["total"] += item["total_profit"]
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
            expense_type = clean(stmt.cell(r, 2).value)
            amount = money(stmt.cell(r, 3).value)
            if not any([date_val, expense_type, amount]):
                continue
            statement["rows"].append({
                "date": date_val.date().isoformat() if isinstance(date_val, datetime) else clean(date_val),
                "expense_type": expense_type,
                "amount": amount,
            })

    items.sort(key=lambda x: (x["date"], x["order_id"]), reverse=True)
    top_merchants = sorted(by_merchant.items(), key=lambda kv: kv[1]["total"], reverse=True)[:5]
    top_cities = sorted(by_city.items(), key=lambda kv: kv[1]["total"], reverse=True)[:5]
    top_carriers = sorted(by_carrier.items(), key=lambda kv: kv[1]["total"], reverse=True)
    daily = sorted(by_date.items(), key=lambda kv: kv[0])
    daily_count = sorted(by_date_count.items(), key=lambda kv: kv[0])
    period_label = june_period_label(all_shipment_dates)
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
    return totals, top_merchants, top_cities, top_carriers, by_status, daily, daily_count, finance, cod_items, statement, period_label


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
            f"<circle class='chart-point' cx='{x:.1f}' cy='{y:.1f}' r='5' fill='{color}' stroke='rgba(255,255,255,0.96)' stroke-width='2.2' />"
        )
        label_width = max(34, min(78, 12 + len(value_text) * 6.2))
        chosen_row = idx % 2
        label_y = 14 + (chosen_row * 20)
        side_offset = -18 if chosen_row == 0 else 18
        label_x = max(20, min(width - label_width - 20, x + side_offset - label_width / 2))
        line_y = label_y + 18
        value_labels.append(
            f"<g class='area-value-wrap'><line x1='{x:.1f}' y1='{y - 6:.1f}' x2='{x:.1f}' y2='{line_y:.1f}' stroke='rgba(177,42,91,0.22)' stroke-width='1' /><rect x='{label_x:.1f}' y='{label_y:.1f}' rx='10' ry='10' width='{label_width:.1f}' height='18' fill='rgba(15,11,19,0.88)' stroke='rgba(255,255,255,0.12)' stroke-width='1' /><text x='{label_x + label_width/2:.1f}' y='{label_y + 13:.1f}' text-anchor='middle' class='area-value'>{value_text}</text></g>"
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
        <path class='chart-line' d='{ " ".join(line_commands) }' fill='none' stroke='url(#areaStroke-{chart_key})' stroke-width='3.8' stroke-linecap='round' stroke-linejoin='round' />
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
    top_merchants = data["top_merchants"]
    top_cities = data["top_cities"]
    top_carriers = data["top_carriers"]
    period_label = data["period_label"]
    statement = data.get("statement", {"summary": {}, "rows": []})
    details_href = "/report.xlsx"
    cod_overrides = {
        "1757": {"collection_date": "2026-06-12", "transfer_date": "2026-06-13"},
        "1756": {"collection_date": "2026-06-11", "transfer_date": "2026-06-13"},
    }

    chart_html = build_step_area_chart(daily[-7:], color="#b12a5b", chart_key="profit")
    count_chart_html = build_step_area_chart(
        daily_count[-7:],
        color="#7a1538",
        value_formatter=lambda value: str(int(value)),
        chart_key="count",
    )
    carrier_total = sum(data["total"] for _, data in top_carriers) or 1
    carrier_colors = ["#b12a5b", "#7a1538", "#8c5f63", "#c09aa0", "#d8c2c8", "#5f5163"]
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
              <div class="carrier-row">
                <span class="carrier-dot" style="background:{color}"></span>
                <div class="carrier-name">{html.escape("ARAMEX - ارامكس" if name == "ارامكس - ARAMEX" else name)}</div>
                <div class="carrier-meta">{pct*100:.1f}%</div>
              </div>
            """
        )
        running_pct += pct
    carrier_legend_html = "".join(carrier_legend_items)
    carrier_donut_style = "; ".join([
        f"background: conic-gradient({', '.join(carrier_stops)})",
    ])

    merchant_rows = "".join(
        f"<tr><td>{idx + 1}</td><td>{html.escape(name or '-')}</td><td>{data['count']}</td><td>{fmt_money(data['total'])}</td></tr>"
        for idx, (name, data) in enumerate(top_merchants)
    )
    city_rows = "".join(
        f"<tr><td>{idx + 1}</td><td>{html.escape(name or '-')}</td><td>{data['count']}</td><td>{fmt_money(data['total'])}</td></tr>"
        for idx, (name, data) in enumerate(top_cities)
    )
    shipment_cards = f"""
      <section class="mini-band">
        <article class="mini-card">
          <span class="metric-label">عدد الشحنات</span>
          <strong>{totals["active"]}</strong>
          <div class="metric-footer">إجمالي الطلبات الداخلة في الربح</div>
        </article>
        <article class="mini-card">
          <span class="metric-label">عدد COD</span>
          <strong>{totals["cod"]}</strong>
          <div class="metric-footer">الطلبات التي عليها تحصيل</div>
        </article>
        <article class="mini-card">
          <span class="metric-label">مبلغ COD</span>
          <strong>{fmt_money(totals["cod_amount"])}</strong>
          <div class="metric-footer">إجمالي مبالغ COD</div>
        </article>
      </section>
    """
    finance_cards = f"""
      <section class="mini-band">
        <article class="mini-card finance-card bank">
          <span class="metric-label">إيداعات البنك</span>
          <strong>{finance['bank']['count']}</strong>
          <div class="metric-footer">{fmt_money(finance['bank']['total'])}</div>
        </article>
        <article class="mini-card finance-card moyasar">
          <span class="metric-label">إيداعات ميسر</span>
          <strong>{finance['moyasar']['count']}</strong>
          <div class="metric-footer">{fmt_money(finance['moyasar']['total'])}</div>
        </article>
        <article class="mini-card finance-card other">
          <span class="metric-label">تفصيل الخصومات والاستردادات</span>
          <div class="breakdown-line">خصم شحنة مرتجعة: {fmt_money(finance['shipping_return']['total'])}</div>
          <div class="breakdown-line">خصم الضريبة: {fmt_money(finance['tax_deduction']['total'])}</div>
          <div class="breakdown-line">استرداد تكلفة شحن: {fmt_money(finance['shipping_refund']['total'])}</div>
          <div class="metric-footer">الصافي: {fmt_money(finance['other_net'])}</div>
        </article>
        <article class="mini-card finance-card finance-total-card other">
          <span class="metric-label">إجمالي الإيداعات</span>
          <strong>{fmt_money(finance['total'])}</strong>
          <div class="metric-footer">البنك + ميسر</div>
        </article>
      </section>
    """
    statement_rows = "".join(
        f"<tr class='expense-row'><td>{row['date']}</td><td>{row['expense_type'] or '-'}</td><td>{fmt_money(row['amount'])}</td></tr>"
        for row in statement["rows"]
    )
    cod_rows = "".join(
        f"<tr><td>{item['order_id']}</td><td>{html.escape(item['merchant'])}</td><td>{fmt_money(item['cod_amount'])}</td><td>{item.get('date') or '-'}</td><td>{(cod_overrides.get(str(item['order_id'])) or {}).get('collection_date') or item.get('date') or '-'}</td><td>{(cod_overrides.get(str(item['order_id'])) or {}).get('transfer_date') or '-'}</td></tr>"
        for item in cod_items
    )
    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>lead</title>
  <style>
    :root {{
      color-scheme: dark;
      --bg: #07070b;
      --bg-soft: #0f0b13;
      --card: rgba(255,255,255,.075);
      --card-strong: rgba(255,255,255,.11);
      --border: rgba(255,255,255,.14);
      --text: #f7f2f5;
      --muted: #a89ca7;
      --burgundy: #7a1538;
      --burgundy-2: #b12a5b;
      --purple: #6d28d9;
      --green: #2dd4bf;
      --red: #fb7185;
      --shadow: 0 24px 80px rgba(0,0,0,.45);
    }}
    * {{ box-sizing: border-box; }}
    html {{ scroll-behavior:smooth; }}
    body {{
      margin:0;
      min-height:100vh;
      color:var(--text);
      font-family: Inter, "Avenir Next", "Segoe UI", "Noto Sans Arabic", system-ui, sans-serif;
      background:
        radial-gradient(circle at 15% 5%, rgba(122,21,56,.42), transparent 28%),
        radial-gradient(circle at 90% 0%, rgba(109,40,217,.18), transparent 24%),
        linear-gradient(135deg, #050509 0%, #0b0810 50%, #130812 100%);
    }}
    a {{ color:inherit; text-decoration:none; }}
    .page {{ width:min(1180px, calc(100% - 32px)); margin:0 auto; padding:20px 0 34px; }}
    .hero {{
      margin:0 0 8px;
      padding:16px 28px 14px;
      border-radius:26px;
      background:
        linear-gradient(90deg, #050509 0%, #150811 35%, #4a0c2b 70%, #b1125a 100%);
      border:1px solid rgba(255,255,255,.16);
      box-shadow:var(--shadow);
      backdrop-filter: blur(18px);
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:18px;
      flex-wrap:wrap;
      direction:ltr;
      position:relative;
      overflow:hidden;
    }}
    .hero::before {{
      content:'';
      position:absolute;
      inset:0;
      background:
        radial-gradient(circle at 72% 22%, rgba(255,255,255,.08), transparent 18%),
        radial-gradient(circle at 88% 20%, rgba(255,255,255,.06), transparent 16%),
        linear-gradient(100deg, transparent 0 56%, rgba(255,255,255,.02) 56% 100%);
      pointer-events:none;
    }}
    .hero::after {{
      content:'';
      position:absolute;
      inset:0;
      background:
        repeating-linear-gradient(110deg, rgba(255,255,255,.03) 0 1px, transparent 1px 26px);
      opacity:.18;
      mask-image: linear-gradient(90deg, transparent 0 52%, black 72% 100%);
      pointer-events:none;
    }}
    .brand-lockup {{ position:relative; z-index:1; display:flex; align-items:center; justify-content:flex-start; gap:0; padding:0; border:0; background:transparent; box-shadow:none; direction:ltr; }}
    .brand-logo {{ width:min(176px, 18vw); max-width:176px; height:auto; display:block; background:transparent; filter: drop-shadow(0 10px 24px rgba(0,0,0,.18)); }}
    .brand-divider {{ width:1px; height:52px; background:rgba(255,255,255,.34); box-shadow:0 0 0 1px rgba(0,0,0,.05); margin:0 10px 0 -2px; }}
    .hero-copy {{ display:grid; gap:0; justify-items:flex-start; margin-inline-start:-1px; }}
    .hero-copy span {{ display:block; margin:0; color:rgba(255,255,255,.92); font-size:clamp(14px, 1.6vw, 18px); line-height:1.0; font-weight:650; letter-spacing:-0.02em; }}
    .hero-copy .brand-tag {{ font-size:clamp(14px, 1.6vw, 18px); }}
    .hero-title {{
      position:absolute;
      right:40px;
      top:50%;
      transform:translateY(-50%);
      z-index:1;
      direction:rtl;
      text-align:right;
      display:flex;
      flex-direction:column;
      align-items:flex-end;
      justify-content:center;
      gap:3px;
      margin:0;
      padding:0;
      padding-left:0;
      padding-inline-start:0;
      padding-inline-end:0;
      min-width:min(20vw, 300px);
      width:max-content;
      flex:0 0 auto;
    }}
    .hero-title h1 {{
      margin:0;
      font-size:clamp(18px, 2.0vw, 23px);
      line-height:1.08;
      letter-spacing:-0.03em;
      font-weight:800;
      color:#fff;
      text-shadow:0 2px 18px rgba(0,0,0,.28);
    }}
    .hero-title p {{
      margin:0;
      color:rgba(255,255,255,.72);
      font-size:11px;
    }}
    .hero-tools {{
      display:flex;
      align-items:center;
      justify-content:flex-end;
      gap:10px;
      flex-wrap:wrap;
      margin:0 0 10px;
    }}
    .badge {{
      display:inline-flex;
      align-items:center;
      min-height:40px;
      padding:0 14px;
      border-radius:999px;
      border:1px solid rgba(255,255,255,.16);
      background:rgba(255,255,255,.08);
      color:rgba(255,255,255,.84);
      font-size:12px;
      backdrop-filter: blur(14px);
    }}
    .primary-button {{
      display:inline-flex;
      align-items:center;
      min-height:38px;
      padding:0 14px;
      border-radius:14px;
      background:
        linear-gradient(135deg, rgba(122,21,56,.98), rgba(177,42,91,.84)),
        linear-gradient(180deg, rgba(255,255,255,.12), rgba(255,255,255,.02));
      color:#fff;
      border:1px solid rgba(255,255,255,.20);
      font-weight:780;
      letter-spacing:-0.01em;
      box-shadow:
        0 14px 28px rgba(122,21,56,.30),
        inset 0 1px 0 rgba(255,255,255,.18);
      backdrop-filter: blur(14px);
      -webkit-backdrop-filter: blur(14px);
    }}
    .primary-button:hover {{ transform: translateY(-1px); filter: brightness(1.05); }}
    .subhero {{
      display:flex;
      justify-content:flex-end;
      margin:0 0 14px;
    }}
    .metric-grid {{ display:grid; grid-template-columns:repeat(4, minmax(0,1fr)); gap:14px; margin-bottom:14px; }}
    .mini-band {{ display:grid; grid-template-columns:repeat(3, minmax(0,1fr)); gap:14px; margin:0 0 16px; }}
    .metric-card, .panel, .mini-card {{
      position:relative;
      overflow:hidden;
      border-radius:24px;
      background:linear-gradient(180deg, rgba(18,13,25,.96), rgba(11,9,16,.92));
      border:1px solid rgba(255,255,255,.16);
      box-shadow:0 18px 50px rgba(0,0,0,.45);
      backdrop-filter: blur(22px) saturate(122%);
      -webkit-backdrop-filter: blur(22px) saturate(122%);
    }}
    .metric-card::before, .panel::before, .mini-card::before {{
      content:'';
      position:absolute;
      inset:0;
      background:
        linear-gradient(135deg, rgba(255,255,255,.13) 0%, rgba(255,255,255,.04) 18%, rgba(255,255,255,0) 42%),
        radial-gradient(circle at 18% 10%, rgba(177,42,91,.12), transparent 28%);
      pointer-events:none;
    }}
    .metric-card > *, .panel > *, .mini-card > * {{ position:relative; z-index:1; }}
    .mini-card {{
      padding:14px;
      min-height:92px;
      display:flex;
      flex-direction:column;
      justify-content:center;
      align-items:center;
      text-align:center;
      gap:2px;
    }}
    .metric-card {{
      padding:20px;
      min-height:120px;
      display:flex;
      flex-direction:column;
      justify-content:space-between;
    }}
    .metric-label {{ display:block; color:var(--muted); font-size:13px; line-height:1.45; }}
    .metric-card strong {{ display:block; margin-top:10px; font-size:24px; line-height:1; font-weight:850; letter-spacing:-0.04em; font-variant-numeric: tabular-nums; color:#fff; text-shadow:0 2px 18px rgba(255,255,255,.16); }}
    .metric-footer {{ margin-top:8px; color:rgba(255,255,255,.58); font-size:12px; line-height:1.35; }}
    .mini-card .metric-label {{ font-size:13px; line-height:1.35; }}
    .mini-card strong {{ display:block; margin-top:6px; font-size:20px; line-height:1.15; font-variant-numeric: tabular-nums; color:#fff; text-shadow:0 2px 16px rgba(255,255,255,.14); }}
    .mini-card .metric-footer {{ font-size:12px; line-height:1.35; }}
    .content-grid {{ display:grid; grid-template-columns:1.3fr .92fr; gap:16px; margin-bottom:16px; }}
    .panel {{ padding:20px; }}
    .panel-header {{ display:flex; align-items:flex-start; justify-content:space-between; gap:14px; margin-bottom:16px; }}
    .eyebrow {{ margin:0; color:rgba(255,255,255,.88); font-size:13px; font-weight:750; letter-spacing:.02em; }}
    .subtle {{ margin:6px 0 0; color:var(--muted); font-size:13px; }}
    .chart-shell {{ height:300px; }}
    .area-chart {{ width:100%; height:100%; display:block; }}
    .area-chart svg {{ width:100%; height:100%; display:block; overflow:visible; }}
    .area-value {{ fill:#fff; font-size:11px; font-weight:850; font-family:Inter, "Avenir Next", "Segoe UI", "Noto Sans Arabic", system-ui, sans-serif; }}
    .area-date {{ fill:rgba(255,255,255,.56); font-size:10px; font-weight:700; font-family:Inter, "Avenir Next", "Segoe UI", "Noto Sans Arabic", system-ui, sans-serif; }}
    .chart-line {{ filter: drop-shadow(0 2px 6px rgba(177,42,91,.18)); }}
    .chart-point {{ filter: drop-shadow(0 2px 5px rgba(0,0,0,.18)); }}
    .carrier-layout {{
      display:grid;
      grid-template-columns:120px 1fr;
      gap:16px;
      align-items:center;
    }}
    .carrier-donut {{
      width:120px;
      aspect-ratio:1;
      border-radius:50%;
      position:relative;
      margin:0 auto;
      padding:10px;
      {carrier_donut_style};
      box-shadow:0 24px 40px rgba(0,0,0,.36), inset 0 1px 0 rgba(255,255,255,.14);
    }}
    .carrier-donut::after {{
      content:'';
      position:absolute;
      inset:22%;
      border-radius:50%;
      background:linear-gradient(180deg, rgba(15,11,19,.98), rgba(7,7,11,.96));
      border:1px solid rgba(255,255,255,.08);
      box-shadow:inset 0 1px 0 rgba(255,255,255,.06);
    }}
    .carrier-center {{
      position:absolute;
      inset:0;
      display:flex;
      align-items:center;
      justify-content:center;
      z-index:1;
      text-align:center;
      padding:18px;
      color:#fff;
      pointer-events:none;
    }}
    .carrier-center span {{ display:block; font-size:11px; color:rgba(255,255,255,.72); }}
    .carrier-center strong {{ display:block; font-size:18px; line-height:1; font-weight:850; }}
    .carrier-legend {{ display:grid; gap:10px; }}
    .carrier-row {{
      display:grid;
      grid-template-columns:12px 1fr auto;
      gap:10px;
      align-items:center;
      padding:12px 14px;
      border-radius:16px;
      border:1px solid rgba(255,255,255,.10);
      background:rgba(255,255,255,.045);
    }}
    .carrier-dot {{ width:12px; height:12px; border-radius:50%; box-shadow:0 0 0 4px rgba(255,255,255,.04); }}
    .carrier-name {{ font-size:13px; font-weight:700; color:var(--text); }}
    .carrier-meta {{ font-size:12px; color:rgba(255,255,255,.62); white-space:nowrap; }}
    .table-grid {{ display:grid; grid-template-columns:1fr 1fr; gap:16px; }}
    .table-wrap {{ overflow:auto; }}
    table {{ width:100%; border-collapse:collapse; }}
    th, td {{
      padding:12px 10px;
      border-bottom:1px solid rgba(255,255,255,.08);
      text-align:right;
      white-space:nowrap;
    }}
    th {{
      color:var(--muted);
      font-size:12px;
      font-weight:700;
    }}
    td {{
      color:rgba(255,255,255,.88);
      font-size:13px;
    }}
    .rank {{
      width:36px;
      color:rgba(255,255,255,.56);
      font-variant-numeric: tabular-nums;
    }}
    .table-card-title {{
      display:flex;
      align-items:flex-start;
      justify-content:space-between;
      gap:12px;
      margin-bottom:14px;
    }}
    .table-card-title p {{ margin:6px 0 0; color:var(--muted); font-size:12px; }}
    .statement-summary {{ display:grid; grid-template-columns:repeat(5,minmax(0,1fr)); gap:12px; margin-bottom:14px; }}
    .statement-summary .mini-card strong {{ font-size:18px; }}
    .statement-panel {{ margin-top:16px; }}
    .finance-card strong {{ font-size:18px; }}
    .finance-total-card strong {{ font-size:16px; }}
    .breakdown-line {{ color:rgba(255,255,255,.70); font-size:12px; line-height:1.5; margin-top:2px; }}
    .footer-note {{ margin-top:16px; color:rgba(255,255,255,.48); font-size:12px; }}
    @media (max-width: 1100px) {{
      .metric-grid, .content-grid, .table-grid, .mini-band, .statement-summary {{ grid-template-columns:1fr; }}
      .carrier-layout {{ grid-template-columns:1fr; }}
      .hero {{ padding:14px 18px 12px; margin-bottom:6px; align-items:flex-start; }}
      .brand-lockup {{ flex:1 1 100%; min-width:0; }}
      .brand-logo {{ width:min(150px, 38vw); }}
      .brand-divider {{ height:42px; margin-inline:8px; }}
      .hero-title {{ position:relative; inset:auto; transform:none; min-width:0; padding:0; align-items:flex-end; width:100%; flex:1 1 100%; }}
      .hero-copy span {{ font-size:14px; }}
      .hero-title h1 {{ font-size:18px; }}
      .hero-title p {{ font-size:10px; }}
    }}
  </style>
</head>
<body>
  <main class="page">
    <header class="hero">
      <div class="brand-lockup">
        <img class="brand-logo" src="gf_logo_current_clean.png" alt="GF Smart Accounting Solutions" />
        <div class="brand-divider" aria-hidden="true"></div>
        <div class="hero-copy">
          <span class="brand-tag">Smart Accounting</span>
          <span>Solutions</span>
        </div>
      </div>
      <div class="hero-title">
        <h1>لوحة التحكم المالية المختصرة</h1>
        <p>نظرة عامة على أداء الأعمال</p>
      </div>
    </header>

    <div class="subhero">
      <div class="hero-tools">
        <span class="badge">{period_label}</span>
        <a class="primary-button" href="{details_href}"><span>فتح ملف Excel</span></a>
      </div>
    </div>

    <section class="metric-grid">
      {metric_card("إجمالي الربح", fmt_money(totals["total"]), f"صافي الشحنات: {period_label}")}
      {metric_card("ربح الشحنات", fmt_money(totals["base"]), "بعد الضريبة حسب شيت الأسعار")}
      {metric_card("ربح الوزن الزائد", fmt_money(totals["extra"]), "بعد 15 كجم")}
      {metric_card("ربح COD", fmt_money(totals["cod_profit"]), "الرسوم الصافية")}
    </section>

    {shipment_cards}

    <section class="content-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <p class="eyebrow">الأداء اليومي</p>
            <p class="subtle">آخر 7 أيام</p>
          </div>
        </div>
        <div class="chart-shell">
          <div class="area-chart" aria-label="Area Chart - Step">
            {chart_html}
          </div>
        </div>
        <div class="chart-shell" style="margin-top:16px; border-top:1px solid rgba(255,255,255,.08); padding-top:16px;">
          <div class="area-chart" aria-label="مخطط عدد الشحنات اليومية">
            {count_chart_html}
          </div>
        </div>
      </article>

      <article class="panel">
        <div class="panel-header">
          <div>
            <p class="eyebrow">شركات الشحن</p>
            <p class="subtle">نسبة الربح لكل ناقل</p>
          </div>
        </div>
        <div class="carrier-layout">
          <div class="carrier-donut">
            <div class="carrier-center">
              <div>
                <strong>{len(top_carriers)}</strong>
                <span>شركات</span>
              </div>
            </div>
          </div>
          <div class="carrier-legend">
            {carrier_legend_html}
          </div>
        </div>
      </article>
    </section>

    <section class="table-grid">
      <article class="panel">
        <div class="table-card-title">
          <div>
            <p class="eyebrow">أفضل العملاء</p>
            <p class="subtle">حسب إجمالي الربح</p>
          </div>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr><th class="rank">#</th><th>العميل</th><th>عدد</th><th>الربح</th></tr>
            </thead>
            <tbody>{merchant_rows}</tbody>
          </table>
        </div>
      </article>

      <article class="panel">
        <div class="table-card-title">
          <div>
            <p class="eyebrow">أفضل المدن</p>
            <p class="subtle">حسب إجمالي الربح</p>
          </div>
        </div>
        <div class="table-wrap">
          <table>
            <thead>
              <tr><th class="rank">#</th><th>المدينة</th><th>عدد</th><th>الربح</th></tr>
            </thead>
            <tbody>{city_rows}</tbody>
          </table>
        </div>
      </article>
    </section>

    <section class="panel" style="margin-top:16px;">
      <div class="panel-header">
        <div>
          <p class="eyebrow">العمليات المالية</p>
          <p class="subtle">إيداعات ومخصومات واستردادات</p>
        </div>
      </div>
      {finance_cards}
    </section>

    <section class="panel statement-panel">
      <div class="panel-header">
        <div>
          <p class="eyebrow">كشف الحساب الجاري</p>
          <p class="subtle">ملخص الحركات من الشيت</p>
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
          <thead>
            <tr><th>التاريخ</th><th>نوع المصروف</th><th>المبلغ</th></tr>
          </thead>
          <tbody>{statement_rows}</tbody>
        </table>
      </div>
    </section>

    <section class="panel" style="margin-top:16px;">
      <div class="panel-header">
        <div>
          <p class="eyebrow">طلبات COD</p>
          <p class="subtle">كل طلب على حدة</p>
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
      ملف Excel المرتب موجود هنا: <a href="{details_href}">lead_report.xlsx</a>
    </div>
  </main>
</body>
</html>"""


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    totals, top_merchants, top_cities, top_carriers, statuses, daily, daily_count, finance, cod_items, statement, period_label = load_data()
    data = {
        "totals": totals,
        "top_merchants": top_merchants,
        "top_cities": top_cities,
        "daily": daily,
        "daily_count": daily_count,
        "finance": finance,
        "cod_items": cod_items,
        "top_carriers": top_carriers,
        "statement": statement,
        "period_label": period_label,
    }
    with open(INDEX, "w", encoding="utf-8") as f:
        f.write(build_html(data))
    print(INDEX)


if __name__ == "__main__":
    main()
