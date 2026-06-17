import html
import json
import os
import re
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
    return f"1-{days[-1]} يونيو"


SPECIAL_MERCHANTS = {"عبدالرحمن المطيري", "مؤسسة اواني القيصرية التجارية"}
CARRIER_PRICE_NAMES = {
    "ارامكس - ARAMEX": "ارامكس",
    "aramex ( استلام من الفرع )": "ارامكس استلام",
    "SMSA - سمسا": "سمسا",
    "SMSA ( استلام من الفرع )": "سمسا استلام",
    "RedBox - ريدبوكس": "ريد بوكس",
    "JT Express": "JT Express",
}
RETURN_DETAIL_OVERRIDES = {
    "767": {
        "merchant": "مؤسسة اواني القيصرية التجارية",
        "carrier": "SMSA - سمسا",
        "weight": 15.0,
    },
    "1376": {
        "merchant": "عبدالرحمن المطيري",
        "carrier": "ارامكس - ARAMEX",
        "weight": 1.0,
    },
    "1389": {
        "merchant": "ليد إكسبرس",
        "carrier": "ارامكس - ARAMEX",
        "weight": 1.0,
    },
    "1256": {
        "merchant": "ليد إكسبرس",
        "carrier": "ارامكس - ARAMEX",
        "weight": 22.0,
    },
    "1315": {
        "merchant": "عبدالرحمن المطيري",
        "carrier": "SMSA - سمسا",
        "weight": 1.0,
    },
    "1356": {
        "merchant": "ليد إكسبرس",
        "carrier": "ارامكس - ARAMEX",
        "weight": 10.0,
    },
}


def price_sheet_headers(ws):
    if ws is None:
        return {}
    return {clean(ws.cell(1, col).value): col for col in range(1, ws.max_column + 1)}


def platform_cost_net_from_prices(prices_ws, headers, merchant, carrier, fallback=0.0):
    if prices_ws is None:
        return fallback
    if merchant in SPECIAL_MERCHANTS and merchant in headers:
        return money(prices_ws.cell(5, headers[merchant]).value) or fallback
    carrier_key = CARRIER_PRICE_NAMES.get(carrier)
    if carrier_key and carrier_key in headers:
        return money(prices_ws.cell(5, headers[carrier_key]).value) or fallback
    return fallback


def infer_return_profit_from_prices(prices_ws, headers, revenue):
    if prices_ws is None:
        revenue_net = revenue * 0.85
        return {
            "merchant": "مقدر",
            "carrier": "مقدر",
            "weight": 0.0,
            "revenue_net": revenue_net,
            "platform_shipping": 0.0,
            "base_profit": revenue_net,
            "extra_profit": 0.0,
            "total_profit": revenue_net,
            "status": "مقدر بدون شيت الأسعار",
        }

    extra_col = headers.get("السعر لكل كيلو  زيادة")
    extra_customer_gross = money(prices_ws.cell(2, extra_col).value) if extra_col else 0.0
    extra_customer_net = money(prices_ws.cell(3, extra_col).value) if extra_col else 0.0
    extra_platform_gross = money(prices_ws.cell(4, extra_col).value) if extra_col else 0.0
    extra_platform_net = money(prices_ws.cell(5, extra_col).value) if extra_col else 0.0
    if not extra_customer_net and extra_customer_gross:
        extra_customer_net = extra_customer_gross * 0.85
    if not extra_platform_net and extra_platform_gross:
        extra_platform_net = extra_platform_gross * 0.85

    exact_candidates = []
    fallback_candidates = []
    skip_names = {"السعر لكل كيلو  زيادة", "سعر توصيل cod", ""}
    for name, col in headers.items():
        if name in skip_names:
            continue
        customer_gross = money(prices_ws.cell(2, col).value)
        platform_net = money(prices_ws.cell(5, col).value)
        if not customer_gross or not platform_net or revenue + 0.001 < customer_gross:
            continue

        customer_net = money(prices_ws.cell(3, col).value)
        if not customer_net:
            customer_net = customer_gross if name in SPECIAL_MERCHANTS else customer_gross * 0.85

        extra_charge = max(revenue - customer_gross, 0.0)
        extra_kg = extra_charge / extra_customer_gross if extra_customer_gross else 0.0
        revenue_net = customer_net + (extra_kg * extra_customer_net)
        platform_shipping = platform_net + (extra_kg * extra_platform_net)
        base_profit = customer_net - platform_net
        extra_profit = extra_kg * (extra_customer_net - extra_platform_net)
        total_profit = revenue_net - platform_shipping
        row = {
            "merchant": name if name in SPECIAL_MERCHANTS else "مقدر",
            "carrier": name if name not in SPECIAL_MERCHANTS else "مقدر",
            "weight": 15 + extra_kg if extra_kg else 0.0,
            "revenue_net": revenue_net,
            "platform_shipping": platform_shipping,
            "base_profit": base_profit,
            "extra_profit": extra_profit,
            "total_profit": total_profit,
            "status": f"مقدر حسب {name}",
        }
        if abs(revenue - customer_gross) < 0.001:
            exact_candidates.append(row)
        else:
            fallback_candidates.append(row)

    candidates = exact_candidates or fallback_candidates
    if not candidates:
        revenue_net = revenue * 0.85
        return {
            "merchant": "مقدر",
            "carrier": "غير موجود في شيت الأسعار",
            "weight": 0.0,
            "revenue_net": revenue_net,
            "platform_shipping": 0.0,
            "base_profit": revenue_net,
            "extra_profit": 0.0,
            "total_profit": revenue_net,
            "status": "مقدر بدون مطابقة سعر",
        }
    return min(candidates, key=lambda item: item["total_profit"])


def return_profit_from_details(prices_ws, headers, revenue, merchant, carrier, weight):
    if prices_ws is None:
        return infer_return_profit_from_prices(prices_ws, headers, revenue)

    extra_col = headers.get("السعر لكل كيلو  زيادة")
    extra_customer_gross = money(prices_ws.cell(2, extra_col).value) if extra_col else 0.0
    extra_platform_gross = money(prices_ws.cell(4, extra_col).value) if extra_col else 0.0
    extra_platform_net = money(prices_ws.cell(5, extra_col).value) if extra_col else 0.0
    if not extra_platform_net and extra_platform_gross:
        extra_platform_net = extra_platform_gross * 0.85

    price_key = merchant if merchant in SPECIAL_MERCHANTS and merchant in headers else CARRIER_PRICE_NAMES.get(carrier, carrier)
    col = headers.get(price_key)
    if col is None:
        inferred = infer_return_profit_from_prices(prices_ws, headers, revenue)
        inferred.update({
            "merchant": merchant,
            "carrier": carrier,
            "weight": weight,
            "status": "محسوب بمطابقة سعر تقديرية",
        })
        return inferred

    base_customer_gross = money(prices_ws.cell(2, col).value)
    platform_base_net = money(prices_ws.cell(5, col).value)
    extra_kg = max(weight - 15, 0.0)
    extra_charge_gross = max(revenue - base_customer_gross, 0.0)
    extra_charge_net = extra_charge_gross if merchant in SPECIAL_MERCHANTS else extra_charge_gross * 0.85
    revenue_net = revenue if merchant in SPECIAL_MERCHANTS else revenue * 0.85
    platform_extra_net = extra_kg * extra_platform_net
    platform_shipping = platform_base_net + platform_extra_net
    total_profit = revenue_net - platform_shipping
    extra_profit = extra_charge_net - platform_extra_net
    base_profit = total_profit - extra_profit
    return {
        "merchant": merchant,
        "carrier": carrier,
        "weight": weight,
        "revenue_net": revenue_net,
        "platform_shipping": platform_shipping,
        "base_profit": base_profit,
        "extra_profit": extra_profit,
        "total_profit": total_profit,
        "status": "محسوب من المحفظة والشحنات",
    }


def load_data():
    wb = openpyxl.load_workbook(SOURCE, data_only=True)
    ws = wb["تفاصيل شهر 6"]
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    ops = wb["العمليات المالية"] if "العمليات المالية" in wb.sheetnames else None
    stmt = wb["كشف الحساب"] if "كشف الحساب" in wb.sheetnames else None
    prices_ws = wb["الاسعار"] if "الاسعار" in wb.sheetnames else None
    prices_headers = price_sheet_headers(prices_ws)

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
    total_customer_shipping = 0.0
    total_records = 0
    all_items_by_order = {}
    return_items = []

    for row in rows:
        if not any(v is not None for v in row):
            continue
        if row[0] in (None, ""):
            continue
        total_records += 1
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
            "customer_gross": money(row[15]),
            "customer_net": money(row[16]),
            "platform_shipping": money(row[17]),
            "shipping_profit": money(row[19]),
            "extra_profit": money(row[20]),
            "fee_profit": money(row[21]),
            "total_profit": money(row[22]),
            "review_diff": money(row[19]),
        }
        if item["order_id"] not in (None, ""):
            total_customer_shipping += item["customer_shipping"]
            all_items_by_order[str(item["order_id"])] = item
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
            if typ == "إيداع":
                key = "bank" if "تحويل بنكي" in note else "moyasar" if "Moyasar" in note else "other"
                finance[key]["count"] += 1
                finance[key]["total"] += amount
            else:
                if typ == "admin_deduction" and "خصم تكلفة شحنة مرتجعة" in note:
                    finance["shipping_return"]["count"] += 1
                    finance["shipping_return"]["total"] += abs(amount)
                    order_match = re.search(r"طلب\s*#?\s*(\d+)", note)
                    order_id = order_match.group(1) if order_match else ""
                    linked_item = all_items_by_order.get(order_id)
                    revenue = abs(amount)
                    if linked_item is not None:
                        net_ratio = (
                            linked_item["customer_net"] / linked_item["customer_gross"]
                            if linked_item["customer_gross"]
                            else 0.85
                        )
                        revenue_net = revenue * net_ratio
                        platform_cost_net = platform_cost_net_from_prices(
                            prices_ws,
                            prices_headers,
                            linked_item["merchant"],
                            linked_item["carrier"],
                            linked_item["platform_shipping"],
                        )
                        base_profit = revenue_net - platform_cost_net
                        extra_profit = linked_item["extra_profit"]
                        total_profit = base_profit + extra_profit
                        return_items.append({
                            "order_id": order_id,
                            "merchant": linked_item["merchant"],
                            "carrier": linked_item["carrier"],
                            "weight": linked_item["weight"],
                            "revenue": revenue,
                            "revenue_net": revenue_net,
                            "platform_shipping": platform_cost_net,
                            "base_profit": base_profit,
                            "extra_profit": extra_profit,
                            "total_profit": total_profit,
                            "matched": True,
                            "status": "محسوب",
                        })
                    else:
                        override = RETURN_DETAIL_OVERRIDES.get(order_id)
                        if override:
                            inferred = return_profit_from_details(
                                prices_ws,
                                prices_headers,
                                revenue,
                                override["merchant"],
                                override["carrier"],
                                override["weight"],
                            )
                        else:
                            inferred = infer_return_profit_from_prices(prices_ws, prices_headers, revenue)
                        return_items.append({
                            "order_id": order_id or "-",
                            "merchant": inferred["merchant"],
                            "carrier": inferred["carrier"],
                            "weight": inferred["weight"],
                            "revenue": revenue,
                            "revenue_net": inferred["revenue_net"],
                            "platform_shipping": inferred["platform_shipping"],
                            "base_profit": inferred["base_profit"],
                            "extra_profit": inferred["extra_profit"],
                            "total_profit": inferred["total_profit"],
                            "matched": False,
                            "status": inferred["status"],
                        })
                elif typ == "admin_deduction" and "خصم ضريبة القيمة المضافة" in note:
                    finance["other_net"] += amount
                    finance["tax_deduction"]["count"] += 1
                    finance["tax_deduction"]["total"] += abs(amount)
                elif typ == "استرداد" and "استرداد تكلفة شحن" in note:
                    finance["other_net"] += amount
                    finance["shipping_refund"]["count"] += 1
                    finance["shipping_refund"]["total"] += amount
                else:
                    finance["other_net"] += amount
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
    return_revenue = sum(item["revenue"] for item in return_items)
    return_profit = sum(item["total_profit"] for item in return_items)
    totals = {
        "records": total_records,
        "active": len(items),
        "shipping": len(items),
        "cod": sum(1 for item in items if item["fee_profit"] > 0),
        "cod_amount": sum(item["cod_amount"] for item in items if item["included"]),
        "revenue": total_customer_shipping + return_revenue,
        "base": sum(item["shipping_profit"] for item in items),
        "extra": sum(item["extra_profit"] for item in items),
        "cod_profit": sum(item["fee_profit"] for item in items),
        "return_revenue": return_revenue,
        "return_profit": return_profit,
        "return_count": len(return_items),
        "total": sum(item["total_profit"] for item in items) + return_profit,
        "excluded": total_records - len(items),
    }
    cod_items.sort(key=lambda x: (x["date"], x["order_id"]), reverse=True)
    return_items.sort(key=lambda x: (x["matched"], x["order_id"]), reverse=True)
    return totals, top_merchants, top_cities, top_carriers, by_status, daily, daily_count, finance, cod_items, return_items, statement, period_label


def fmt_money(value):
    return f"{value:,.2f} ريال"


def metric_card(label, value, note="", tone="navy"):
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
    return_items = data["return_items"]
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

    daily_profit_points = daily[-7:]
    daily_count_map = {date: count for date, count in daily_count}
    daily_chart_labels = [
        str(date)[5:] if len(str(date)) >= 10 else str(date)
        for date, _ in daily_profit_points
    ]
    daily_chart_profit = [round(float(value), 2) for _, value in daily_profit_points]
    daily_chart_count = [
        int(daily_count_map.get(date, 0))
        for date, _ in daily_profit_points
    ]
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
        for idx, (name, data) in enumerate(top_merchants[:4])
    )
    city_rows = "".join(
        f"<tr><td>{idx + 1}</td><td>{html.escape(name or '-')}</td><td>{data['count']}</td><td>{fmt_money(data['total'])}</td></tr>"
        for idx, (name, data) in enumerate(top_cities[:4])
    )
    profit_details = f"""
      <section class="profit-details-card">
        <article class="panel details-panel">
          <div class="panel-header compact-header">
            <div>
              <p class="eyebrow">تفاصيل الأرباح</p>
              <p class="subtle">تفصيل القيم الداخلة في إجمالي الربح</p>
            </div>
          </div>
          <div class="details-grid">
            <div class="detail-chip"><span>ربح الشحنات</span><strong>{fmt_money(totals["base"])}</strong></div>
            <div class="detail-chip"><span>ربح المرتجعات</span><strong>{fmt_money(totals["return_profit"])}</strong></div>
            <div class="detail-chip"><span>ربح الوزن الزائد</span><strong>{fmt_money(totals["extra"])}</strong></div>
            <div class="detail-chip"><span>ربح COD</span><strong>{fmt_money(totals["cod_profit"])}</strong></div>
            <div class="detail-chip"><span>عدد COD</span><strong>{totals["cod"]}</strong></div>
          </div>
        </article>
      </section>
    """
    finance_cards = f"""
      <section class="mini-band finance-band">
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
          <div class="breakdown-line">خصم الضريبة: {fmt_money(finance['tax_deduction']['total'])}</div>
          <div class="breakdown-line">العمليات الملغية: {fmt_money(finance['shipping_refund']['total'])}</div>
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

    def cod_collection_date(item):
        override = cod_overrides.get(str(item["order_id"])) or {}
        if override.get("collection_date"):
            return override["collection_date"]
        if item.get("date") == "2026-06-15":
            return "-"
        return item.get("date") or "-"

    cod_rows = "".join(
        f"<tr><td>{item['order_id']}</td><td>{html.escape(item['merchant'])}</td><td>{fmt_money(item['cod_amount'])}</td><td>{item.get('date') or '-'}</td><td>{cod_collection_date(item)}</td><td>{(cod_overrides.get(str(item['order_id'])) or {}).get('transfer_date') or '-'}</td></tr>"
        for item in cod_items
    )
    return_rows = "".join(
        f"<tr><td>{item['order_id']}</td><td>{html.escape(item['merchant'])}</td><td>{html.escape(item['carrier'])}</td><td>{item['weight']:.2f} كجم</td><td>{fmt_money(item['revenue'])}</td><td>{fmt_money(item['platform_shipping'])}</td><td>{fmt_money(item['base_profit'])}</td><td>{fmt_money(item['extra_profit'])}</td><td>{fmt_money(item['total_profit'])}</td><td>{item['status']}</td></tr>"
        for item in return_items
    )
    return f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>lead</title>
  <script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js"></script>
  <style>
    :root {{
      color-scheme: dark;
      --bg: #04070d;
      --bg-soft: #0b101c;
      --card: rgba(18,24,38,.72);
      --card-strong: rgba(28,35,54,.82);
      --border: rgba(255,255,255,.12);
      --text: #f6fbff;
      --muted: #9aa8ba;
      --burgundy: #8b1744;
      --burgundy-2: #e43f80;
      --purple: #7c4dff;
      --cyan: #20e7d6;
      --amber: #f3b354;
      --shadow: 0 24px 80px rgba(0,0,0,.52);
    }}
    * {{ box-sizing: border-box; }}
    html {{ scroll-behavior:smooth; }}
    body {{
      margin:0;
      min-height:100vh;
      color:var(--text);
      font-family: Inter, "Avenir Next", "Segoe UI", "Noto Sans Arabic", system-ui, sans-serif;
      background:
        radial-gradient(circle at 16% 8%, rgba(228,63,128,.20), transparent 30%),
        radial-gradient(circle at 84% 0%, rgba(32,231,214,.12), transparent 26%),
        radial-gradient(circle at 50% 105%, rgba(124,77,255,.16), transparent 34%),
        linear-gradient(135deg, #02050a 0%, #08111d 48%, #0b0713 100%);
    }}
    a {{ color:inherit; text-decoration:none; }}
    .page {{ width:min(1180px, calc(100% - 32px)); margin:0 auto; padding:20px 0 34px; }}
    .hero {{
      margin:0 0 10px;
      padding:14px 26px 12px;
      border-radius:16px;
      background:
        linear-gradient(100deg, rgba(4,7,13,.96) 0%, rgba(21,16,38,.90) 44%, rgba(139,23,68,.68) 100%);
      border:1px solid rgba(255,255,255,.16);
      box-shadow:var(--shadow), inset 0 1px 0 rgba(255,255,255,.08);
      backdrop-filter: blur(22px) saturate(130%);
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
      border-radius:10px;
      border:1px solid rgba(255,255,255,.16);
      background:rgba(18,24,38,.70);
      color:rgba(255,255,255,.84);
      font-size:12px;
      backdrop-filter: blur(14px);
    }}
    .primary-button {{
      display:inline-flex;
      align-items:center;
      min-height:38px;
      padding:0 14px;
      border-radius:10px;
      background:
        linear-gradient(135deg, rgba(139,23,68,.98), rgba(228,63,128,.82)),
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
    .metric-grid {{ display:grid; grid-template-columns:repeat(4, minmax(0,1fr)); gap:10px; margin-bottom:10px; }}
    .mini-band {{ display:grid; grid-template-columns:repeat(3, minmax(0,1fr)); gap:14px; margin:0 0 16px; }}
    .finance-band {{ grid-template-columns:repeat(4, minmax(0,1fr)); }}
    .metric-card, .panel, .mini-card {{
      position:relative;
      overflow:hidden;
      border-radius:10px;
      background:
        linear-gradient(180deg, rgba(17,24,39,.74), rgba(8,13,23,.68)),
        radial-gradient(circle at 12% 8%, rgba(255,255,255,.10), transparent 26%);
      border:1px solid rgba(255,255,255,.11);
      box-shadow:0 18px 44px rgba(0,0,0,.42), inset 0 1px 0 rgba(255,255,255,.06);
      backdrop-filter: blur(24px) saturate(135%);
      -webkit-backdrop-filter: blur(24px) saturate(135%);
    }}
    .metric-card::before, .panel::before, .mini-card::before {{
      content:'';
      position:absolute;
      inset:0;
      background:
        linear-gradient(90deg, rgba(228,63,128,.62), rgba(124,77,255,.50), rgba(32,231,214,.58)) top / 100% 2px no-repeat,
        linear-gradient(135deg, rgba(255,255,255,.10) 0%, rgba(255,255,255,.035) 22%, rgba(255,255,255,0) 48%);
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
      padding:16px 18px;
      min-height:102px;
      display:flex;
      flex-direction:column;
      justify-content:space-between;
    }}
    .metric-label {{ display:block; color:var(--muted); font-size:13px; line-height:1.45; }}
    .metric-card strong {{ display:block; margin-top:8px; font-size:22px; line-height:1; font-weight:850; letter-spacing:-0.02em; font-variant-numeric: tabular-nums; color:#fff; text-shadow:0 2px 18px rgba(32,231,214,.14); }}
    .metric-footer {{ margin-top:8px; color:rgba(255,255,255,.58); font-size:12px; line-height:1.35; }}
    .mini-card .metric-label {{ font-size:13px; line-height:1.35; }}
    .mini-card strong {{ display:block; margin-top:6px; font-size:20px; line-height:1.15; font-variant-numeric: tabular-nums; color:#fff; text-shadow:0 2px 16px rgba(255,255,255,.14); }}
    .mini-card .metric-footer {{ font-size:12px; line-height:1.35; }}
    .profit-details-card {{ margin:0 0 14px; }}
    .details-panel {{ padding:16px; }}
    .compact-header {{ margin-bottom:12px; }}
    .details-grid {{ display:grid; grid-template-columns:repeat(5, minmax(0,1fr)); gap:10px; }}
    .detail-chip {{
      min-height:74px;
      padding:12px;
      border-radius:10px;
      background:rgba(255,255,255,.045);
      border:1px solid rgba(255,255,255,.09);
      display:flex;
      flex-direction:column;
      justify-content:center;
      gap:6px;
    }}
    .detail-chip span {{ color:var(--muted); font-size:12px; }}
    .detail-chip strong {{ color:#fff; font-size:18px; line-height:1.1; font-variant-numeric:tabular-nums; }}
    .content-grid {{ display:grid; grid-template-columns:1.3fr .92fr; gap:14px; margin-bottom:14px; }}
    .panel {{ padding:18px; }}
    .panel-header {{ display:flex; align-items:flex-start; justify-content:space-between; gap:14px; margin-bottom:16px; }}
    .eyebrow {{ margin:0; color:rgba(255,255,255,.88); font-size:13px; font-weight:750; letter-spacing:.02em; }}
    .subtle {{ margin:6px 0 0; color:var(--muted); font-size:13px; }}
    .chart-shell {{ height:300px; }}
    .combo-chart {{ width:100%; height:100%; display:block; }}
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
      border-radius:10px;
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
      .metric-grid, .content-grid, .table-grid, .mini-band, .finance-band, .statement-summary, .details-grid {{ grid-template-columns:1fr; }}
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
      {metric_card("إجمالي الإيرادات", fmt_money(totals["revenue"]), "الشحنات + إيرادات المرتجعات")}
      {metric_card("إجمالي الربح", fmt_money(totals["total"]), f"الشحنات المحسوبة: {period_label}")}
      {metric_card("عدد الشحنات", totals["active"], f"غير داخلة في الربح: {totals['excluded']}")}
      {metric_card("مبلغ COD", fmt_money(totals["cod_amount"]), "إجمالي مبالغ COD")}
    </section>

    {profit_details}

    <section class="content-grid">
      <article class="panel">
        <div class="panel-header">
          <div>
            <p class="eyebrow">الأداء اليومي</p>
            <p class="subtle">آخر 7 أيام</p>
          </div>
        </div>
        <div class="chart-shell">
          <canvas id="dailyComboChart" class="combo-chart" aria-label="الأداء اليومي"></canvas>
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
  <script>
    const dailyLabels = {json.dumps(daily_chart_labels, ensure_ascii=False)};
    const dailyProfit = {json.dumps(daily_chart_profit, ensure_ascii=False)};
    const dailyCount = {json.dumps(daily_chart_count, ensure_ascii=False)};
    const chartCanvas = document.getElementById('dailyComboChart');
    function drawCanvasComboChart(canvas, labels, profit, counts) {{
      const ctx = canvas.getContext('2d');
      const dpr = window.devicePixelRatio || 1;
      const rect = canvas.getBoundingClientRect();
      canvas.width = Math.max(1, Math.floor(rect.width * dpr));
      canvas.height = Math.max(1, Math.floor(rect.height * dpr));
      ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
      const width = rect.width;
      const height = rect.height;
      const pad = {{ top: 34, right: 36, bottom: 34, left: 44 }};
      const plotW = width - pad.left - pad.right;
      const plotH = height - pad.top - pad.bottom;
      const maxProfit = Math.max(...profit, 1);
      const maxCount = Math.max(...counts, 1);

      ctx.clearRect(0, 0, width, height);
      ctx.font = '10px Inter, "Noto Sans Arabic", system-ui';
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';

      for (let i = 0; i <= 4; i += 1) {{
        const y = pad.top + (plotH * i / 4);
        ctx.strokeStyle = 'rgba(255,255,255,.06)';
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(pad.left, y);
        ctx.lineTo(width - pad.right, y);
        ctx.stroke();
      }}

      const step = plotW / Math.max(labels.length, 1);
      const barW = Math.min(30, step * 0.46);
      const gradient = ctx.createLinearGradient(0, pad.top, 0, height - pad.bottom);
      gradient.addColorStop(0, 'rgba(228,63,128,.92)');
      gradient.addColorStop(1, 'rgba(228,63,128,.24)');

      labels.forEach((label, index) => {{
        const centerX = pad.left + step * index + step / 2;
        const barH = (profit[index] / maxProfit) * plotH;
        const x = centerX - barW / 2;
        const y = height - pad.bottom - barH;
        ctx.fillStyle = gradient;
        ctx.beginPath();
        ctx.roundRect(x, y, barW, barH, 6);
        ctx.fill();
        ctx.fillStyle = 'rgba(246,251,255,.58)';
        ctx.fillText(label, centerX, height - 16);
      }});

      const linePoints = counts.map((count, index) => {{
        const centerX = pad.left + step * index + step / 2;
        const y = height - pad.bottom - ((count / maxCount) * plotH);
        return [centerX, y];
      }});
      ctx.strokeStyle = 'rgba(32,231,214,1)';
      ctx.lineWidth = 2.5;
      ctx.beginPath();
      linePoints.forEach(([x, y], index) => {{
        if (index === 0) ctx.moveTo(x, y);
        else ctx.lineTo(x, y);
      }});
      ctx.stroke();
      linePoints.forEach(([x, y]) => {{
        ctx.fillStyle = 'rgba(32,231,214,1)';
        ctx.beginPath();
        ctx.arc(x, y, 4, 0, Math.PI * 2);
        ctx.fill();
        ctx.strokeStyle = '#07111d';
        ctx.lineWidth = 2;
        ctx.stroke();
      }});

      ctx.textAlign = 'right';
      ctx.fillStyle = 'rgba(246,251,255,.68)';
      ctx.fillText('الربح اليومي', width - pad.right, 16);
      ctx.fillStyle = 'rgba(32,231,214,.78)';
      ctx.fillText('عدد الشحنات', width - pad.right - 88, 16);
    }}

    if (chartCanvas && window.Chart) {{
      const chartGradient = chartCanvas.getContext('2d').createLinearGradient(0, 0, 0, 280);
      chartGradient.addColorStop(0, 'rgba(228,63,128,0.92)');
      chartGradient.addColorStop(1, 'rgba(228,63,128,0.22)');
      new Chart(chartCanvas, {{
        data: {{
          labels: dailyLabels,
          datasets: [
            {{
              type: 'bar',
              label: 'الربح اليومي',
              data: dailyProfit,
              yAxisID: 'profit',
              backgroundColor: chartGradient,
              borderColor: 'rgba(228,63,128,1)',
              borderWidth: 1,
              borderRadius: 6,
              maxBarThickness: 30
            }},
            {{
              type: 'line',
              label: 'عدد الشحنات',
              data: dailyCount,
              yAxisID: 'count',
              borderColor: 'rgba(32,231,214,1)',
              backgroundColor: 'rgba(32,231,214,0.16)',
              borderWidth: 2.5,
              tension: 0.38,
              pointRadius: 3.5,
              pointHoverRadius: 5,
              pointBackgroundColor: 'rgba(32,231,214,1)',
              pointBorderColor: '#07111d',
              pointBorderWidth: 2
            }}
          ]
        }},
        options: {{
          responsive: true,
          maintainAspectRatio: false,
          interaction: {{ mode: 'index', intersect: false }},
          plugins: {{
            legend: {{
              position: 'top',
              align: 'end',
              labels: {{
                color: 'rgba(246,251,255,.76)',
                boxWidth: 10,
                boxHeight: 10,
                usePointStyle: true,
                font: {{ family: 'Inter, "Noto Sans Arabic", system-ui', size: 11 }}
              }}
            }},
            tooltip: {{
              rtl: true,
              textDirection: 'rtl',
              backgroundColor: 'rgba(5,8,14,.94)',
              borderColor: 'rgba(255,255,255,.14)',
              borderWidth: 1,
              titleColor: '#fff',
              bodyColor: 'rgba(246,251,255,.82)'
            }}
          }},
          scales: {{
            x: {{
              grid: {{ color: 'rgba(255,255,255,.05)' }},
              ticks: {{ color: 'rgba(246,251,255,.58)', font: {{ size: 10 }} }}
            }},
            profit: {{
              position: 'left',
              grid: {{ color: 'rgba(255,255,255,.06)' }},
              ticks: {{ color: 'rgba(246,251,255,.58)', font: {{ size: 10 }} }}
            }},
            count: {{
              position: 'right',
              grid: {{ drawOnChartArea: false }},
              ticks: {{ color: 'rgba(32,231,214,.72)', precision: 0, font: {{ size: 10 }} }}
            }}
          }}
        }}
      }});
    }} else if (chartCanvas) {{
      drawCanvasComboChart(chartCanvas, dailyLabels, dailyProfit, dailyCount);
    }}
  </script>
</body>
</html>"""


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    totals, top_merchants, top_cities, top_carriers, statuses, daily, daily_count, finance, cod_items, return_items, statement, period_label = load_data()
    data = {
        "totals": totals,
        "top_merchants": top_merchants,
        "top_cities": top_cities,
        "daily": daily,
        "daily_count": daily_count,
        "finance": finance,
        "cod_items": cod_items,
        "return_items": return_items,
        "top_carriers": top_carriers,
        "statement": statement,
        "period_label": period_label,
    }
    with open(INDEX, "w", encoding="utf-8") as f:
        f.write(build_html(data))
    print(INDEX)


if __name__ == "__main__":
    main()
