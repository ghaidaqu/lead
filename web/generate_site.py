import re
import sys
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from scripts.db_store import fetch_dashboard_rows, get_conn


OUT_DIR = Path(__file__).resolve().parent
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


ARABIC_MONTHS = ["", "يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو",
                 "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]


def period_label_from_dates(dates):
    """Derive the period label from the actual shipment dates — real month
    name(s) and the true min–max day range (handles cross-month spans).
    Nothing about the month or range is hardcoded."""
    parsed = []
    for value in dates:
        if not value:
            continue
        match = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(value))
        if not match:
            continue
        try:
            parsed.append(date(int(match.group(1)), int(match.group(2)), int(match.group(3))))
        except ValueError:
            continue
    if not parsed:
        return ""
    lo, hi = min(parsed), max(parsed)
    if (lo.year, lo.month) == (hi.year, hi.month):
        month = ARABIC_MONTHS[lo.month]
        return month if lo.day == hi.day else f"{lo.day}-{hi.day} {month}"
    return f"{lo.day} {ARABIC_MONTHS[lo.month]} - {hi.day} {ARABIC_MONTHS[hi.month]}"


def _row_value(row, index, default=None):
    try:
        value = row[index]
    except Exception:
        return default
    return default if value is None else value


def _row_date(value):
    if isinstance(value, datetime):
        return value.date().isoformat()
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return clean(value)


def _to_date(value):
    match = re.search(r"(\d{4})[-/](\d{1,2})[-/](\d{1,2})", str(value or ""))
    if not match:
        return None
    try:
        return date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    except ValueError:
        return None


def _in_range(value, lo, hi):
    """True if value's date falls within [lo, hi]. When no range is set, all
    rows pass; when a range IS set, undated rows are excluded."""
    if lo is None and hi is None:
        return True
    day = _to_date(value)
    if day is None:
        return False
    if lo and day < lo:
        return False
    if hi and day > hi:
        return False
    return True


def load_data_from_db(date_from=None, date_to=None):
    with get_conn() as conn:
        payload = fetch_dashboard_rows(conn)

    shipment_rows = payload["shipments"]
    wallet_rows = payload["wallet"]
    cod_rows_db = payload["cod"]

    # Real collection/transfer dates come from the cod_collections table
    # (scraped from collect-cod.php), keyed by order_id for lookup.
    cod_date_map = {}
    for cr in cod_rows_db:
        oid = str(cr.get("order_id") or "").strip()
        if oid:
            cod_date_map[oid] = {
                "collection_date": cr.get("collection_date"),
                "transfer_date": cr.get("transfer_date"),
            }

    items = []
    cod_items = []
    by_merchant = defaultdict(lambda: {"count": 0, "total": 0.0})
    by_city = defaultdict(lambda: {"count": 0, "total": 0.0})
    by_carrier = defaultdict(lambda: {"count": 0, "total": 0.0})
    by_status = defaultdict(int)
    by_date = defaultdict(float)
    by_date_revenue = defaultdict(float)
    by_date_count = defaultdict(int)
    all_shipment_dates = set()
    all_items_by_order = {}
    in_range_count = 0  # shipments within the selected date range (included or not)
    finance = {
        "bank": {"count": 0, "total": 0.0},
        "moyasar": {"count": 0, "total": 0.0},
        "shipping_return": {"count": 0, "total": 0.0},
        "tax_deduction": {"count": 0, "total": 0.0},
        "shipping_refund": {"count": 0, "total": 0.0},
        "other_net": 0.0,
        "other": {"count": 0, "total": 0.0},
    }
    statement = {"summary": {}, "rows": []}
    return_items = []
    total_customer_shipping = 0.0

    for row in shipment_rows:
        raw = row.get("raw_payload") or []
        if not raw:
            raw = [
                row.get("order_id"),
                row.get("tracking_number"),
                row.get("merchant_name"),
                row.get("store_name"),
                row.get("customer_name"),
                row.get("city"),
                row.get("shipping_charge"),
                row.get("cod_amount"),
                None,
                row.get("payment_type"),
                row.get("carrier"),
                row.get("weight"),
                row.get("status"),
                row.get("shipment_date"),
                "نعم" if row.get("included_in_profit") else "لا",
                row.get("customer_price_gross"),
                row.get("customer_price_net"),
                row.get("platform_cost_net"),
                row.get("extra_kg"),
                row.get("base_profit"),
                row.get("extra_profit"),
                row.get("cod_profit"),
                row.get("total_profit"),
                row.get("source_row"),
            ]
        if not any(v is not None for v in raw):
            continue
        if raw[0] in (None, ""):
            continue
        customer_net = money(row.get("customer_price_net") if row.get("customer_price_net") is not None else _row_value(raw, 16))
        platform_shipping = money(row.get("platform_cost_net") if row.get("platform_cost_net") is not None else _row_value(raw, 17))
        shipping_profit = money(row.get("base_profit") if row.get("base_profit") is not None else _row_value(raw, 19))
        extra_profit = money(row.get("extra_profit") if row.get("extra_profit") is not None else _row_value(raw, 20))
        fee_profit = money(row.get("cod_profit") if row.get("cod_profit") is not None else _row_value(raw, 21))
        total_profit = money(row.get("total_profit") if row.get("total_profit") is not None else _row_value(raw, 22))
        actual_profit = money(row.get("actual_profit")) if row.get("actual_profit") is not None else 0.0

        item = {
            "order_id": raw[0],
            "tracking": clean(row.get("tracking_number") or _row_value(raw, 1)),
            "merchant": clean(row.get("merchant_name") or _row_value(raw, 2)),
            "customer": clean(row.get("customer_name") or _row_value(raw, 4)),
            "city": clean(row.get("city") or _row_value(raw, 5)),
            "carrier": clean(row.get("carrier") or _row_value(raw, 10)),
            "date": _row_date(row.get("shipment_date") or _row_value(raw, 13)),
            "status": clean(row.get("status") or _row_value(raw, 12)),
            "weight": money(row.get("weight") if row.get("weight") is not None else _row_value(raw, 11)),
            "included": bool(row.get("included_in_profit")),
            "customer_shipping": money(row.get("shipping_charge") if row.get("shipping_charge") is not None else _row_value(raw, 6)),
            "cod_amount": money(row.get("cod_amount") if row.get("cod_amount") is not None else _row_value(raw, 7)),
            "customer_gross": money(row.get("customer_price_gross") if row.get("customer_price_gross") is not None else _row_value(raw, 15)),
            "customer_net": customer_net,
            "platform_shipping": platform_shipping,
            "shipping_profit": shipping_profit,
            "extra_profit": extra_profit,
            "fee_profit": fee_profit,
            "total_profit": total_profit,
            "actual_profit": actual_profit,
            "review_diff": shipping_profit,
            "collection_date": cod_date_map.get(str(raw[0]), {}).get("collection_date"),
            "transfer_date": cod_date_map.get(str(raw[0]), {}).get("transfer_date"),
        }
        if not _in_range(item["date"], date_from, date_to):
            continue
        in_range_count += 1
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
        by_merchant[item["merchant"]]["total"] += item["actual_profit"]
        by_city[item["city"]]["count"] += 1
        by_city[item["city"]]["total"] += item["actual_profit"]
        by_carrier[item["carrier"]]["count"] += 1
        by_carrier[item["carrier"]]["total"] += item["actual_profit"]
        by_status[item["status"]] += 1
        if item["date"]:
            by_date[item["date"]] += item["actual_profit"]
            by_date_revenue[item["date"]] += item["customer_shipping"]
            by_date_count[item["date"]] += 1

    for row in wallet_rows:
        raw = row.get("raw_payload") or []
        if not _in_range(_row_value(raw, 1), date_from, date_to):
            continue
        typ = clean(_row_value(raw, 3))
        note = clean(_row_value(raw, 5))
        amount = money(_row_value(raw, 4))
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
                    net_ratio = (linked_item["customer_net"] / linked_item["customer_gross"]) if linked_item["customer_gross"] else 0.85
                    revenue_net = revenue * net_ratio
                    platform_cost_net = linked_item["platform_shipping"]
                    base_profit = revenue_net - platform_cost_net
                    extra_profit = linked_item["extra_profit"]
                    total_profit = base_profit + extra_profit
                else:
                    revenue_net = revenue * 0.85
                    platform_cost_net = 0.0
                    base_profit = revenue_net
                    extra_profit = 0.0
                    total_profit = revenue_net
                return_items.append({
                    "order_id": order_id or "-",
                    "merchant": linked_item["merchant"] if linked_item else "مقدر",
                    "carrier": linked_item["carrier"] if linked_item else "مقدر",
                    "weight": linked_item["weight"] if linked_item else 0.0,
                    "revenue": revenue,
                    "revenue_net": revenue_net,
                    "platform_shipping": platform_cost_net,
                    "base_profit": base_profit,
                    "extra_profit": extra_profit,
                    "total_profit": total_profit,
                    "matched": linked_item is not None,
                    "status": "محسوب" if linked_item else "مقدر",
                })
            elif typ == "admin_deduction" and "خصم ضريبة القيمة المضافة" in note:
                finance["other_net"] += amount
                finance["tax_deduction"]["count"] += 1
                finance["tax_deduction"]["total"] += abs(amount)
            elif typ == "استرداد" and "استرداد تكلفة شحن" in note:
                finance["other_net"] += amount
                finance["shipping_refund"]["count"] += 1
                finance["shipping_refund"]["total"] += amount
            elif typ == "إيداع":
                # a deposit not via bank/Moyasar (handled above) shouldn't reach here,
                # but keep the bucket consistent if the site adds new deposit notes.
                finance["other"]["count"] += 1
                finance["other"]["total"] += amount
            else:
                # شحن (shipping charges) and anything else is NOT a deposit — track its
                # net only, never in the deposit buckets.
                finance["other_net"] += amount

    # Total deposits = every إيداع (bank + Moyasar + other), matching the site's
    # "إجمالي الإيداعات". Shipping charges are excluded.
    finance["total"] = finance["bank"]["total"] + finance["moyasar"]["total"] + finance["other"]["total"]
    top_merchants = sorted(by_merchant.items(), key=lambda kv: kv[1]["total"], reverse=True)[:5]
    top_cities = sorted(by_city.items(), key=lambda kv: kv[1]["total"], reverse=True)[:5]
    top_carriers = sorted(by_carrier.items(), key=lambda kv: kv[1]["total"], reverse=True)
    daily = sorted(by_date.items(), key=lambda kv: kv[0])
    daily_revenue = sorted(by_date_revenue.items(), key=lambda kv: kv[0])
    daily_count = sorted(by_date_count.items(), key=lambda kv: kv[0])
    period_label = period_label_from_dates(all_shipment_dates)
    return_revenue = sum(item["revenue"] for item in return_items)
    return_profit = sum(item.get("actual_profit", 0) for item in return_items)
    return_platform_cost = sum(item["platform_shipping"] for item in return_items)
    totals = {
        "records": in_range_count,
        "active": len(items),
        "shipping": len(items),
        "cod": sum(1 for item in items if item["fee_profit"] > 0),
        "cod_amount": sum(item["cod_amount"] for item in items if item["included"]),
        "revenue": total_customer_shipping + return_revenue,
        "cost": sum(item["platform_shipping"] for item in items) + return_platform_cost,
        "base": sum(item["shipping_profit"] for item in items),
        "extra": sum(item["extra_profit"] for item in items),
        "cod_profit": sum(item["fee_profit"] for item in items),
        "return_revenue": return_revenue,
        "return_profit": return_profit,
        "return_count": len(return_items),
        "total": sum(item["actual_profit"] for item in items) + return_profit,
        "excluded": in_range_count - len(items),
    }
    cod_items.sort(key=lambda x: (x["date"], x["order_id"]), reverse=True)
    return_items.sort(key=lambda x: (x["matched"], x["order_id"]), reverse=True)
    order_numbers = sorted(
        {
            int(item["order_id"])
            for item in all_items_by_order.values()
            if str(item["order_id"]).isdigit()
        }
    )
    missing_sequence_numbers = []
    if order_numbers:
        expected = set(range(order_numbers[0], order_numbers[-1] + 1))
        missing_sequence_numbers = sorted(expected.difference(order_numbers))
    return totals, top_merchants, top_cities, top_carriers, by_status, daily, daily_revenue, daily_count, finance, cod_items, return_items, statement, period_label, missing_sequence_numbers
