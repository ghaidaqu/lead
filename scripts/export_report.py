"""On-demand Excel export built in-memory from PostgreSQL.

No .xlsx file is ever written to disk — the workbook is streamed straight to the
client. `openpyxl` is used here and only here.
"""

from __future__ import annotations

import io
from typing import Any

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

SHIPMENT_COLUMNS = [
    ("order_id", "رقم الطلب"),
    ("merchant_name", "التاجر"),
    ("store_name", "المتجر"),
    ("customer_name", "العميل"),
    ("city", "المدينة"),
    ("carrier", "شركة الشحن"),
    ("payment_type", "الدفع"),
    ("status", "الحالة"),
    ("shipment_date", "تاريخ الشحن"),
    ("weight", "الوزن"),
    ("cod_amount", "مبلغ COD"),
    ("customer_price_net", "صافي العميل"),
    ("platform_cost_net", "صافي المنصة"),
    ("base_profit", "ربح الشحنة"),
    ("extra_profit", "ربح الوزن الزائد"),
    ("cod_profit", "ربح COD"),
    ("total_profit", "إجمالي الربح"),
    ("actual_revenue", "الإيراد الفعلي"),
    ("actual_base_cost", "التكلفة الأساسية الفعلية"),
    ("actual_extra_cost", "رسوم وزن/COD محمَّلة"),
    ("actual_profit", "صافي الربح الفعلي"),
    ("included_in_profit", "محتسب"),
]

WALLET_COLUMNS = [
    ("transaction_key", "Transaction Key"),
    ("transaction_date", "Transaction Date"),
    ("user_name", "User"),
    ("description", "Description"),
    ("amount", "Amount"),
    ("transaction_type", "Type"),
    ("balance_before", "Balance Before"),
    ("balance_after", "Balance After"),
    ("source_page", "Source Page"),
    ("raw_payload", "Raw Payload"),
]

PAYMENT_COLUMNS = [
    ("payment_key", "Payment Key"),
    ("payment_date", "Payment Date"),
    ("customer_name", "Customer"),
    ("amount", "Amount"),
    ("method", "Method"),
    ("status", "Status"),
    ("source_page", "Source Page"),
    ("raw_payload", "Raw Payload"),
]

_HEADER_FILL = PatternFill("solid", fgColor="7A2438")
_HEADER_FONT = Font(color="FFFFFF", bold=True)


def _style_header(ws) -> None:
    for cell in ws[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def build_report_xlsx(conn, date_from=None, date_to=None) -> bytes:
    """Build the report workbook from the DB and return it as bytes."""
    shipment_where, shipment_params = _date_where("shipment_date", date_from, date_to)
    wallet_where, wallet_params = _date_where("transaction_date", date_from, date_to)
    payment_where, payment_params = _date_where("payment_date", date_from, date_to)
    with conn.cursor() as cur:
        cur.execute(
            f"SELECT {', '.join(c for c, _ in SHIPMENT_COLUMNS)} "
            f"FROM shipments {shipment_where} ORDER BY shipment_date NULLS LAST, order_id",
            shipment_params,
        )
        rows = [dict(r) for r in cur.fetchall()]
        cur.execute(
            f"SELECT {', '.join(c for c, _ in WALLET_COLUMNS)} "
            f"FROM wallet_transactions {wallet_where} ORDER BY transaction_date NULLS LAST, transaction_key",
            wallet_params,
        )
        wallet_rows = [dict(r) for r in cur.fetchall()]
        cur.execute(
            f"SELECT {', '.join(c for c, _ in PAYMENT_COLUMNS)} "
            f"FROM payments {payment_where} ORDER BY payment_date NULLS LAST, payment_key",
            payment_params,
        )
        payment_rows = [dict(r) for r in cur.fetchall()]

    wb = Workbook()
    ws = wb.active
    ws.title = "الشحنات"
    ws.sheet_view.rightToLeft = True
    ws.append([label for _, label in SHIPMENT_COLUMNS])
    included_count = 0
    excluded_count = 0
    cod_amount_total = 0.0
    revenue_total = 0.0
    platform_cost_total = 0.0
    base_profit_total = 0.0
    extra_profit_total = 0.0
    cod_profit_total = 0.0
    total_profit = 0.0
    actual_revenue_total = 0.0
    actual_base_total = 0.0
    absorbed_fees_total = 0.0
    actual_profit_total = 0.0
    has_actual_values = False
    for row in rows:
        ws.append([_fmt(row.get(col)) for col, _ in SHIPMENT_COLUMNS])
        if row.get("included_in_profit"):
            included_count += 1
            cod_amount_total += float(row.get("cod_amount") or 0)
            revenue_total += float(row.get("customer_price_net") or 0)
            platform_cost_total += float(row.get("platform_cost_net") or 0)
            base_profit_total += float(row.get("base_profit") or 0)
            extra_profit_total += float(row.get("extra_profit") or 0)
            cod_profit_total += float(row.get("cod_profit") or 0)
            total_profit += float(row.get("total_profit") or 0)
            if any(row.get(col) is not None for col in ("actual_revenue", "actual_base_cost", "actual_extra_cost", "actual_profit")):
                has_actual_values = True
            actual_revenue_total += float(row.get("actual_revenue") or 0)
            actual_base_total += float(row.get("actual_base_cost") or 0)
            absorbed_fees_total += float(row.get("actual_extra_cost") or 0)
            actual_profit_total += float(row.get("actual_profit") or 0)
        else:
            excluded_count += 1
    _style_header(ws)

    summary = wb.create_sheet("الملخص")
    summary.sheet_view.rightToLeft = True
    summary.append(["البند", "القيمة"])
    summary.append(["الفترة من", date_from.isoformat() if date_from else ""])
    summary.append(["الفترة إلى", date_to.isoformat() if date_to else ""])
    summary.append(["عدد الشحنات", included_count])
    summary.append(["غير داخلة في الربح", excluded_count])
    summary.append(["مبلغ COD", round(cod_amount_total, 2)])
    summary.append(["إجمالي الإيرادات", round(actual_revenue_total if has_actual_values else revenue_total, 2)])
    summary.append(["التكلفة الأساسية", round(actual_base_total if has_actual_values else platform_cost_total, 2)])
    summary.append(["رسوم وزن/COD محمَّلة", round(absorbed_fees_total, 2)])
    summary.append(["صافي الربح الفعلي", round(actual_profit_total if has_actual_values else total_profit, 2)])
    summary.append(["ربح الشحنة", round(base_profit_total, 2)])
    summary.append(["ربح الوزن الزائد", round(extra_profit_total, 2)])
    summary.append(["ربح COD", round(cod_profit_total, 2)])
    summary.append(["إجمالي الربح", round(total_profit, 2)])
    summary.append(["المحتسبة في الربح", sum(1 for r in rows if r.get("included_in_profit"))])
    _style_header(summary)

    _append_raw_sheet(wb, "Raw_Wallet", WALLET_COLUMNS, wallet_rows)
    _append_raw_sheet(wb, "Raw_Payments", PAYMENT_COLUMNS, payment_rows)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _date_where(column: str, date_from, date_to) -> tuple[str, list[Any]]:
    clauses = []
    params: list[Any] = []
    if date_from:
        clauses.append(f"{column} >= %s")
        params.append(date_from)
    if date_to:
        clauses.append(f"{column} < %s")
        import datetime as _dt
        params.append(date_to + _dt.timedelta(days=1))
    return (("WHERE " + " AND ".join(clauses)) if clauses else "", params)


def _append_raw_sheet(wb: Workbook, title: str, columns: list[tuple[str, str]], rows: list[dict[str, Any]]) -> None:
    ws = wb.create_sheet(title)
    ws.append([label for _, label in columns])
    for row in rows:
        ws.append([_fmt(row.get(col)) for col, _ in columns])
    _style_header(ws)


def _fmt(value: Any):
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()[:10]
    if isinstance(value, bool):
        return "نعم" if value else "لا"
    if isinstance(value, (list, dict)):
        import json
        return json.dumps(value, ensure_ascii=False, default=str)
    try:
        return float(value)
    except (TypeError, ValueError):
        return str(value)
