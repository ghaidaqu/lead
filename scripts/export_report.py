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
    ("included_in_profit", "محتسب"),
]

_HEADER_FILL = PatternFill("solid", fgColor="7A2438")
_HEADER_FONT = Font(color="FFFFFF", bold=True)


def _style_header(ws) -> None:
    for cell in ws[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "A2"


def build_report_xlsx(conn) -> bytes:
    """Build the report workbook from the DB and return it as bytes."""
    with conn.cursor() as cur:
        cur.execute(
            f"SELECT {', '.join(c for c, _ in SHIPMENT_COLUMNS)} "
            "FROM shipments ORDER BY shipment_date NULLS LAST, order_id"
        )
        rows = [dict(r) for r in cur.fetchall()]

    wb = Workbook()
    ws = wb.active
    ws.title = "الشحنات"
    ws.sheet_view.rightToLeft = True
    ws.append([label for _, label in SHIPMENT_COLUMNS])
    total = 0.0
    for row in rows:
        ws.append([_fmt(row.get(col)) for col, _ in SHIPMENT_COLUMNS])
        total += float(row.get("total_profit") or 0)
    _style_header(ws)

    summary = wb.create_sheet("الملخص")
    summary.sheet_view.rightToLeft = True
    summary.append(["البند", "القيمة"])
    summary.append(["عدد الشحنات", len(rows)])
    summary.append(["إجمالي الربح", round(total, 2)])
    summary.append(["المحتسبة في الربح", sum(1 for r in rows if r.get("included_in_profit"))])
    _style_header(summary)

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def _fmt(value: Any):
    if value is None:
        return ""
    if hasattr(value, "isoformat"):
        return value.isoformat()[:10]
    if isinstance(value, bool):
        return "نعم" if value else "لا"
    try:
        return float(value)
    except (TypeError, ValueError):
        return str(value)
