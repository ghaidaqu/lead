"""Synthetic financial analysis used exclusively by the GF Demo presentation."""

from __future__ import annotations

from collections import defaultdict


# This is deliberately synthetic presentation data.  It has no connection to
# production systems, PostgreSQL, customer systems, or uploaded client files.
SYNTHETIC_FINANCE = {
    "revenue": 9_993.75,
    "costs": {
        "مواد غذائية": 4_050.00,
        "رواتب": 1_650.00,
        "إيجارات": 1_100.00,
        "توصيل": 830.00,
        "هدر": 217.00,
    },
    "branches": (
        {"name": "فرع ١", "revenue": 2_720.00, "operating_cost": 4_980.00, "waste": 220.00, "profit": 500.00},
        {"name": "فرع ٢", "revenue": 2_430.00, "operating_cost": 5_360.00, "waste": 410.00, "profit": 186.75},
        {"name": "فرع ٣", "revenue": 2_480.00, "operating_cost": 4_720.00, "waste": 180.00, "profit": 512.50},
        {"name": "فرع ٤", "revenue": 2_363.75, "operating_cost": 5_110.00, "waste": 260.00, "profit": 447.50},
    ),
    "carriers": (
        {"name": "جاهز", "orders": 116, "fees": 520.25},
        {"name": "هنقرستيشن", "orders": 142, "fees": 625.00},
        {"name": "كيتا", "orders": 109, "fees": 501.50},
        {"name": "نينجا", "orders": 108, "fees": 500.00},
    ),
    "products": (
        {"name": "شاورما دجاج", "orders": 142, "profit": 184.50},
        {"name": "برجر لحم", "orders": 129, "profit": 163.25},
        {"name": "سلطة جانبية", "orders": 256, "profit": 18.50},
        {"name": "بطاطس مقلية", "orders": 320, "profit": 22.75},
        {"name": "مشروب غازي", "orders": 298, "profit": 24.00},
        {"name": "إضافات صوص", "orders": 255, "profit": 19.50},
    ),
}


def pct(numerator: float, denominator: float) -> float:
    return round((numerator / denominator * 100) if denominator else 0, 1)


def demo_analysis() -> dict:
    """Calculate decision-ready insights from the local synthetic data."""
    data = SYNTHETIC_FINANCE
    revenue = data["revenue"]
    costs = data["costs"]
    net_profit = round(revenue - sum(costs.values()), 2)
    food_cost_pct = pct(costs["مواد غذائية"], revenue)
    waste_pct = pct(costs["هدر"], revenue)
    margin_pct = pct(net_profit, revenue)

    branch_rows = []
    for branch in data["branches"]:
        row = dict(branch)
        row["waste_pct"] = pct(row["waste"], row["operating_cost"])
        row["margin_pct"] = pct(row["profit"], row["revenue"])
        branch_rows.append(row)
    best_branch = max(branch_rows, key=lambda row: row["profit"])
    weakest_branch = min(branch_rows, key=lambda row: row["profit"])
    highest_waste_branch = max(branch_rows, key=lambda row: row["waste_pct"])

    carrier_rows = []
    total_orders = sum(row["orders"] for row in data["carriers"])
    for carrier in data["carriers"]:
        row = dict(carrier)
        row["order_share_pct"] = pct(row["orders"], total_orders)
        row["fee_per_order"] = round(row["fees"] / row["orders"], 2)
        carrier_rows.append(row)
    top_carrier = max(carrier_rows, key=lambda row: row["orders"])
    highest_fee_carrier = max(carrier_rows, key=lambda row: row["fee_per_order"])

    product_rows = []
    for product in data["products"]:
        row = dict(product)
        row["profit_per_order"] = round(row["profit"] / row["orders"], 2)
        product_rows.append(row)
    review_products = sorted(product_rows, key=lambda row: (-row["orders"], row["profit_per_order"]))[:3]

    alert_count = sum((food_cost_pct > 38, waste_pct > 2, margin_pct < 24, highest_waste_branch["waste_pct"] > 5))
    health_score = max(0, min(100, round(100 - (food_cost_pct - 35) * 2 - max(0, waste_pct - 2) * 5 - max(0, 24 - margin_pct) * 3)))

    return {
        "synthetic": True,
        "scope": {"branches": len(branch_rows), "carriers": len(carrier_rows), "cost_categories": len(costs)},
        "kpis": {
            "revenue": revenue,
            "net_profit": net_profit,
            "margin_pct": margin_pct,
            "food_cost_pct": food_cost_pct,
            "waste_pct": waste_pct,
            "health_score": health_score,
            "alerts": alert_count,
        },
        "diagnosis": [
            {
                "title": "الضغط الرئيسي على الربح",
                "value": f"تكلفة الطعام {food_cost_pct}%",
                "detail": f"أعلى من حد المتابعة التجريبي 38% بمقدار {round(food_cost_pct - 38, 1)} نقطة.",
                "level": "warning",
            },
            {
                "title": "الهدر الذي يحتاج إجراء",
                "value": f"{highest_waste_branch['name']} — {highest_waste_branch['waste_pct']}%",
                "detail": "الأعلى بين الفروع؛ راجع الاستلام والتحضير في هذا الفرع أولًا.",
                "level": "warning",
            },
            {
                "title": "قناة الطلب الأكثر أثرًا",
                "value": f"{top_carrier['name']} — {top_carrier['order_share_pct']}%",
                "detail": f"الأكثر طلبًا، بينما أعلى رسم لكل طلب لدى {highest_fee_carrier['name']} ({highest_fee_carrier['fee_per_order']:.2f} ريال).",
                "level": "neutral",
            },
        ],
        "actions": [
            {"priority": "أولوية ١", "action": f"خفض هدر {highest_waste_branch['name']} إلى 5%", "impact": "يحسن الهامش ويقلل الهدر التشغيلي", "owner": "مدير الفرع"},
            {"priority": "أولوية ٢", "action": "مراجعة تكلفة وصفات الأصناف عالية البيع", "impact": "حماية الربح دون المساس بحجم الطلبات", "owner": "التشغيل"},
            {"priority": "أولوية ٣", "action": f"مراجعة رسوم {highest_fee_carrier['name']}", "impact": "تحسين صافي ربح قنوات التوصيل", "owner": "الشراكات"},
        ],
        "branches": branch_rows,
        "carriers": carrier_rows,
        "review_products": review_products,
    }
