"""Synthetic financial analysis used exclusively by the GF Demo presentation."""

from __future__ import annotations

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

MONTHLY_COMPARISON = {
    "current": {
        "label": "يوليو 2026",
        "revenue": 9_993.75,
        "orders": 475,
        "expense_invoices": 18,
        "costs": {
            "مواد غذائية": 4_050.00,
            "رواتب": 1_650.00,
            "إيجارات": 1_100.00,
            "توصيل": 830.00,
            "هدر": 217.00,
        },
    },
    "previous": {
        "label": "يونيو 2026",
        "revenue": 9_025.40,
        "orders": 436,
        "expense_invoices": 16,
        "costs": {
            "مواد غذائية": 3_321.35,
            "رواتب": 1_590.00,
            "إيجارات": 1_075.00,
            "توصيل": 576.95,
            "هدر": 154.00,
        },
    },
    "trend": (
        {"month": "فبراير", "revenue": 8_240.00, "net_profit": 1_884.00, "margin_pct": 22.9},
        {"month": "مارس", "revenue": 8_610.00, "net_profit": 2_012.00, "margin_pct": 23.4},
        {"month": "أبريل", "revenue": 8_430.00, "net_profit": 1_926.00, "margin_pct": 22.8},
        {"month": "مايو", "revenue": 8_790.00, "net_profit": 2_118.00, "margin_pct": 24.1},
        {"month": "يونيو", "revenue": 9_025.40, "net_profit": 2_308.10, "margin_pct": 25.6},
        {"month": "يوليو", "revenue": 9_993.75, "net_profit": 2_146.75, "margin_pct": 21.5},
    ),
    "branches": (
        {"name": "فرع ١", "current_profit": 500.00, "previous_profit": 454.00, "current_margin": 18.4, "previous_margin": 18.1, "waste_pct": 4.4},
        {"name": "فرع ٢", "current_profit": 186.75, "previous_profit": 342.50, "current_margin": 7.7, "previous_margin": 14.8, "waste_pct": 7.6},
        {"name": "فرع ٣", "current_profit": 512.50, "previous_profit": 480.20, "current_margin": 20.7, "previous_margin": 20.0, "waste_pct": 3.8},
        {"name": "فرع ٤", "current_profit": 447.50, "previous_profit": 398.40, "current_margin": 18.9, "previous_margin": 17.5, "waste_pct": 5.1},
    ),
    "products": (
        {"name": "شاورما دجاج", "orders": 142, "revenue": 1_732.40, "profit": 184.50, "margin_pct": 31.8, "previous_profit": 156.80},
        {"name": "برجر لحم", "orders": 129, "revenue": 1_548.00, "profit": 163.25, "margin_pct": 29.6, "previous_profit": 151.10},
        {"name": "بطاطس مقلية", "orders": 320, "revenue": 1_120.00, "profit": 22.75, "margin_pct": 8.1, "previous_profit": 39.20},
        {"name": "مشروب غازي", "orders": 298, "revenue": 894.00, "profit": 24.00, "margin_pct": 9.4, "previous_profit": 34.60},
        {"name": "إضافات صوص", "orders": 255, "revenue": 510.00, "profit": 19.50, "margin_pct": 7.6, "previous_profit": 27.10},
        {"name": "سلطة جانبية", "orders": 256, "revenue": 768.00, "profit": 18.50, "margin_pct": 6.9, "previous_profit": 30.40},
    ),
    "carriers": (
        {"name": "جاهز", "orders": 116, "previous_orders": 104, "fees": 520.25, "previous_fees": 451.20},
        {"name": "هنقرستيشن", "orders": 142, "previous_orders": 126, "fees": 625.00, "previous_fees": 538.40},
        {"name": "كيتا", "orders": 109, "previous_orders": 101, "fees": 501.50, "previous_fees": 428.30},
        {"name": "نينجا", "orders": 108, "previous_orders": 105, "fees": 500.00, "previous_fees": 446.90},
    ),
}


def pct(numerator: float, denominator: float) -> float:
    return round((numerator / denominator * 100) if denominator else 0, 1)


def change_pct(current: float, previous: float) -> float:
    return round(((current - previous) / previous * 100) if previous else 0, 1)


def financial_comparison() -> dict:
    """Return a complete month-over-month analysis for the presentation page."""
    model = MONTHLY_COMPARISON
    current = model["current"]
    previous = model["previous"]
    current_costs = sum(current["costs"].values())
    previous_costs = sum(previous["costs"].values())
    current_profit = round(current["revenue"] - current_costs, 2)
    previous_profit = round(previous["revenue"] - previous_costs, 2)

    def metric(label: str, current_value: float, previous_value: float, unit: str, *, percentage_points: bool = False) -> dict:
        difference = round(current_value - previous_value, 2)
        return {
            "label": label,
            "current": current_value,
            "previous": previous_value,
            "difference": difference,
            "change_pct": round(difference, 1) if percentage_points else change_pct(current_value, previous_value),
            "unit": unit,
            "comparison_unit": "نقطة" if percentage_points else "%",
            "direction": "up" if difference > 0 else "down" if difference < 0 else "flat",
        }

    current_margin = pct(current_profit, current["revenue"])
    previous_margin = pct(previous_profit, previous["revenue"])
    current_food_cost = pct(current["costs"]["مواد غذائية"], current["revenue"])
    previous_food_cost = pct(previous["costs"]["مواد غذائية"], previous["revenue"])
    current_waste = pct(current["costs"]["هدر"], current["revenue"])
    previous_waste = pct(previous["costs"]["هدر"], previous["revenue"])

    kpis = [
        metric("الإيرادات", current["revenue"], previous["revenue"], "ريال"),
        metric("صافي الربح", current_profit, previous_profit, "ريال"),
        metric("هامش الربح", current_margin, previous_margin, "%", percentage_points=True),
        metric("تكلفة الطعام", current_food_cost, previous_food_cost, "%", percentage_points=True),
        metric("الهدر", current["costs"]["هدر"], previous["costs"]["هدر"], "ريال"),
        metric("عدد الطلبات", current["orders"], previous["orders"], "طلب"),
        metric("متوسط قيمة الطلب", current["revenue"] / current["orders"], previous["revenue"] / previous["orders"], "ريال"),
        metric("متوسط فاتورة المصروف", current_costs / current["expense_invoices"], previous_costs / previous["expense_invoices"], "ريال"),
    ]

    cost_variance = []
    budget_share = {"مواد غذائية": 38.0, "رواتب": 17.0, "إيجارات": 11.5, "توصيل": 7.0, "هدر": 1.8}
    for category, current_value in current["costs"].items():
        previous_value = previous["costs"][category]
        budget_value = round(current["revenue"] * budget_share[category] / 100, 2)
        cost_variance.append({
            "category": category,
            "current": current_value,
            "previous": previous_value,
            "difference": round(current_value - previous_value, 2),
            "change_pct": change_pct(current_value, previous_value),
            "share_pct": pct(current_value, current["revenue"]),
            "budget_share_pct": budget_share[category],
            "budget_value": budget_value,
            "budget_variance": round(current_value - budget_value, 2),
            "over_budget": current_value > budget_value,
        })

    branches = []
    for branch in model["branches"]:
        row = dict(branch)
        row["profit_change_pct"] = change_pct(row["current_profit"], row["previous_profit"])
        row["margin_variance"] = round(row["current_margin"] - row["previous_margin"], 1)
        branches.append(row)

    products = []
    for product in model["products"]:
        row = dict(product)
        row["profit_change_pct"] = change_pct(row["profit"], row["previous_profit"])
        row["profit_per_order"] = round(row["profit"] / row["orders"], 2)
        products.append(row)

    carriers = []
    for carrier in model["carriers"]:
        row = dict(carrier)
        row["orders_change_pct"] = change_pct(row["orders"], row["previous_orders"])
        row["fee_per_order"] = round(row["fees"] / row["orders"], 2)
        row["previous_fee_per_order"] = round(row["previous_fees"] / row["previous_orders"], 2)
        row["fee_per_order_change_pct"] = change_pct(row["fee_per_order"], row["previous_fee_per_order"])
        carriers.append(row)

    revenue_gain = round(current["revenue"] - previous["revenue"], 2)
    added_cost = round(current_costs - previous_costs, 2)
    profit_change = round(current_profit - previous_profit, 2)

    return {
        "synthetic": True,
        "period": {"current": current["label"], "previous": previous["label"]},
        "kpis": kpis,
        "summary": {
            "headline": "المبيعات ارتفعت، لكن الربحية انخفضت",
            "detail": f"ارتفعت الإيرادات {change_pct(current['revenue'], previous['revenue'])}%، بينما تراجع صافي الربح {abs(change_pct(current_profit, previous_profit))}% بسبب نمو تكلفة الطعام والهدر ورسوم التوصيل أسرع من الإيراد.",
            "health_score": 71,
            "alerts": 4,
        },
        "profit_bridge": {
            "previous_profit": previous_profit,
            "revenue_gain": revenue_gain,
            "added_cost": added_cost,
            "profit_change": profit_change,
            "current_profit": current_profit,
        },
        "cost_variance": cost_variance,
        "trend": list(model["trend"]),
        "branches": branches,
        "products": products,
        "carriers": carriers,
        "highlights": {
            "best_product": max(products, key=lambda row: row["profit"]),
            "lowest_product": min(products, key=lambda row: row["profit"]),
            "high_sales_low_margin": sorted(products, key=lambda row: (-row["orders"], row["margin_pct"]))[:3],
            "best_branch": max(branches, key=lambda row: row["current_profit"]),
            "weakest_branch": min(branches, key=lambda row: row["current_profit"]),
        },
        "alerts": [
            {"level": "high", "title": "تكلفة الطعام أعلى من حد المتابعة", "detail": f"ارتفعت من {previous_food_cost}% إلى {current_food_cost}% (+{round(current_food_cost - previous_food_cost, 1)} نقطة)."},
            {"level": "high", "title": "الهدر ينمو أسرع من الإيراد", "detail": f"ارتفع الهدر {change_pct(current['costs']['هدر'], previous['costs']['هدر'])}% مقابل نمو الإيراد {change_pct(current['revenue'], previous['revenue'])}%."},
            {"level": "medium", "title": "تراجع هامش الربح", "detail": f"انخفض الهامش من {previous_margin}% إلى {current_margin}% ({round(current_margin - previous_margin, 1)} نقطة)."},
            {"level": "medium", "title": "فرع ٢ يحتاج تدخلًا", "detail": "أعلى هدر وأكبر تراجع شهري في الربح بين الفروع."},
        ],
        "analysis_rules": [
            {"metric": "تكلفة الطعام", "rule": "تنبيه عند تجاوز 38% من الإيراد", "actual": f"{current_food_cost}%", "status": "متجاوز"},
            {"metric": "الهدر", "rule": "تنبيه عند تجاوز 1.8% من الإيراد", "actual": f"{current_waste}%", "status": "متجاوز"},
            {"metric": "هامش الربح", "rule": "تنبيه عند الانخفاض عن 24%", "actual": f"{current_margin}%", "status": "دون الحد"},
            {"metric": "رسوم التوصيل", "rule": "تنبيه عند تجاوز 7% من الإيراد", "actual": f"{pct(current['costs']['توصيل'], current['revenue'])}%", "status": "متجاوز"},
            {"metric": "تراجع الفرع", "rule": "تنبيه عند انخفاض الربح أكثر من 10% شهريًا", "actual": "فرع ٢: -45.5%", "status": "متجاوز"},
            {"metric": "ربحية الصنف", "rule": "مراجعة الصنف إذا الهامش أقل من 10% مع طلبات مرتفعة", "actual": "4 أصناف", "status": "مراجعة"},
        ],
        "actions": [
            {"priority": "أولوية ١", "action": "مراجعة وصفات وتوريد المواد الغذائية", "impact": "استعادة جزء من 3.7 نقطة المفقودة في تكلفة الطعام"},
            {"priority": "أولوية ٢", "action": "خطة خفض هدر فرع ٢", "impact": "تقليل الانحراف التشغيلي الأعلى بين الفروع"},
            {"priority": "أولوية ٣", "action": "إعادة تسعير الأصناف عالية البيع منخفضة الهامش", "impact": "رفع الربح دون الحاجة إلى طلبات إضافية"},
            {"priority": "أولوية ٤", "action": "مراجعة رسوم شركات التوصيل لكل طلب", "impact": "منع نمو رسوم القنوات أسرع من نمو الطلبات"},
        ],
    }


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
