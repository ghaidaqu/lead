from __future__ import annotations

from datetime import date


BANK_TRANSFER_FEES = [
    {
        "date": date(2026, 5, 30),
        "expense_type": "رسوم حوالة فورية صادرة",
        "amount": 0.58,
        "source": "كشف حساب جاري-3.PDF",
    },
    {
        "date": date(2026, 6, 3),
        "expense_type": "رسوم حوالة فورية صادرة",
        "amount": 1.15,
        "source": "كشف حساب جاري.PDF",
    },
    {
        "date": date(2026, 6, 10),
        "expense_type": "رسوم حوالة فورية صادرة",
        "amount": 0.58,
        "source": "كشف حساب جاري.PDF",
    },
    {
        "date": date(2026, 6, 13),
        "expense_type": "رسوم حوالة فورية صادرة",
        "amount": 0.58,
        "source": "كشف حساب جاري-2.PDF",
    },
    {
        "date": date(2026, 6, 22),
        "expense_type": "رسوم حوالة فورية صادرة",
        "amount": 1.15,
        "source": "كشف حساب جاري-2.PDF",
    },
]


def filtered_bank_transfer_fees(date_from=None, date_to=None):
    rows = []
    for row in BANK_TRANSFER_FEES:
        day = row["date"]
        if date_from and day < date_from:
            continue
        if date_to and day > date_to:
            continue
        rows.append(
            {
                "date": day.isoformat(),
                "expense_type": row["expense_type"],
                "amount": row["amount"],
                "source": row["source"],
            }
        )
    return sorted(rows, key=lambda r: r["date"], reverse=True)


def summarize_bank_statement(date_from=None, date_to=None):
    rows = filtered_bank_transfer_fees(date_from, date_to)
    transfer_fees_total = round(sum(float(r["amount"] or 0) for r in rows), 2)
    return {
        "summary": {
            "deposits_total": 0.0,
            "expenses_total": transfer_fees_total,
            "transfer_fees_total": transfer_fees_total,
            "net_total": -transfer_fees_total,
            "expenses_count": len(rows),
        },
        "rows": rows,
    }
