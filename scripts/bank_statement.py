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


def _database_rows(date_from=None, date_to=None):
    try:
        from scripts import db_store
        if not db_store.db_enabled():
            return []
        with db_store.get_conn() as conn, conn.cursor() as cur:
            cur.execute("SELECT to_regclass('public.bank_transactions') AS table_name")
            if not cur.fetchone()["table_name"]:
                return []
            cur.execute(
                """SELECT transaction_date, description, amount, direction, category, source_file
                   FROM bank_transactions
                   WHERE approved = TRUE
                     AND (%s::date IS NULL OR transaction_date >= %s::date)
                     AND (%s::date IS NULL OR transaction_date <= %s::date)
                   ORDER BY transaction_date DESC, id DESC""",
                (date_from, date_from, date_to, date_to),
            )
            return [
                {
                    "date": row["transaction_date"].isoformat(),
                    "expense_type": row["category"],
                    "description": row["description"],
                    "amount": float(row["amount"]),
                    "direction": row["direction"],
                    "source": row["source_file"],
                }
                for row in cur.fetchall()
            ]
    except Exception:
        return []


def summarize_bank_statement(date_from=None, date_to=None):
    imported = _database_rows(date_from, date_to)
    rows = filtered_bank_transfer_fees(date_from, date_to)
    known = {(r["date"], r["expense_type"], round(float(r["amount"]), 2)) for r in rows}
    known_transfer_fees = {(r["date"], round(float(r["amount"]), 2)) for r in rows}
    rows.extend(
        r for r in imported
        if (r["date"], r["expense_type"], round(float(r["amount"]), 2)) not in known
        and not (
            r["expense_type"] == "رسوم حوالات"
            and (r["date"], round(float(r["amount"]), 2)) in known_transfer_fees
        )
    )
    rows.sort(key=lambda r: r["date"], reverse=True)
    deposits_total = round(sum(float(r["amount"] or 0) for r in rows if r.get("direction") == "credit"), 2)
    expenses_total = round(sum(float(r["amount"] or 0) for r in rows if r.get("direction", "debit") == "debit"), 2)
    return {
        "summary": {
            "deposits_total": deposits_total,
            "expenses_total": expenses_total,
            "transfer_fees_total": round(sum(float(r["amount"] or 0) for r in rows if r.get("category") == "رسوم حوالات" or "رسوم حوالة" in r.get("expense_type", "")), 2),
            "net_total": round(deposits_total - expenses_total, 2),
            "expenses_count": sum(1 for r in rows if r.get("direction", "debit") == "debit"),
        },
        "rows": rows,
    }
