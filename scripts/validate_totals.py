#!/usr/bin/env python3
"""Validate our DB totals/aggregates against the LIVE Lead site KPIs.

Logs in, reads the platform's own KPI cards (dashboard / reports / wallet), and
diffs them against what we compute from PostgreSQL — so we can be sure we use the
same base and logic as the site. Modelled on the sister project's validate_db.py.

    DATABASE_URL=... python3 scripts/validate_totals.py
"""

from __future__ import annotations

import datetime as dt
import http.cookiejar
import re
import sys
import urllib.request
from html.parser import HTMLParser
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

import scripts.sync_from_lead as S  # noqa: E402
from scripts import db_store  # noqa: E402
from web.generate_site import load_data_from_db  # noqa: E402

TOL_ABS = 3       # absolute tolerance (the site rounds amounts to whole riyals)
TOL_PCT = 0.015   # 1.5% relative tolerance for large amounts


class _Text(HTMLParser):
    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_data(self, data):
        data = data.strip()
        if data:
            self.parts.append(data)


def flat(html: str) -> str:
    p = _Text()
    p.feed(html)
    return re.sub(r"\s+", " ", " ".join(p.parts))


def after(text: str, label: str):
    m = re.search(re.escape(label) + r"\s*([\d,]+(?:\.\d+)?)", text)
    return float(m.group(1).replace(",", "")) if m else None


def before(text: str, label: str):
    m = re.search(r"([\d,]+(?:\.\d+)?)\s*ريال\s*" + re.escape(label), text)
    return float(m.group(1).replace(",", "")) if m else None


def login_opener(env):
    base = env["LEAD_BASE_URL"].rstrip("/")
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    _, _, lh = S.http_get_meta(opener, f"{base}/login.php")
    fp = S.FormParser(); fp.feed(lh)
    form = fp.forms[0]; post = {}
    for inp in form["inputs"]:
        name = inp.get("name"); typ = (inp.get("type") or "").lower()
        if not name:
            continue
        if typ == "password":
            post[name] = env["LEAD_PASSWORD"]
        elif typ in ("text", "email") or "user" in name.lower() or "login" in name.lower():
            post[name] = env["LEAD_USERNAME"]
        elif typ == "hidden":
            post[name] = inp.get("value", "")
    for btn in form.get("buttons", []):
        if btn.get("name"):
            post.setdefault(btn["name"], btn.get("value", "") or "1")
    action = form.get("action") or "/login.php"
    if not action.startswith("http"):
        action = f"{base}/{action.lstrip('/')}"
    S.http_post_meta(opener, action, post)
    return opener, base


def db_counts(conn):
    with conn.cursor() as cur:
        cur.execute("SELECT count(*) c FROM shipments")
        total = int(cur.fetchone()["c"])
        cur.execute("SELECT count(*) c FROM shipments WHERE status LIKE %s OR status LIKE %s", ("%توصيل%", "%تسليم%"))
        delivered = int(cur.fetchone()["c"])
        cur.execute("SELECT count(*) c FROM shipments WHERE status = ANY(%s)",
                    (["مؤكد", "قيد الشحن", "قيد التوصيل", "تم الاستلام", "معلق"],))
        inprog = int(cur.fetchone()["c"])
        cur.execute("SELECT count(DISTINCT merchant_name) c FROM shipments WHERE merchant_name <> ''")
        customers = int(cur.fetchone()["c"])
    return total, delivered, inprog, customers


def main() -> int:
    env = S.load_env(PROJECT_DIR / ".env")
    if not db_store.db_enabled():
        print("DATABASE_URL not configured.", file=sys.stderr)
        return 2

    opener, base = login_opener(env)
    dash = flat(S.http_get(opener, f"{base}/admin/dashboard.php"))
    rep = flat(S.http_get(opener, f"{base}/admin/reports.php"))
    wal = flat(S.http_get(opener, f"{base}/admin/wallet.php"))

    # reports.php is the current-month view → compute our equivalents for this month.
    today = dt.date.today()
    month_from = today.replace(day=1)
    t_month = load_data_from_db(month_from, today)[0]
    fin_all = load_data_from_db()[8]                     # all-time finance
    deposits_all = fin_all["bank"]["total"] + fin_all["moyasar"]["total"] + fin_all["other"]["total"]

    with db_store.get_conn() as conn:
        total, delivered, inprog, customers = db_counts(conn)

    # is_amount=True → a small rounding tolerance is fine (the site rounds to whole
    # riyals); counts must match exactly.
    checks = [
        ("اجمالي الطلبات (الكل)",   after(dash, "اجمالي الطلبات"),        total, False),
        ("تم التسليم (الكل)",        after(dash, "تم التسليم"),            delivered, False),
        ("قيد التنفيذ (الكل)",       after(dash, "قيد التنفيذ"),           inprog, False),
        ("إجمالي العملاء",          after(dash, "إجمالي العملاء"),        customers, False),
        ("إجمالي الشحنات (الشهر)",   after(rep, "إجمالي الشحنات"),         t_month["records"], False),
        ("إجمالي الإيرادات (الشهر)", after(rep, "إجمالي الإيرادات"),       round(t_month["revenue"], 2), True),
        ("صافي الأرباح (الشهر)",     after(rep, "صافي الأرباح"),           round(t_month["total"], 2), True),
        ("إجمالي الإيداعات (الكل)",  after(wal, "إجمالي الإيداعات"),       round(deposits_all, 2), True),
    ]

    wlab = max(len(c[0]) for c in checks)
    print(f"{'KPI'.ljust(wlab)}  {'site':>12}  {'ours':>12}  {'diff':>9}  match")
    print("-" * (wlab + 44))
    ok = 0
    for label, site, ours, is_amount in checks:
        if site is None or ours is None:
            diff, match, ds = None, False, "—"
        else:
            diff = ours - site
            tol = max(TOL_ABS, abs(site) * TOL_PCT) if is_amount else 0  # counts: exact
            match = abs(diff) <= tol
            ds = f"{diff:+,.0f}"
        ok += match
        sv = "—" if site is None else f"{site:,.0f}"
        print(f"{label.ljust(wlab)}  {sv:>12}  {ours:>12,.0f}  {ds:>9}   {'OK' if match else 'XX'}")
    print("-" * (wlab + 44))
    print(f"{ok}/{len(checks)} KPIs match (counts require an exact match)")
    return 0 if ok == len(checks) else 1


if __name__ == "__main__":
    raise SystemExit(main())
