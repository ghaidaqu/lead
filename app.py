from __future__ import annotations

import os
import json
import logging
from functools import wraps
from io import BytesIO
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, request, send_file, send_from_directory, session, url_for, make_response

try:
    from web.generate_site import load_data_from_db
except Exception:  # pragma: no cover
    load_data_from_db = None

try:
    from scripts import db_store
    from scripts.db_store import db_url as pg_db_url
except Exception:  # pragma: no cover
    db_store = None
    pg_db_url = lambda: None


ROOT = Path(__file__).resolve().parent
WEB_DIR = ROOT / "web"
REPORT_PATH = ROOT / "output" / "lead6_report.xlsx"
AUTH_USER = os.environ.get("LEAD_AUTH_USER", "")
AUTH_PASS = os.environ.get("LEAD_AUTH_PASS", "")

app = Flask(__name__)
app.secret_key = os.environ.get("LEAD_SESSION_SECRET", os.environ.get("SECRET_KEY", "lead-local-session"))
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("lead")
LAST_DATA_SOURCE = "unknown"
LAST_POSTGRES_AVAILABLE = False
LAST_ERROR_TYPE = ""

# --- Lead reports (actuals) ---------------------------------------------------
# The dashboard sources cost/profit from Lead's own reports.php (exact for any
# period), not from a price sheet. The WORKER (whose Lead session is reliable)
# captures the reports per range into the `lead_reports` DB setting; here we just
# read that and match the selected range — no Lead login in the request path.
import datetime as _dt


def _range_label(date_from, date_to):
    """Map a selected (from,to) to the worker's semantic snapshot label.
    date_from/date_to are date objects or None. Returns 'all' | 'month' | '30d' |
    'YYYY-MM' (closed month) | None (custom range with no captured snapshot)."""
    today = _dt.date.today()
    if not date_from:
        return "all"
    end = date_to or today
    # this month (1st .. today)
    if date_from == today.replace(day=1) and end == today:
        return "month"
    # last 30 days
    if date_from == today - _dt.timedelta(days=29) and end == today:
        return "30d"
    # a full calendar month (1st .. last day)
    if date_from.day == 1:
        nxt = _dt.date(date_from.year + (date_from.month == 12),
                       1 if date_from.month == 12 else date_from.month + 1, 1)
        if end == nxt - _dt.timedelta(days=1):
            return f"{date_from.year:04d}-{date_from.month:02d}"
    return None


_PRESET_LABELS = {"all": "all", "month": "month", "30": "30d"}


def get_lead_reports(date_from, date_to, preset=None):
    """Lead's actual revenue/cost/profit (+ per-carrier/-merchant) for the selected
    range, read from the worker-captured `lead_reports` setting. `preset` (from the
    dashboard tab) maps directly and avoids any server/browser timezone mismatch;
    custom ranges fall back to month-boundary detection. Returns None when the range
    wasn't snapshotted (dashboard then keeps its own figure)."""
    if db_store is None:
        return None
    label = _PRESET_LABELS.get(preset) or _range_label(date_from, date_to)
    if label is None:
        return None
    try:
        with db_store.get_conn() as conn:
            raw = db_store.get_setting(conn, "lead_reports", "")
        store = json.loads(raw) if raw else None
    except Exception:
        return None
    if not store:
        return None
    return store.get("ranges", {}).get(label)


def get_actuals(date_from, date_to):
    """Per-shipment TRUE profit for the range, aggregated from the stored actuals.
    True profit = Policy charge - (Base Cost + Over + COD), i.e. it subtracts the
    over-weight/COD fees Lead pays نفوذ but never bills merchants. Also returns the
    reports.php-style profit (charge - Base) and the absorbed fees, so the dashboard
    can show both the headline number and the leakage."""
    if db_store is None or not db_store.db_enabled():
        return None
    start = date_from.isoformat() if date_from else "2026-02-01"
    end = date_to.isoformat() if date_to else _dt.date.today().isoformat()
    try:
        with db_store.get_conn() as conn, conn.cursor() as cur:
            cur.execute(
                """SELECT sum(actual_revenue) rev,
                          sum(actual_base_cost) FILTER (WHERE actual_revenue>0) base,
                          sum(actual_extra_cost) extra, sum(actual_profit) profit,
                          count(*) FILTER (WHERE actual_revenue>0) n
                   FROM shipments WHERE shipment_date >= %s AND shipment_date <= %s""",
                (start, end),
            )
            r = cur.fetchone()
    except Exception:
        return None
    rev = float(r["rev"] or 0); base = float(r["base"] or 0)
    extra = float(r["extra"] or 0); profit = float(r["profit"] or 0)
    return {
        "revenue": round(rev, 2),
        "base_cost": round(base, 2),
        "absorbed_fees": round(extra, 2),
        "reports_profit": round(rev - base, 2),
        "true_profit": round(profit, 2),
        "count": int(r["n"] or 0),
    }


def get_daily_profit(date_from, date_to):
    """Daily profit using the stored per-shipment profit rule."""
    if db_store is None or not db_store.db_enabled():
        return []
    start = date_from.isoformat() if date_from else "2026-02-01"
    end = date_to.isoformat() if date_to else _dt.date.today().isoformat()
    try:
        with db_store.get_conn() as conn, conn.cursor() as cur:
            cur.execute(
                """SELECT shipment_date::date day,
                          COALESCE(sum(actual_profit) FILTER (WHERE included_in_profit), 0) AS profit
                   FROM shipments
                   WHERE shipment_date >= %s AND shipment_date <= %s
                   GROUP BY shipment_date::date
                   ORDER BY shipment_date::date""",
                (start, end),
            )
            rows = cur.fetchall()
    except Exception:
        return []
    return [[r["day"].isoformat(), round(float(r["profit"] or 0), 2)] for r in rows]


def _auth_enabled() -> bool:
    return bool(AUTH_USER and AUTH_PASS)


def _check_credentials(user: str, password: str) -> bool:
    # Fail closed: if no credentials are configured, deny all logins rather than
    # accepting any non-empty pair (the previous behaviour was a security hole).
    if not _auth_enabled():
        logger.warning("login denied: LEAD_AUTH_USER/LEAD_AUTH_PASS not configured")
        return False
    return user == AUTH_USER and password == AUTH_PASS


def require_auth(view):
    """Gate write endpoints behind an authenticated session."""
    @wraps(view)
    def wrapped(*args, **kwargs):
        if not session.get("lead_authenticated"):
            return jsonify({"ok": False, "error": "unauthorized"}), 401
        return view(*args, **kwargs)
    return wrapped


def _update_data_source(source: str, postgres_available: bool, error_type: str = "") -> None:
    global LAST_DATA_SOURCE, LAST_POSTGRES_AVAILABLE, LAST_ERROR_TYPE
    LAST_DATA_SOURCE = source
    LAST_POSTGRES_AVAILABLE = postgres_available
    LAST_ERROR_TYPE = error_type
    logger.info("DATA_SOURCE=%s postgres_available=%s%s", source, str(postgres_available).lower(), f" error_type={error_type}" if error_type else "")


def _probe_postgres() -> tuple[str, bool, str | None]:
    url = pg_db_url()
    if not url:
        return "missing_database_url", False, None
    if load_data_from_db is None:
        return "postgres_unavailable", False, "ImportError"
    try:
        from scripts.db_store import get_conn, compare_counts
        with get_conn() as conn:
            compare_counts(conn)
        return "postgres", True, None
    except Exception as exc:
        return "postgres_error", False, type(exc).__name__


def _parse_date_arg(value):
    import datetime as _dt
    if not value:
        return None
    try:
        return _dt.date.fromisoformat(value.strip()[:10])
    except (ValueError, AttributeError):
        return None


def _resolve_date_range(date_from, date_to, preset=None):
    """Resolve dashboard/export presets to concrete dates.

    Explicit custom from/to values win; presets only fill an otherwise empty
    range so existing custom filters keep behaving as selected by the user.
    """
    if date_from or date_to:
        return date_from, date_to
    today = _dt.date.today()
    if preset == "month":
        return today.replace(day=1), today
    if preset == "30":
        return today - _dt.timedelta(days=29), today
    return date_from, date_to


def _dashboard_payload_from_db(date_from=None, date_to=None):
    totals, top_merchants, top_cities, top_carriers, statuses, daily, daily_revenue, daily_count, finance, cod_items, return_items, statement, period_label, missing_sequence_numbers = load_data_from_db(date_from, date_to)
    return {
        "totals": totals,
        "top_merchants": top_merchants,
        "top_cities": top_cities,
        "daily": daily,
        "daily_revenue": daily_revenue,
        "daily_count": daily_count,
        "finance": finance,
        "cod_items": cod_items,
        "return_items": return_items,
        "top_carriers": top_carriers,
        "statement": statement,
        "period_label": period_label,
        "missing_sequence_numbers": missing_sequence_numbers,
        "statuses": statuses,
    }


def _json_safe(value):
    import datetime as _dt
    from decimal import Decimal
    if isinstance(value, Decimal):
        return float(value)
    if isinstance(value, (_dt.datetime, _dt.date)):
        return value.isoformat()
    return str(value)


@app.get("/api/dashboard")
@require_auth
def api_dashboard():
    if load_data_from_db is None:
        return jsonify({"ok": False, "error": "dashboard_unavailable"}), 503
    try:
        date_from = _parse_date_arg(request.args.get("from"))
        date_to = _parse_date_arg(request.args.get("to"))
        date_from, date_to = _resolve_date_range(date_from, date_to, request.args.get("preset"))
        payload = _dashboard_payload_from_db(date_from, date_to)
        # our own current-month figures, so the dashboard can always show a
        # like-for-like comparison against the platform's month KPIs (site_kpis),
        # no matter which range the user is currently viewing.
        try:
            import datetime as _dt
            _today = _dt.date.today()
            _mtotals = load_data_from_db(_today.replace(day=1), _today)[0]
            payload["month_self"] = {
                "revenue": _mtotals.get("revenue"),
                "profit": _mtotals.get("total"),
                "count": _mtotals.get("records"),
                "cod_amount": _mtotals.get("cod_amount"),
            }
        except Exception:
            payload["month_self"] = None
        site_kpis = None
        if db_store is not None:
            try:
                with db_store.get_conn() as conn:
                    raw = db_store.get_setting(conn, "site_kpis", "")
                site_kpis = json.loads(raw) if raw else None
            except Exception:
                site_kpis = None
        payload["site_kpis"] = site_kpis
        # Lead's own actuals for the selected range (exact cost/profit, no sheet).
        try:
            payload["lead_reports"] = get_lead_reports(date_from, date_to, request.args.get("preset"))
        except Exception:
            payload["lead_reports"] = None
        try:
            payload["actuals"] = get_actuals(date_from, date_to)
        except Exception:
            payload["actuals"] = None
        try:
            payload["daily_profit"] = get_daily_profit(date_from, date_to)
            payload["daily_actual_profit"] = payload["daily_profit"]
        except Exception:
            payload["daily_profit"] = []
            payload["daily_actual_profit"] = []
        _update_data_source("postgres", True)
    except Exception as exc:
        logger.exception("dashboard payload failed")
        _update_data_source("error", False, type(exc).__name__)
        return jsonify({"ok": False, "error": type(exc).__name__}), 500
    body = json.dumps({"ok": True, **payload}, ensure_ascii=False, default=_json_safe)
    return Response(body, mimetype="application/json")


@app.post("/api/login")
def api_login():
    payload = request.get_json(silent=True) or {}
    user = str(payload.get("username", ""))
    password = str(payload.get("password", ""))
    if _check_credentials(user, password):
        session["lead_authenticated"] = True
        session["lead_user"] = user
        return jsonify({"ok": True})
    return jsonify({"ok": False, "error": "invalid_credentials"}), 401


@app.post("/api/logout")
def api_logout():
    session.clear()
    return jsonify({"ok": True})


@app.get("/api/session")
def api_session():
    return jsonify({"authenticated": bool(session.get("lead_authenticated"))})


# ---------------------------------------------------------------------------
# Pricing admin API. Carrier prices are auto-synced from the Lead site and are
# read-only here; the COD fee, extra-kilo pricing, and per-merchant overrides
# are admin-editable and persisted in PostgreSQL.
# ---------------------------------------------------------------------------

@app.get("/api/pricing")
def api_pricing_get():
    """Read-only pricing reference. Everything is sourced automatically from Lead
    (carrier prices from shipping-companies.php, COD fee from cod-settings.php);
    nothing here is editable — profit is computed from Lead's actuals."""
    if db_store is None or not db_store.db_enabled():
        return jsonify({"ok": False, "error": "database_unavailable"}), 503
    cod_fee = None
    with db_store.get_conn() as conn:
        snap = db_store.load_pricing_snapshot(conn)
        data_floor = db_store.get_setting(conn, "min_sync_date", "")
        raw = db_store.get_setting(conn, "lead_reports", "")
    if raw:
        try:
            cod_fee = json.loads(raw).get("cod_fee")
        except Exception:
            cod_fee = None
    return jsonify({"ok": True, "carriers": snap.get("carriers", {}),
                    "cod_fee": cod_fee, "data_floor": data_floor})


@app.get("/admin")
@app.get("/admin/")
def admin_page() -> Response:
    # Same single-page app as the dashboard; the SPA opens the pricing view
    # client-side based on the /admin path (no separate page, no reload).
    spa = WEB_DIR / "dashboard" / "index.html"
    if spa.exists():
        return send_file(spa)
    return Response("Not found", status=404, mimetype="text/plain")


def _dashboard_dir() -> Path | None:
    index = WEB_DIR / "index.html"
    if index.exists():
        return WEB_DIR
    return None


def _fallback_page() -> str:
    report_status = "found" if REPORT_PATH.exists() else "missing"
    dashboard_status = "found" if (WEB_DIR / "index.html").exists() else "missing"
    return f"""<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>lead</title>
    <style>
      body {{
        margin: 0;
        font-family: Inter, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        background: #f6f1eb;
        color: #2f2724;
      }}
      .wrap {{
        max-width: 760px;
        margin: 64px auto;
        padding: 0 24px;
      }}
      .card {{
        background: rgba(255, 255, 255, 0.82);
        border: 1px solid rgba(96, 72, 63, 0.14);
        border-radius: 20px;
        padding: 28px;
        box-shadow: 0 20px 45px rgba(57, 38, 31, 0.08);
        backdrop-filter: blur(18px);
      }}
      h1 {{
        margin: 0 0 12px;
        font-size: 28px;
      }}
      p {{
        line-height: 1.7;
        margin: 10px 0;
      }}
      code {{
        background: #f0e7df;
        padding: 2px 8px;
        border-radius: 8px;
      }}
      .meta {{
        margin-top: 18px;
        color: #6b5a53;
        font-size: 14px;
      }}
      .btn {{
        display: inline-block;
        margin-top: 18px;
        padding: 12px 18px;
        border-radius: 999px;
        text-decoration: none;
        color: #fff;
        background: linear-gradient(135deg, #70584d, #a58374);
      }}
    </style>
  </head>
  <body>
    <main class="wrap">
      <section class="card">
        <h1>lead</h1>
        <p>Dashboard source is present when <code>web/index.html</code> exists.</p>
        <p>Report file is present when the generated workbook exists.</p>
        <a class="btn" href="{url_for("dashboard_index")}">Open dashboard</a>
        <div class="meta">dashboard: {dashboard_status} · workbook: {report_status}</div>
      </section>
    </main>
  </body>
</html>
"""


@app.get("/")
def home() -> Response:
    return redirect(url_for("dashboard_index"))


DASHBOARD_DIR = WEB_DIR / "dashboard"


@app.get("/dashboard/")
@app.get("/dashboard/index.html")
def dashboard_index() -> Response:
    # Static SPA — it fetches its data from GET /api/dashboard.
    spa = DASHBOARD_DIR / "index.html"
    if spa.exists():
        response = make_response(send_file(spa))
        response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
        response.headers["Pragma"] = "no-cache"
        response.headers["Expires"] = "0"
        return response
    return Response(_fallback_page(), mimetype="text/html")


@app.get("/dashboard/<path:filename>")
def dashboard_assets(filename: str):
    # Serve SPA-local files first (none yet), then shared assets/logos from web/.
    if (DASHBOARD_DIR / filename).exists():
        return send_from_directory(DASHBOARD_DIR, filename)
    if (WEB_DIR / filename).exists():
        return send_from_directory(WEB_DIR, filename)
    return Response("Not found", status=404, mimetype="text/plain")


@app.get("/report.xlsx")
def report_download():
    # Built on demand from PostgreSQL — no .xlsx is stored on disk.
    if db_store is None or not db_store.db_enabled():
        return Response("Database unavailable.", status=503, mimetype="text/plain")
    try:
        from scripts.export_report import build_report_xlsx
        date_from = _parse_date_arg(request.args.get("from"))
        date_to = _parse_date_arg(request.args.get("to"))
        date_from, date_to = _resolve_date_range(date_from, date_to, request.args.get("preset"))
        with db_store.get_conn() as conn:
            data = build_report_xlsx(conn, date_from, date_to)
    except Exception as exc:
        logger.exception("report export failed")
        return Response(f"Report export failed: {type(exc).__name__}", status=500, mimetype="text/plain")
    return send_file(
        BytesIO(data),
        as_attachment=True,
        download_name="lead_report.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/health")
def health():
    data_source, postgres_available, last_error_type = _probe_postgres()
    _update_data_source(data_source, postgres_available, last_error_type or "")
    response = make_response(jsonify({
        "status": "ok",
        "data_source": data_source,
        "postgres_available": postgres_available,
        "last_error_type": last_error_type,
    }))
    response.headers["Cache-Control"] = "no-store, no-cache, must-revalidate, max-age=0"
    response.headers["Pragma"] = "no-cache"
    response.headers["X-Data-Source"] = data_source
    return response


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port, debug=False)
