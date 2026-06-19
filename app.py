from __future__ import annotations

import os
import logging
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, request, send_file, send_from_directory, session, url_for, make_response

try:
    from web.generate_site import build_html, load_data_from_db
except Exception:  # pragma: no cover
    build_html = None
    load_data_from_db = None

try:
    from scripts.db_store import db_url as pg_db_url
except Exception:  # pragma: no cover
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


def _auth_enabled() -> bool:
    return bool(AUTH_USER and AUTH_PASS)


def _check_credentials(user: str, password: str) -> bool:
    if not _auth_enabled():
        return bool(user.strip() and password.strip())
    return user == AUTH_USER and password == AUTH_PASS


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
    if build_html is None or load_data_from_db is None:
        return "postgres_unavailable", False, "ImportError"
    try:
        from scripts.db_store import get_conn, compare_counts
        with get_conn() as conn:
            compare_counts(conn)
        return "postgres", True, None
    except Exception as exc:
        return "postgres_error", False, type(exc).__name__


def _dashboard_payload_from_db():
    totals, top_merchants, top_cities, top_carriers, statuses, daily, daily_revenue, daily_count, finance, cod_items, return_items, statement, period_label, missing_sequence_numbers = load_data_from_db()
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


@app.get("/dashboard/")
@app.get("/dashboard/index.html")
def dashboard_index() -> Response:
    if build_html is not None and load_data_from_db is not None:
        try:
            data = _dashboard_payload_from_db()
            _update_data_source("postgres", True)
            return Response(build_html(data), mimetype="text/html")
        except Exception as exc:
            _update_data_source("excel_fallback", False, type(exc).__name__)
    if _dashboard_dir() is not None:
        _update_data_source("excel_fallback", False)
        return send_from_directory(WEB_DIR, "index.html")
    _update_data_source("excel_fallback", False)
    return Response(_fallback_page(), mimetype="text/html")


@app.get("/dashboard/<path:filename>")
def dashboard_assets(filename: str):
    if _dashboard_dir() is None:
        return Response(_fallback_page(), mimetype="text/html")
    return send_from_directory(WEB_DIR, filename)


@app.get("/report.xlsx")
def report_download():
    if REPORT_PATH.exists():
        return send_file(REPORT_PATH, as_attachment=True, download_name="lead_report.xlsx")
    return Response("Report not found. Run scripts/build_report.py first.", status=404, mimetype="text/plain")


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
