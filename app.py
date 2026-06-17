from __future__ import annotations

import os
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, request, send_file, send_from_directory, session, url_for


ROOT = Path(__file__).resolve().parent
WEB_DIR = ROOT / "web"
REPORT_PATH = ROOT / "output" / "lead6_report.xlsx"
AUTH_USER = os.environ.get("LEAD_AUTH_USER", "")
AUTH_PASS = os.environ.get("LEAD_AUTH_PASS", "")

app = Flask(__name__)
app.secret_key = os.environ.get("LEAD_SESSION_SECRET", os.environ.get("SECRET_KEY", "lead-local-session"))


def _auth_enabled() -> bool:
    return bool(AUTH_USER and AUTH_PASS)


def _check_credentials(user: str, password: str) -> bool:
    if not _auth_enabled():
        return bool(user.strip() and password.strip())
    return user == AUTH_USER and password == AUTH_PASS


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
    if _dashboard_dir() is not None:
        return send_from_directory(WEB_DIR, "index.html")
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
    return {"status": "ok"}


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    app.run(host="0.0.0.0", port=port, debug=False)
