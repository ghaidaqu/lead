from __future__ import annotations

import hmac
import os
import time
from collections import defaultdict, deque
from pathlib import Path

from flask import Flask, jsonify, redirect, request, send_file, send_from_directory, session


ROOT = Path(__file__).resolve().parent
DASHBOARD = ROOT / "dashboard"
WORKBOOK = ROOT / "GF_Demo.xlsx"
AUTH_USER = os.environ.get("GF_DEMO_AUTH_USER", "")
AUTH_PASS = os.environ.get("GF_DEMO_AUTH_PASS", "")
SESSION_SECRET = os.environ.get("GF_DEMO_SESSION_SECRET", "")

if not AUTH_USER or not AUTH_PASS or not SESSION_SECRET:
    raise RuntimeError("GF Demo authentication environment variables are required")

app = Flask(__name__)
app.secret_key = SESSION_SECRET
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Strict",
    SESSION_COOKIE_SECURE=os.environ.get("GF_DEMO_SECURE_COOKIE", "0") == "1",
    PERMANENT_SESSION_LIFETIME=3600,
    MAX_CONTENT_LENGTH=16 * 1024,
)

failed_attempts: dict[str, deque[float]] = defaultdict(deque)
WINDOW_SECONDS = 300
MAX_ATTEMPTS = 8


def client_key() -> str:
    forwarded = request.headers.get("X-Forwarded-For", "").split(",")[0].strip()
    return forwarded or request.remote_addr or "unknown"


def rate_limited(key: str) -> bool:
    now = time.monotonic()
    attempts = failed_attempts[key]
    while attempts and now - attempts[0] > WINDOW_SECONDS:
        attempts.popleft()
    return len(attempts) >= MAX_ATTEMPTS


def authenticated() -> bool:
    return bool(session.get("gf_demo_authenticated"))


@app.after_request
def security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Permissions-Policy"] = "camera=(), microphone=(), geolocation=()"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; script-src 'self' 'unsafe-inline' https://cdn.jsdelivr.net; "
        "style-src 'self' 'unsafe-inline'; img-src 'self' data:; connect-src 'self'; "
        "font-src 'self' data:; frame-ancestors 'none'; base-uri 'self'; form-action 'self'"
    )
    if request.path.startswith("/api/") or request.path.endswith(".xlsx"):
        response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/")
def home():
    return redirect("/dashboard/")


@app.get("/dashboard/")
@app.get("/dashboard/index.html")
def dashboard():
    return send_from_directory(DASHBOARD, "index.html")


@app.get("/dashboard/<path:filename>")
def dashboard_asset(filename: str):
    return send_from_directory(DASHBOARD, filename)


@app.post("/api/login")
def login():
    key = client_key()
    if rate_limited(key):
        return jsonify({"ok": False, "error": "too_many_attempts"}), 429
    payload = request.get_json(silent=True) or {}
    username = str(payload.get("username", ""))
    password = str(payload.get("password", ""))
    valid = hmac.compare_digest(username, AUTH_USER) and hmac.compare_digest(password, AUTH_PASS)
    if not valid:
        failed_attempts[key].append(time.monotonic())
        return jsonify({"ok": False, "error": "invalid_credentials"}), 401
    failed_attempts.pop(key, None)
    session.clear()
    session.permanent = True
    session["gf_demo_authenticated"] = True
    return jsonify({"ok": True})


@app.post("/api/logout")
def logout():
    session.clear()
    return jsonify({"ok": True})


@app.get("/api/session")
def session_status():
    return jsonify({"authenticated": authenticated()})


@app.get("/GF_Demo.xlsx")
def workbook():
    if not authenticated():
        return redirect("/dashboard/")
    return send_file(WORKBOOK, as_attachment=True, download_name="GF_Demo.xlsx")


@app.get("/health")
def health():
    return jsonify({"status": "ok", "service": "gf-demo"})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "8768")), debug=False)
