from __future__ import annotations

import hmac
import hashlib
import json
import mimetypes
import os
import sqlite3
import time
from urllib.parse import quote
from collections import defaultdict, deque
from pathlib import Path

from flask import Flask, Response, jsonify, redirect, request, send_file, send_from_directory, session

from analysis_engine import demo_analysis, financial_comparison


ROOT = Path(__file__).resolve().parent
DASHBOARD = ROOT / "dashboard"
WORKBOOK = ROOT / "GF_Demo.xlsx"
AUTH_USER = os.environ.get("GF_DEMO_AUTH_USER", "")
AUTH_PASS = os.environ.get("GF_DEMO_AUTH_PASS", "")
SESSION_SECRET = os.environ.get("GF_DEMO_SESSION_SECRET", "")
DATA_DIR = Path(os.environ.get("GF_DEMO_DATA_DIR", ROOT / "data"))
INVOICE_DB = DATA_DIR / "expense_invoices.sqlite3"

if not AUTH_USER or not AUTH_PASS or not SESSION_SECRET:
    raise RuntimeError("GF Demo authentication environment variables are required")

app = Flask(__name__)
app.secret_key = SESSION_SECRET
app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Strict",
    SESSION_COOKIE_SECURE=os.environ.get("GF_DEMO_SECURE_COOKIE", "0") == "1",
    PERMANENT_SESSION_LIFETIME=3600,
    MAX_CONTENT_LENGTH=10 * 1024 * 1024,
)

failed_attempts: dict[str, deque[float]] = defaultdict(deque)
WINDOW_SECONDS = 300
MAX_ATTEMPTS = 8
EXPENSE_TYPES = (
    "مواد غذائية", "قهوة ومشروبات", "تعبئة وتغليف", "إيجارات", "رواتب",
    "كهرباء ومياه", "صيانة", "توصيل", "تسويق", "خدمات تقنية",
    "رسوم بنكية", "مصروفات تشغيلية", "غير معروف",
)


def invoice_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(INVOICE_DB)
    conn.row_factory = sqlite3.Row
    conn.execute(
        """CREATE TABLE IF NOT EXISTS expense_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            signature TEXT NOT NULL UNIQUE,
            display_name TEXT NOT NULL,
            source_file TEXT NOT NULL,
            mime_type TEXT NOT NULL,
            file_data BLOB NOT NULL,
            classification TEXT NOT NULL,
            expense_type TEXT NOT NULL,
            supplier TEXT,
            invoice_number TEXT,
            invoice_date TEXT,
            amount REAL,
            extracted_json TEXT NOT NULL DEFAULT '{}',
            uploaded_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
        )"""
    )
    conn.commit()
    return conn


def canonical_signature(filename: str, data: bytes) -> str:
    if filename.lower().endswith(".xlsx"):
        try:
            from io import BytesIO
            from datetime import date, datetime, time as dt_time
            from decimal import Decimal
            import openpyxl

            def stable(value):
                if value is None:
                    return None
                if isinstance(value, (datetime, date, dt_time)):
                    return [type(value).__name__, value.isoformat()]
                if isinstance(value, bool):
                    return ["bool", value]
                if isinstance(value, (int, float, Decimal)):
                    return ["number", format(Decimal(str(value)).normalize(), "f")]
                return ["text", str(value).strip()]

            wb = openpyxl.load_workbook(BytesIO(data), read_only=True, data_only=False)
            try:
                sheets = []
                for ws in wb.worksheets:
                    rows = [[stable(v) for v in row] for row in ws.iter_rows(values_only=True)]
                    while rows and all(v is None for v in rows[-1]):
                        rows.pop()
                    for row in rows:
                        while row and row[-1] is None:
                            row.pop()
                    sheets.append([ws.title, ws.sheet_state, rows])
                payload = json.dumps(sheets, ensure_ascii=False, separators=(",", ":"))
                return hashlib.sha256(payload.encode("utf-8")).hexdigest()
            finally:
                wb.close()
        except Exception:
            pass
    return hashlib.sha256(data).hexdigest()


def classify_invoice(filename: str) -> str:
    value = filename.lower().replace("_", " ").replace("-", " ")
    rules = (
        ("مواد غذائية", ("تموين", "غذاء", "food", "meat", "خضار", "لحوم")),
        ("قهوة ومشروبات", ("قهوة", "coffee", "مشروب", "beverage")),
        ("تعبئة وتغليف", ("تغليف", "عبوات", "packaging", "cups")),
        ("إيجارات", ("إيجار", "rent")), ("صيانة", ("صيانة", "maintenance")),
        ("توصيل", ("توصيل", "delivery")), ("تسويق", ("تسويق", "marketing")),
        ("خدمات تقنية", ("software", "تقنية", "نظام", "subscription")),
        ("رسوم بنكية", ("bank", "بنك", "رسوم تحويل")),
    )
    return next((category for category, words in rules if any(word in value for word in words)), "غير معروف")


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


@app.get("/api/demo-analysis")
def get_demo_analysis():
    if not authenticated():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    return jsonify({"ok": True, "analysis": demo_analysis()})


@app.get("/api/financial-analysis")
def get_financial_analysis():
    if not authenticated():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    return jsonify({"ok": True, "analysis": financial_comparison()})


@app.get("/api/expense-invoices")
def list_expense_invoices():
    if not authenticated():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    with invoice_db() as conn:
        rows = conn.execute(
            """SELECT id, display_name, source_file, classification, expense_type,
                      supplier, invoice_number, invoice_date, amount, uploaded_at
               FROM expense_invoices ORDER BY uploaded_at DESC, id DESC"""
        ).fetchall()
    return jsonify({"ok": True, "invoices": [dict(row) for row in rows],
                    "expense_types": EXPENSE_TYPES})


@app.post("/api/expense-invoices")
def upload_expense_invoice():
    if not authenticated():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    upload = request.files.get("file")
    if not upload or not upload.filename:
        return jsonify({"ok": False, "error": "file_required"}), 400
    suffix = Path(upload.filename).suffix.lower()
    if suffix not in {".pdf", ".png", ".jpg", ".jpeg", ".xlsx"}:
        return jsonify({"ok": False, "error": "unsupported_file_type"}), 400
    data = upload.read()
    if not data:
        return jsonify({"ok": False, "error": "empty_file"}), 400
    signature = canonical_signature(upload.filename, data)
    classification = classify_invoice(upload.filename)
    display_name = Path(upload.filename).stem[:160] or "فاتورة بدون اسم"
    mime_type = upload.mimetype or mimetypes.guess_type(upload.filename)[0] or "application/octet-stream"
    try:
        with invoice_db() as conn:
            cursor = conn.execute(
                """INSERT INTO expense_invoices
                   (signature, display_name, source_file, mime_type, file_data,
                    classification, expense_type)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (signature, display_name, upload.filename[:255], mime_type, data,
                 classification, classification),
            )
            invoice_id = cursor.lastrowid
    except sqlite3.IntegrityError:
        with invoice_db() as conn:
            existing = conn.execute(
                """SELECT id, display_name, source_file, uploaded_at
                   FROM expense_invoices WHERE signature=?""", (signature,)
            ).fetchone()
        return jsonify({"ok": False, "error": "invoice_already_uploaded",
                        "invoice": dict(existing)}), 409
    return jsonify({"ok": True, "id": invoice_id, "classification": classification})


@app.patch("/api/expense-invoices/<int:invoice_id>")
def update_expense_invoice(invoice_id):
    if not authenticated():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    payload = request.get_json(silent=True) or {}
    display_name = str(payload.get("display_name") or "").strip()[:160]
    expense_type = str(payload.get("expense_type") or "غير معروف").strip()
    supplier = str(payload.get("supplier") or "").strip()[:160]
    invoice_number = str(payload.get("invoice_number") or "").strip()[:100]
    invoice_date = str(payload.get("invoice_date") or "").strip()[:10]
    try:
        amount = float(payload["amount"]) if payload.get("amount") not in (None, "") else None
    except (TypeError, ValueError):
        return jsonify({"ok": False, "error": "invalid_amount"}), 400
    if not display_name or expense_type not in EXPENSE_TYPES or (amount is not None and amount < 0):
        return jsonify({"ok": False, "error": "invalid_invoice_metadata"}), 400
    with invoice_db() as conn:
        cursor = conn.execute(
            """UPDATE expense_invoices
               SET display_name=?, expense_type=?, supplier=?, invoice_number=?,
                   invoice_date=?, amount=?, updated_at=CURRENT_TIMESTAMP
               WHERE id=?""",
            (display_name, expense_type, supplier, invoice_number, invoice_date, amount, invoice_id),
        )
        if cursor.rowcount == 0:
            return jsonify({"ok": False, "error": "invoice_not_found"}), 404
    return jsonify({"ok": True})


@app.get("/api/expense-invoices/<int:invoice_id>/file")
def view_expense_invoice(invoice_id):
    if not authenticated():
        return jsonify({"ok": False, "error": "unauthorized"}), 401
    with invoice_db() as conn:
        row = conn.execute(
            "SELECT source_file, mime_type, file_data FROM expense_invoices WHERE id=?",
            (invoice_id,),
        ).fetchone()
    if row is None:
        return jsonify({"ok": False, "error": "invoice_not_found"}), 404
    return Response(row["file_data"], mimetype=row["mime_type"], headers={
        "Content-Disposition": f"inline; filename*=UTF-8''{quote(row['source_file'])}"
    })


@app.get("/health")
def health():
    return jsonify({"status": "ok", "service": "gf-demo"})


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=int(os.environ.get("PORT", "8768")), debug=False)
