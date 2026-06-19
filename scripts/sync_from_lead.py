#!/usr/bin/env python3
from __future__ import annotations

import datetime as dt
import json
import os
import re
import shutil
import sys
import urllib.parse
import urllib.request
import http.cookiejar
import traceback
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl import Workbook

try:
    from scripts.db_store import compare_counts, db_enabled, ensure_schema, finish_sync_run, get_conn, make_sync_run, upsert_rows
except Exception:  # pragma: no cover
    compare_counts = None
    db_enabled = lambda: False  # type: ignore
    ensure_schema = None
    finish_sync_run = None
    get_conn = None
    make_sync_run = None
    upsert_rows = None


PROJECT_DIR = Path(__file__).resolve().parents[1]
REPORT_XLSX = PROJECT_DIR / "output" / "lead6_report.xlsx"
BACKUP_DIR = PROJECT_DIR / "backups"
FALLBACK_BACKUP_DIR = Path(os.environ.get("LEAD_BACKUP_DIR", Path("/private/tmp/lead6_backups")))
STATE_PATH = PROJECT_DIR / "sync_state.json"
ENV_PATH = PROJECT_DIR / ".env"
AUTH_DIR = PROJECT_DIR / ".auth"
AUTH_STATE_PATH = AUTH_DIR / "lead_state.json"
MIN_SYNC_DATE = dt.date(2026, 6, 18)


def load_env(path: Path) -> dict[str, str]:
    env: dict[str, str] = {k: v for k, v in os.environ.items() if k.startswith("LEAD_")}
    if not path.exists():
        return env
    for raw in path.read_text(encoding="utf-8").splitlines():
        line = raw.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        env[key.strip()] = value.strip().strip('"').strip("'")
    return env


def norm(v: Any) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).strip())


def money(v: Any) -> float:
    m = re.search(r"-?\d+(?:\.\d+)?", norm(v).replace(",", ""))
    return float(m.group()) if m else 0.0


def parse_date_text(value: Any) -> dt.date | None:
    text = norm(value)
    if not text:
        return None
    match = re.search(r"(20\d{2})[/-](\d{2})[/-](\d{2})", text)
    if match:
        return dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
    match = re.search(r"(\d{2})/(\d{2})", text)
    if match:
        month = int(match.group(1))
        day = int(match.group(2))
        year = 2026
        return dt.date(year, month, day)
    return None


def row_date_at_or_after(row: list[Any], indexes: tuple[int, ...], minimum: dt.date = MIN_SYNC_DATE) -> bool:
    for idx in indexes:
        if idx < len(row):
            parsed = parse_date_text(row[idx])
            if parsed and parsed >= minimum:
                return True
    return False


def make_backup(src: Path) -> Path:
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    candidates = [BACKUP_DIR, FALLBACK_BACKUP_DIR]
    last_error: Exception | None = None
    for folder in candidates:
        try:
            folder.mkdir(parents=True, exist_ok=True)
            dst = folder / f"{src.stem}-{stamp}{src.suffix}"
            shutil.copy2(src, dst)
            return dst
        except Exception as exc:
            last_error = exc
            continue
    if last_error is not None:
        raise last_error
    raise RuntimeError("Unable to create backup")


def desired_source_mode() -> str:
    mode = os.environ.get("LEAD_SYNC_SOURCE", "website").strip().lower()
    return mode if mode in {"website", "excel"} else "website"


def create_workbook_with_raw_sheets() -> Workbook:
    wb = Workbook()
    default = wb.active
    wb.remove(default)
    for name in ("Raw_Shipments", "Raw_Wallet", "Raw_Payments", "Raw_COD"):
        wb.create_sheet(title=name)
    return wb


def create_source_workbook_from_website(ship_rows: list[list[Any]], wallet_rows: list[list[Any]]) -> Workbook:
    wb = create_workbook_with_raw_sheets()
    for name in ("الشحنات", "العمليات المالية"):
        if name in wb.sheetnames:
            del wb[name]
    ship_sheet = wb.create_sheet("الشحنات")
    ops_sheet = wb.create_sheet("العمليات المالية")
    if ship_rows:
        write_sheet(ship_sheet, ship_rows)
    if wallet_rows:
        write_sheet(ops_sheet, wallet_rows)
    return wb


def load_auth_state() -> dict[str, Any]:
    if AUTH_STATE_PATH.exists():
        try:
            return json.loads(AUTH_STATE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_auth_state(state: dict[str, Any]) -> None:
    AUTH_DIR.mkdir(parents=True, exist_ok=True)
    AUTH_STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


def load_state() -> dict[str, Any]:
    if STATE_PATH.exists():
        try:
            return json.loads(STATE_PATH.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def save_state(state: dict[str, Any]) -> None:
    STATE_PATH.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")


class TableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.in_table = False
        self.in_row = False
        self.in_cell = False
        self.tables: list[list[list[str]]] = []
        self.current_table: list[list[str]] = []
        self.current_row: list[str] = []
        self.buf: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self.in_table = True
            self.current_table = []
        elif self.in_table and tag == "tr":
            self.in_row = True
            self.current_row = []
        elif self.in_row and tag in ("td", "th"):
            self.in_cell = True
            self.buf = []

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self.in_cell:
            self.current_row.append(norm("".join(self.buf)))
            self.in_cell = False
        elif tag == "tr" and self.in_row:
            if any(self.current_row):
                self.current_table.append(self.current_row[:])
            self.in_row = False
        elif tag == "table" and self.in_table:
            if self.current_table:
                self.tables.append(self.current_table[:])
            self.in_table = False

    def handle_data(self, data):
        if self.in_cell:
            self.buf.append(data)


def parse_tables(html: str) -> list[list[list[str]]]:
    parser = TableParser()
    parser.feed(html)
    return parser.tables


def write_sheet(ws, rows: list[list[Any]]) -> None:
    ws.delete_rows(1, ws.max_row)
    for row in rows:
        ws.append(row)


def find_sheet(wb, names: list[str]):
    wanted = {re.sub(r"\s+", "", n).replace("أ", "ا").replace("إ", "ا").replace("آ", "ا") for n in names}
    for s in wb.sheetnames:
        key = re.sub(r"\s+", "", s).replace("أ", "ا").replace("إ", "ا").replace("آ", "ا")
        if key in wanted:
            return wb[s]
    return wb.create_sheet(names[0])


def data_key(row: list[Any], idxs: tuple[int, ...]) -> str:
    return " | ".join(norm(row[i]) if i < len(row) else "" for i in idxs)


def upsert_sheet_rows(ws, incoming: list[list[Any]], key_indexes: tuple[int, ...]) -> tuple[int, int]:
    existing: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        key = data_key([ws.cell(r, c).value for c in range(1, ws.max_column + 1)], key_indexes)
        if key.strip("| "):
            existing[key] = r
    added = updated = 0
    for row in incoming:
        key = data_key(row, key_indexes)
        if not key.strip("| "):
            continue
        if key in existing:
            target = existing[key]
            for c, value in enumerate(row, start=1):
                ws.cell(target, c).value = value
            updated += 1
        else:
            ws.append(row)
            added += 1
    return added, updated


class FormParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.forms: list[dict[str, Any]] = []
        self._in_form = False
        self._form: dict[str, Any] = {}

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if tag == "form":
            self._in_form = True
            self._form = {"method": attrs.get("method", "get").upper(), "action": attrs.get("action", ""), "inputs": [], "buttons": []}
        elif self._in_form and tag == "input":
            self._form["inputs"].append(attrs)
        elif self._in_form and tag == "button":
            self._form["buttons"].append(attrs)

    def handle_endtag(self, tag):
        if tag == "form" and self._in_form:
            self.forms.append(self._form)
            self._in_form = False


class LinkParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.links: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag == "a":
            attrs = dict(attrs)
            href = attrs.get("href")
            if href:
                self.links.append(href)


def http_get_meta(opener, url: str) -> tuple[str, int, str]:
    req = urllib.request.Request(url)
    with opener.open(req) as resp:
        html = resp.read().decode("utf-8", errors="ignore")
        final_url = getattr(resp, "geturl", lambda: url)()
        status = int(getattr(resp, "status", resp.getcode()))
        return final_url, status, html


def http_get(opener, url: str) -> str:
    _, _, html = http_get_meta(opener, url)
    return html


def http_post(opener, url: str, data: dict[str, str]) -> str:
    body = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with opener.open(req) as resp:
        return resp.read().decode("utf-8", errors="ignore")


def http_post_meta(opener, url: str, data: dict[str, str]) -> tuple[str, int, str]:
    body = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with opener.open(req) as resp:
        html = resp.read().decode("utf-8", errors="ignore")
        final_url = getattr(resp, "geturl", lambda: url)()
        status = int(getattr(resp, "status", resp.getcode()))
        return final_url, status, html


def log_page_sample(logs: list[dict[str, Any]], label: str, requested_url: str, final_url: str, status: int, html: str, reason: str = "") -> None:
    snippet = html[:500]
    logs.append({
        "kind": "page_sample",
        "page": label,
        "requested_url": requested_url,
        "final_url": final_url,
        "status": status,
        "reason": reason,
        "html_500": snippet,
    })


def scrape_site(env: dict[str, str]) -> dict[str, Any]:
    base = env["LEAD_BASE_URL"].rstrip("/")
    username = env["LEAD_USERNAME"]
    password = env["LEAD_PASSWORD"]
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(cj))
    logs: list[dict[str, Any]] = []
    auth_state = load_auth_state()

    def page_looks_logged_in(html: str) -> tuple[bool, dict[str, bool]]:
        parser = LinkParser()
        parser.feed(html)
        lower = html.lower()
        checks = {
            "has_admin_links": any(token in html for token in ("/admin/shipments.php", "/admin/wallet.php", "/admin/collect-cod.php", "/admin/dashboard.php")),
            "shipments_link": any("/admin/shipments.php" in link for link in parser.links),
            "wallet_link": any("/admin/wallet.php" in link for link in parser.links),
            "cod_link": any("/admin/collect-cod.php" in link for link in parser.links),
            "not_login_form": "password" not in lower or "/login" not in lower,
        }
        return all(checks.values()), checks

    def page_is_login_like(html: str, final_url: str) -> bool:
        lower = html.lower()
        return "login" in final_url.lower() or ("password" in lower and "username" in lower)

    def fetch_page(label: str, url: str) -> tuple[str, int, str]:
        final_url, status, html = http_get_meta(opener, url)
        reason = ""
        if final_url != url:
            reason = f"redirected_to={final_url}"
        elif page_is_login_like(html, final_url):
            reason = "login_page_detected"
        log_page_sample(logs, label, url, final_url, status, html, reason)
        return final_url, status, html

    if auth_state.get("cookie"):
        opener.addheaders = [("Cookie", auth_state["cookie"])]
        _, dashboard_status, cached_dashboard = fetch_page("dashboard.php", f"{base}/admin/dashboard.php")
        ok, cached_checks = page_looks_logged_in(cached_dashboard)
        logs.append({"kind": "probe", "page": "dashboard.php", "cached_session": ok, "checks": cached_checks})
        shipments_final_url, shipments_status, shipments_probe = fetch_page("shipments.php", f"{base}/admin/shipments.php")
        shipment_ok, shipment_checks = page_looks_logged_in(shipments_probe)
        logs.append({"kind": "probe", "page": "shipments.php", "cached_session": shipment_ok, "checks": shipment_checks, "final_url": shipments_final_url, "status": shipments_status})
        _, wallet_status, wallet_probe = fetch_page("wallet.php", f"{base}/admin/wallet.php")
        wallet_ok, wallet_checks = page_looks_logged_in(wallet_probe)
        logs.append({"kind": "probe", "page": "wallet.php", "cached_session": wallet_ok, "checks": wallet_checks, "status": wallet_status})
        _, cod_status, cod_probe = fetch_page("collect-cod.php", f"{base}/admin/collect-cod.php")
        cod_ok, cod_checks = page_looks_logged_in(cod_probe)
        logs.append({"kind": "probe", "page": "collect-cod.php", "cached_session": cod_ok, "checks": cod_checks, "status": cod_status})
        if ok and shipment_ok and wallet_ok and cod_ok:
            html = {
                "dashboard.php": cached_dashboard,
                "shipments.php": shipments_probe,
                "wallet.php": wallet_probe,
                "collect-cod.php": cod_probe,
                "sync-status.php": http_get(opener, f"{base}/admin/sync-status.php"),
            }
            logs.append({"kind": "session", "source": "saved_state", "used": True})
            return {
                "logs": logs,
                "html": html,
                "checks": {
                    "current_url": f"{base}/admin/dashboard.php",
                    "has_dashboard_links": ok,
                    "shipments_opened": shipment_ok,
                    "wallet_opened": wallet_ok,
                    "cod_opened": cod_ok,
                    "session_used": True,
                    "dashboard_checks": cached_checks,
                    "shipments_checks": shipment_checks,
                    "wallet_checks": wallet_checks,
                    "cod_checks": cod_checks,
                },
            }

    login_final_url, login_status, login_html = fetch_page("login.php", f"{base}/login.php")
    logs.append({"kind": "request", "url": f"{base}/login.php", "final_url": login_final_url, "status": login_status})
    parser = FormParser()
    parser.feed(login_html)
    form = parser.forms[0] if parser.forms else {"method": "POST", "action": "/login.php", "inputs": []}
    post_data: dict[str, str] = {}
    for inp in form["inputs"]:
        name = inp.get("name")
        if not name:
            continue
        typ = (inp.get("type") or "").lower()
        if typ == "password":
            post_data[name] = password
        elif typ in ("text", "email") or "user" in name.lower() or "login" in name.lower():
            post_data[name] = username
        elif typ == "hidden":
            post_data[name] = inp.get("value", "")
    for btn in form.get("buttons", []):
        btn_name = btn.get("name")
        if btn_name and btn.get("type", "").lower() in ("submit", ""):
            post_data.setdefault(btn_name, btn.get("value", "") or "1")
    action = form.get("action") or "/login.php"
    if not action.startswith("http"):
        action = f"{base}/{action.lstrip('/')}"
    login_post_final_url, login_post_status, login_result = http_post_meta(opener, action, post_data)
    login_result_html = login_result[:500]
    logs.append({
        "kind": "request",
        "url": action,
        "final_url": login_post_final_url,
        "status": login_post_status,
        "response_500": login_result_html,
    })

    dashboard_final_url, dashboard_status, dashboard_html = fetch_page("dashboard.php", f"{base}/admin/dashboard.php")
    logs.append({"kind": "request", "url": f"{base}/admin/dashboard.php", "final_url": dashboard_final_url, "status": dashboard_status})
    has_dashboard_links, dashboard_checks = page_looks_logged_in(dashboard_html)
    shipments_final_url, shipments_status, shipments_probe = fetch_page("shipments.php", f"{base}/admin/shipments.php")
    logs.append({"kind": "request", "url": f"{base}/admin/shipments.php", "final_url": shipments_final_url, "status": shipments_status})
    shipments_checks = {
        "status_200": shipments_status == 200,
        "redirected": shipments_final_url != f"{base}/admin/shipments.php",
        "login_like": page_is_login_like(shipments_probe, shipments_final_url),
    }
    shipments_ok = shipments_status == 200 and not shipments_checks["login_like"]
    wallet_final_url, wallet_status, wallet_probe = fetch_page("wallet.php", f"{base}/admin/wallet.php")
    logs.append({"kind": "request", "url": f"{base}/admin/wallet.php", "final_url": wallet_final_url, "status": wallet_status})
    wallet_checks = {
        "status_200": wallet_status == 200,
        "redirected": wallet_final_url != f"{base}/admin/wallet.php",
        "login_like": page_is_login_like(wallet_probe, wallet_final_url),
    }
    wallet_ok = wallet_status == 200 and not wallet_checks["login_like"]
    cod_final_url, cod_status, cod_probe = fetch_page("collect-cod.php", f"{base}/admin/collect-cod.php")
    logs.append({"kind": "request", "url": f"{base}/admin/collect-cod.php", "final_url": cod_final_url, "status": cod_status})
    cod_checks = {
        "status_200": cod_status == 200,
        "redirected": cod_final_url != f"{base}/admin/collect-cod.php",
        "login_like": page_is_login_like(cod_probe, cod_final_url),
    }
    cod_ok = cod_status == 200 and not cod_checks["login_like"]
    login_issue = None
    if not shipments_ok:
        login_issue = "shipments page was not accessible"
    elif not wallet_ok:
        login_issue = "wallet page was not accessible"
    elif not cod_ok:
        login_issue = "cod page was not accessible"
    if login_issue:
        logs.append({
            "kind": "login_issue",
            "page": "shipments.php" if not shipments_ok else "wallet.php" if not wallet_ok else "collect-cod.php",
            "reason": login_issue,
            "shipments": {"final_url": shipments_final_url, "status": shipments_status, "checks": shipments_checks},
            "wallet": {"final_url": wallet_final_url, "status": wallet_status, "checks": wallet_checks},
            "cod": {"final_url": cod_final_url, "status": cod_status, "checks": cod_checks},
        })

    saved_cookie = None
    for cookie in cj:
        if cookie.name == "PHPSESSID":
            saved_cookie = f"{cookie.name}={cookie.value}"
            break
    if saved_cookie:
        save_auth_state({"cookie": saved_cookie, "saved_at": dt.datetime.now().isoformat(timespec="seconds")})

    html: dict[str, str] = {
        "dashboard.php": dashboard_html,
        "shipments.php": shipments_probe,
        "wallet.php": wallet_probe,
        "collect-cod.php": cod_probe,
        "sync-status.php": http_get(opener, f"{base}/admin/sync-status.php"),
    }
    logs.append({"kind": "request", "url": f"{base}/admin/sync-status.php"})
    return {
        "logs": logs,
        "html": html,
        "checks": {
            "current_url": f"{base}/admin/dashboard.php",
            "has_dashboard_links": has_dashboard_links,
            "shipments_opened": shipments_ok,
            "wallet_opened": wallet_ok,
            "cod_opened": cod_ok,
            "session_used": False,
            "dashboard_checks": dashboard_checks,
            "shipments_checks": shipments_checks,
            "wallet_checks": wallet_checks,
            "cod_checks": cod_checks,
            "dashboard_status": dashboard_status,
            "shipments_status": shipments_status,
            "wallet_status": wallet_status,
            "cod_status": cod_status,
            "dashboard_final_url": dashboard_final_url,
            "shipments_final_url": shipments_final_url,
            "wallet_final_url": wallet_final_url,
            "cod_final_url": cod_final_url,
            "login_status": login_status,
            "login_final_url": login_final_url,
            "login_http_status": login_status,
            "login_redirected_to_login": "login" in login_final_url.lower(),
            "login_html_500": login_html[:500],
        },
    }


def normalize_headers(row: list[str]) -> list[str]:
    return [norm(v) for v in row]


def shipment_record(row: list[list[Any]]) -> dict[str, Any]:
    return {
        "order_id": norm(row[0]) if len(row) > 0 else "",
        "tracking_number": norm(row[1]) if len(row) > 1 else "",
        "merchant_name": norm(row[2]) if len(row) > 2 else "",
        "store_name": norm(row[3]) if len(row) > 3 else "",
        "customer_name": norm(row[4]) if len(row) > 4 else "",
        "city": norm(row[5]) if len(row) > 5 else "",
        "carrier": norm(row[10]) if len(row) > 10 else "",
        "payment_type": norm(row[9]) if len(row) > 9 else "",
        "status": norm(row[12]) if len(row) > 12 else "",
        "shipment_date": parse_date_text(row[13]) if len(row) > 13 else None,
        "delivery_date": parse_date_text(row[14]) if len(row) > 14 else None,
        "weight": money(row[11]) if len(row) > 11 else 0.0,
        "cod_amount": money(row[7]) if len(row) > 7 else 0.0,
        "shipping_charge": money(row[6]) if len(row) > 6 else 0.0,
        "raw_payload": row,
        "source_row": None,
        "source_hash": norm(row[0]) if row else "",
    }


def wallet_record(row: list[list[Any]], page: str, kind: str) -> dict[str, Any]:
    return {
        "transaction_key": data_key(row, (0, 1, 2, 3)),
        "transaction_date": parse_date_text(row[1]) if len(row) > 1 else None,
        "user_name": norm(row[2]) if len(row) > 2 else "",
        "description": norm(row[3]) if len(row) > 3 else "",
        "amount": money(row[4]) if len(row) > 4 else 0.0,
        "transaction_type": kind,
        "balance_before": None,
        "balance_after": None,
        "source_page": page,
        "raw_payload": row,
    }


def payment_record(row: list[list[Any]]) -> dict[str, Any]:
    return {
        "payment_key": data_key(row, (0, 1, 2, 3)),
        "payment_date": parse_date_text(row[1]) if len(row) > 1 else None,
        "customer_name": norm(row[2]) if len(row) > 2 else "",
        "amount": money(row[4]) if len(row) > 4 else 0.0,
        "method": norm(row[6]) if len(row) > 6 else "",
        "status": norm(row[5]) if len(row) > 5 else "",
        "source_page": "wallet.php",
        "raw_payload": row,
    }


def cod_record(row: list[list[Any]]) -> dict[str, Any]:
    return {
        "order_id": norm(row[0]) if len(row) > 0 else "",
        "tracking_number": norm(row[1]) if len(row) > 1 else "",
        "merchant_name": norm(row[2]) if len(row) > 2 else "",
        "customer_name": norm(row[3]) if len(row) > 3 else "",
        "cod_amount": money(row[4]) if len(row) > 4 else 0.0,
        "collection_date": parse_date_text(row[5]) if len(row) > 5 else None,
        "transfer_date": parse_date_text(row[6]) if len(row) > 6 else None,
        "settlement_status": norm(row[7]) if len(row) > 7 else "",
        "shipment_status": norm(row[8]) if len(row) > 8 else "",
        "raw_payload": row,
    }


def extract_shipments(html: str) -> tuple[list[list[Any]], int]:
    tables = parse_tables(html)
    if not tables:
        return [], 0
    table = max(tables, key=len)
    rows = [normalize_headers(r) for r in table]
    return rows, max(0, len(rows) - 1)


def extract_wallet(html: str) -> tuple[list[list[Any]], int, int]:
    tables = parse_tables(html)
    tx_rows: list[list[Any]] = []
    cod_rows: list[list[Any]] = []
    for table in tables:
        headers = normalize_headers(table[0]) if table else []
        joined = " ".join(headers)
        if "الرقم" in joined and "الوصف" in joined:
            tx_rows = [normalize_headers(r) for r in table]
        if "تاريخ التحصيل" in joined and "تاريخ التحويل" in joined:
            cod_rows = [normalize_headers(r) for r in table]
    return tx_rows, max(0, len(tx_rows) - 1), max(0, len(cod_rows) - 1)


def extract_cod(html: str) -> list[list[Any]]:
    tables = parse_tables(html)
    for table in tables:
        headers = normalize_headers(table[0]) if table else []
        if "تاريخ التحصيل" in " ".join(headers) and "تاريخ التحويل" in " ".join(headers):
            return [normalize_headers(r) for r in table]
    return []


def normalize_row(values: list[Any]) -> list[Any]:
    return [None if v == "" else v for v in values]


def dedupe_rows(rows: list[list[Any]], key_indexes: tuple[int, ...]) -> tuple[list[list[Any]], int]:
    seen: dict[str, list[Any]] = {}
    duplicates = 0
    for row in rows:
        key = data_key(row, key_indexes)
        if not key.strip("| "):
            continue
        if key in seen:
            duplicates += 1
            continue
        seen[key] = normalize_row(row)
    return list(seen.values()), duplicates


def merge_raw_sheet(ws, incoming: list[list[Any]], key_indexes: tuple[int, ...]) -> tuple[int, int, int]:
    incoming_unique, dup_incoming = dedupe_rows(incoming[1:] if incoming else [], key_indexes)
    existing: dict[str, int] = {}
    for r in range(2, ws.max_row + 1):
        row = [ws.cell(r, c).value for c in range(1, max(ws.max_column, len(incoming[0]) if incoming else 0) + 1)]
        key = data_key(row, key_indexes)
        if key.strip("| "):
            existing[key] = r
    added = updated = 0
    for row in incoming_unique:
        key = data_key(row, key_indexes)
        if key in existing:
            target = existing[key]
            for c, value in enumerate(row, start=1):
                ws.cell(target, c).value = value
            updated += 1
        else:
            ws.append(row)
            added += 1
    return added, updated, dup_incoming


def ensure_raw_sheets(wb):
    for name in ("Raw_Shipments", "Raw_Wallet", "Raw_Payments", "Raw_COD"):
        find_sheet(wb, [name])


def main() -> int:
    env = load_env(ENV_PATH)
    for key in ("LEAD_USERNAME", "LEAD_PASSWORD", "LEAD_BASE_URL"):
        if not env.get(key):
            print(f"Missing {key} in .env", file=sys.stderr)
            return 2
    payload = scrape_site(env)
    source_mode = desired_source_mode()
    source_xlsx = None
    source_type = "website"
    if source_mode == "excel":
        source_xlsx = REPORT_XLSX if REPORT_XLSX.exists() else None
    if source_mode == "excel" and source_xlsx is not None and source_xlsx.exists():
        source_type = "excel"
        backup = make_backup(source_xlsx)
        wb = load_workbook(source_xlsx)
    else:
        wb = create_workbook_with_raw_sheets()
        source_xlsx = REPORT_XLSX
        backup = Path("/private/tmp/lead6-no-backup.xlsx")

    raw_ship = find_sheet(wb, ["Raw_Shipments"])
    raw_wallet = find_sheet(wb, ["Raw_Wallet"])
    raw_payments = find_sheet(wb, ["Raw_Payments"])
    raw_cod = find_sheet(wb, ["Raw_COD"])

    ship_rows, shipments_total = extract_shipments(payload["html"].get("shipments.php", ""))
    wallet_rows, wallet_total, payments_total = extract_wallet(payload["html"].get("wallet.php", ""))
    cod_rows = extract_cod(payload["html"].get("collect-cod.php", ""))

    if source_type == "website":
        wb = create_source_workbook_from_website(ship_rows, wallet_rows)
        raw_ship = find_sheet(wb, ["Raw_Shipments"])
        raw_wallet = find_sheet(wb, ["Raw_Wallet"])
        raw_payments = find_sheet(wb, ["Raw_Payments"])
        raw_cod = find_sheet(wb, ["Raw_COD"])
    else:
        raw_ship = find_sheet(wb, ["Raw_Shipments"])
        raw_wallet = find_sheet(wb, ["Raw_Wallet"])
        raw_payments = find_sheet(wb, ["Raw_Payments"])
        raw_cod = find_sheet(wb, ["Raw_COD"])
    shipments_added = shipments_updated = shipments_dups = 0
    wallet_added = wallet_updated = wallet_dups = 0
    payments_added = payments_updated = payments_dups = 0
    cod_added = cod_updated = cod_dups = 0

    if ship_rows:
        ship_rows = [ship_rows[0]] + [row for row in ship_rows[1:] if row_date_at_or_after(row, (13,))]
        shipments_added, shipments_updated, shipments_dups = merge_raw_sheet(raw_ship, ship_rows, (1,))
    if wallet_rows:
        wallet_rows = [wallet_rows[0]] + [row for row in wallet_rows[1:] if row_date_at_or_after(row, (1,))]
        wallet_added, wallet_updated, wallet_dups = merge_raw_sheet(raw_wallet, wallet_rows, (0, 1, 2, 3))
        payments_added, payments_updated, payments_dups = merge_raw_sheet(raw_payments, wallet_rows, (0, 1, 2, 3))
    if cod_rows:
        cod_rows = [cod_rows[0]] + [row for row in cod_rows[1:] if row_date_at_or_after(row, (3, 4))]
        cod_added, cod_updated, cod_dups = merge_raw_sheet(raw_cod, cod_rows, (0,))

    wb.save(source_xlsx)
    canonical_source = PROJECT_DIR / "source" / "lead6.xlsx"
    if source_type == "excel" and source_xlsx != canonical_source:
        canonical_source.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source_xlsx, canonical_source)

    db_report = {"enabled": False, "synced": False, "comparison": {}}
    if db_enabled():
        try:
            with get_conn() as conn:
                ensure_schema(conn)
                run_id = make_sync_run(conn, "sync_from_lead.py")
                shipments_payload = [shipment_record(row) for row in ship_rows[1:]]
                wallet_payload = [wallet_record(row, "wallet.php", "transaction") for row in wallet_rows[1:]]
                payments_payload = [payment_record(row) for row in wallet_rows[1:]]
                cod_payload = [cod_record(row) for row in cod_rows[1:]]
                shipment_ins, shipment_upd = upsert_rows(conn, "shipments", shipments_payload, ["order_id"], [
                    "tracking_number", "merchant_name", "store_name", "customer_name", "city", "carrier", "payment_type",
                    "status", "shipment_date", "delivery_date", "weight", "cod_amount", "shipping_charge",
                    "customer_price_gross", "customer_price_net", "platform_cost_gross", "platform_cost_net",
                    "base_profit", "extra_kg", "extra_profit", "cod_profit", "total_profit", "included_in_profit",
                    "source_row", "source_hash", "raw_payload",
                ])
                wallet_ins, wallet_upd = upsert_rows(conn, "wallet_transactions", wallet_payload, ["transaction_key"], [
                    "transaction_date", "user_name", "description", "amount", "transaction_type", "balance_before",
                    "balance_after", "source_page", "raw_payload",
                ])
                payment_ins, payment_upd = upsert_rows(conn, "payments", payments_payload, ["payment_key"], [
                    "payment_date", "customer_name", "amount", "method", "status", "source_page", "raw_payload",
                ])
                cod_ins, cod_upd = upsert_rows(conn, "cod_collections", cod_payload, ["order_id"], [
                    "tracking_number", "merchant_name", "customer_name", "cod_amount", "collection_date",
                    "transfer_date", "settlement_status", "shipment_status", "raw_payload",
                ])
                comparison = compare_counts(conn)
                finish_sync_run(
                    conn,
                    run_id,
                    "ok",
                    inserted=shipment_ins + wallet_ins + payment_ins + cod_ins,
                    updated=shipment_upd + wallet_upd + payment_upd + cod_upd,
                    skipped=shipments_dups + wallet_dups + payments_dups + cod_dups,
                )
                db_report = {
                    "enabled": True,
                    "synced": True,
                    "comparison": comparison,
                    "db_rows_inserted": shipment_ins + wallet_ins + payment_ins + cod_ins,
                    "db_rows_updated": shipment_upd + wallet_upd + payment_upd + cod_upd,
                }
        except Exception as exc:
            db_report = {"enabled": True, "synced": False, "error": str(exc)}

    build_report_error = None
    site_error = None
    build_completed = None
    site_completed = None
    try:
        env = os.environ.copy()
        env["PYTHONPATH"] = str(PROJECT_DIR) + (os.pathsep + env["PYTHONPATH"] if env.get("PYTHONPATH") else "")
        build_completed = subprocess.run(
            [sys.executable, str(PROJECT_DIR / "scripts" / "build_report.py")],
            cwd=str(PROJECT_DIR),
            check=False,
            capture_output=True,
            text=True,
            env=env,
        )
        if build_completed.returncode != 0:
            build_report_error = {
                "type": "build_report_failed",
                "exit_code": build_completed.returncode,
                "stdout": build_completed.stdout[-5000:],
                "stderr": build_completed.stderr[-5000:],
            }
    except Exception as exc:
        build_report_error = {
            "type": type(exc).__name__,
            "traceback": traceback.format_exc()[-5000:],
        }

    try:
        env = os.environ.copy()
        env["PYTHONPATH"] = str(PROJECT_DIR) + (os.pathsep + env["PYTHONPATH"] if env.get("PYTHONPATH") else "")
        site_completed = subprocess.run(
            [sys.executable, str(PROJECT_DIR / "web" / "generate_site.py")],
            cwd=str(PROJECT_DIR),
            check=False,
            capture_output=True,
            text=True,
            env=env,
        )
        if site_completed.returncode != 0:
            site_error = {
                "type": "generate_site_failed",
                "exit_code": site_completed.returncode,
                "stdout": site_completed.stdout[-5000:],
                "stderr": site_completed.stderr[-5000:],
            }
    except Exception as exc:
        site_error = {
            "type": type(exc).__name__,
            "traceback": traceback.format_exc()[-5000:],
        }

    report = {
        "backup": str(backup),
        "source_workbook": str(source_xlsx),
        "source_type": source_type,
        "output_excel": str(REPORT_XLSX),
        "shipments_rows": max(0, len(ship_rows) - 1),
        "shipments_new": shipments_added,
        "shipments_updated": shipments_updated,
        "shipments_duplicates": shipments_dups,
        "wallet_rows": wallet_total,
        "wallet_new": wallet_added,
        "wallet_updated": wallet_updated,
        "wallet_duplicates": wallet_dups,
        "payments_rows": payments_total,
        "payments_new": payments_added,
        "payments_updated": payments_updated,
        "payments_duplicates": payments_dups,
        "cod_rows": max(0, len(cod_rows) - 1),
        "cod_new": cod_added,
        "cod_updated": cod_updated,
        "cod_duplicates": cod_dups,
        "errors": 0,
        "api_observations": len(payload.get("logs", [])),
        "checks": payload.get("checks", {}),
        "login_status": payload.get("checks", {}).get("login_status"),
        "login_final_url": payload.get("checks", {}).get("login_final_url"),
        "login_http_status": payload.get("checks", {}).get("login_http_status"),
        "login_redirected_to_login": payload.get("checks", {}).get("login_redirected_to_login"),
        "login_html_500": payload.get("checks", {}).get("login_html_500"),
        "postgres": db_report,
        "dashboard_refreshed": True,
        "build_report_error": build_report_error,
        "site_error": site_error,
        "build_report_exit_code": build_completed.returncode if build_completed is not None else None,
        "site_exit_code": site_completed.returncode if site_completed is not None else None,
    }

    state = load_state()
    state["last_run_at"] = dt.datetime.now().isoformat(timespec="seconds")
    state["backup"] = str(backup)
    state["network_events"] = len(payload.get("logs", []))
    save_state(state)
    print(f"[lead-sync] source={source_type}", flush=True)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
