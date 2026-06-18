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
from html.parser import HTMLParser
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


PROJECT_DIR = Path(__file__).resolve().parents[1]
SOURCE_XLSX = PROJECT_DIR / "source" / "lead6.xlsx"
REPORT_XLSX = PROJECT_DIR / "output" / "lead6_report.xlsx"
BACKUP_DIR = PROJECT_DIR / "backups"
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
    BACKUP_DIR.mkdir(parents=True, exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    dst = BACKUP_DIR / f"{src.stem}-{stamp}{src.suffix}"
    shutil.copy2(src, dst)
    return dst


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
            self._form = {"method": attrs.get("method", "get").upper(), "action": attrs.get("action", ""), "inputs": []}
        elif self._in_form and tag == "input":
            self._form["inputs"].append(attrs)

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


def http_get(opener, url: str) -> str:
    with opener.open(url) as resp:
        return resp.read().decode("utf-8", errors="ignore")


def http_post(opener, url: str, data: dict[str, str]) -> str:
    body = urllib.parse.urlencode(data).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers={"Content-Type": "application/x-www-form-urlencoded"})
    with opener.open(req) as resp:
        return resp.read().decode("utf-8", errors="ignore")


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
        checks = {
            "dashboard_links": any(token in html for token in ("لوحة التحكم", "/admin/shipments.php", "/admin/wallet.php", "/admin/collect-cod.php")),
            "shipments_link": any("/admin/shipments.php" in link for link in parser.links),
            "wallet_link": any("/admin/wallet.php" in link for link in parser.links),
            "cod_link": any("/admin/collect-cod.php" in link for link in parser.links),
        }
        return all(checks.values()), checks

    if auth_state.get("cookie"):
        opener.addheaders = [("Cookie", auth_state["cookie"])]
        cached_dashboard = http_get(opener, f"{base}/admin/dashboard.php")
        ok, cached_checks = page_looks_logged_in(cached_dashboard)
        logs.append({"kind": "probe", "page": "dashboard.php", "cached_session": ok, "checks": cached_checks})
        shipments_probe = http_get(opener, f"{base}/admin/shipments.php")
        shipment_ok, shipment_checks = page_looks_logged_in(shipments_probe)
        logs.append({"kind": "probe", "page": "shipments.php", "cached_session": shipment_ok, "checks": shipment_checks})
        wallet_probe = http_get(opener, f"{base}/admin/wallet.php")
        wallet_ok, wallet_checks = page_looks_logged_in(wallet_probe)
        logs.append({"kind": "probe", "page": "wallet.php", "cached_session": wallet_ok, "checks": wallet_checks})
        cod_probe = http_get(opener, f"{base}/admin/collect-cod.php")
        cod_ok, cod_checks = page_looks_logged_in(cod_probe)
        logs.append({"kind": "probe", "page": "collect-cod.php", "cached_session": cod_ok, "checks": cod_checks})
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

    login_html = http_get(opener, f"{base}/login.php")
    logs.append({"kind": "request", "url": f"{base}/login.php"})
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
    action = form.get("action") or "/login.php"
    if not action.startswith("http"):
        action = f"{base}/{action.lstrip('/')}"
    login_result = http_post(opener, action, post_data)
    logs.append({"kind": "request", "url": action})

    dashboard_html = http_get(opener, f"{base}/admin/dashboard.php")
    logs.append({"kind": "request", "url": f"{base}/admin/dashboard.php"})
    has_dashboard_links, dashboard_checks = page_looks_logged_in(dashboard_html)
    shipments_probe = http_get(opener, f"{base}/admin/shipments.php")
    logs.append({"kind": "request", "url": f"{base}/admin/shipments.php"})
    shipments_ok, shipments_checks = page_looks_logged_in(shipments_probe)
    wallet_probe = http_get(opener, f"{base}/admin/wallet.php")
    logs.append({"kind": "request", "url": f"{base}/admin/wallet.php"})
    wallet_ok, wallet_checks = page_looks_logged_in(wallet_probe)
    cod_probe = http_get(opener, f"{base}/admin/collect-cod.php")
    logs.append({"kind": "request", "url": f"{base}/admin/collect-cod.php"})
    cod_ok, cod_checks = page_looks_logged_in(cod_probe)
    login_issue = None
    if not has_dashboard_links:
        login_issue = "dashboard markers not found"
    elif not shipments_ok:
        login_issue = "shipments page was not accessible"
    elif not wallet_ok:
        login_issue = "wallet page was not accessible"
    elif not cod_ok:
        login_issue = "cod page was not accessible"
    if login_issue:
        raise RuntimeError(f"Login/session check failed: {login_issue}")

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
        },
    }


def normalize_headers(row: list[str]) -> list[str]:
    return [norm(v) for v in row]


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
    if not SOURCE_XLSX.exists():
        print(f"Missing workbook: {SOURCE_XLSX}", file=sys.stderr)
        return 2

    backup = make_backup(SOURCE_XLSX)
    payload = scrape_site(env)

    wb = load_workbook(SOURCE_XLSX)
    raw_ship = find_sheet(wb, ["Raw_Shipments"])
    raw_wallet = find_sheet(wb, ["Raw_Wallet"])
    raw_payments = find_sheet(wb, ["Raw_Payments"])
    raw_cod = find_sheet(wb, ["Raw_COD"])

    ship_rows, shipments_total = extract_shipments(payload["html"].get("shipments.php", ""))
    wallet_rows, wallet_total, payments_total = extract_wallet(payload["html"].get("wallet.php", ""))
    cod_rows = extract_cod(payload["html"].get("collect-cod.php", ""))

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

    wb.save(SOURCE_XLSX)

    subprocess.run([sys.executable, str(PROJECT_DIR / "scripts" / "build_report.py")], cwd=str(PROJECT_DIR), check=True)
    subprocess.run([sys.executable, str(PROJECT_DIR / "web" / "generate_site.py")], cwd=str(PROJECT_DIR), check=True)

    report = {
        "backup": str(backup),
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
        "dashboard_refreshed": True,
    }

    state = load_state()
    state["last_run_at"] = dt.datetime.now().isoformat(timespec="seconds")
    state["backup"] = str(backup)
    state["network_events"] = len(payload.get("logs", []))
    save_state(state)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
