from __future__ import annotations

import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from scripts.db_store import db_enabled, ensure_schema, get_conn


def main() -> int:
    if not db_enabled():
        print("DATABASE_URL is not set or psycopg is unavailable.", file=sys.stderr)
        return 2
    with get_conn() as conn:
        ensure_schema(conn)
    print("PostgreSQL schema ensured.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
