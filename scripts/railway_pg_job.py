from __future__ import annotations

import json
import sys
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parents[1]
if str(PROJECT_DIR) not in sys.path:
    sys.path.insert(0, str(PROJECT_DIR))

from scripts.migrate_pg import main as migrate_main
from scripts.validation_report import main as validation_main


def main() -> int:
    print(json.dumps({"step": "migration", "status": "starting"}, ensure_ascii=False))
    migrate_code = migrate_main()
    if migrate_code != 0:
        print(json.dumps({"step": "migration", "status": "failed", "code": migrate_code}, ensure_ascii=False))
        return migrate_code
    print(json.dumps({"step": "validation", "status": "starting"}, ensure_ascii=False))
    validation_code = validation_main()
    if validation_code != 0:
        print(json.dumps({"step": "validation", "status": "failed", "code": validation_code}, ensure_ascii=False))
        return validation_code
    print(json.dumps({"step": "complete", "status": "ok"}, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
