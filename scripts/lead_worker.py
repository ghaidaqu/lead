from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
SYNC_SCRIPT = PROJECT_DIR / "scripts" / "sync_from_lead.py"
DEFAULT_INTERVAL_SECONDS = max(300, int(os.environ.get("LEAD_WORKER_INTERVAL_SECONDS", "3600")))


def _run_once() -> int:
    env = os.environ.copy()
    env.setdefault("PYTHONUNBUFFERED", "1")
    completed = subprocess.run(
        [sys.executable, str(SYNC_SCRIPT)],
        cwd=str(PROJECT_DIR),
        env=env,
        check=False,
    )
    return int(completed.returncode)


def _sleep_until_next_run(interval_seconds: int) -> None:
    now = time.time()
    next_run = ((int(now) // interval_seconds) + 1) * interval_seconds
    time.sleep(max(1, next_run - now))


def main() -> int:
    interval_seconds = DEFAULT_INTERVAL_SECONDS
    print(f"[lead-worker] interval_seconds={interval_seconds}", flush=True)
    while True:
        started = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[lead-worker] sync_start={started}", flush=True)
        code = _run_once()
        finished = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[lead-worker] sync_end={finished} exit_code={code}", flush=True)
        _sleep_until_next_run(interval_seconds)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
