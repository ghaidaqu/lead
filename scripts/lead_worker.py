from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
SYNC_SCRIPT = PROJECT_DIR / "scripts" / "sync_from_lead.py"
DEFAULT_INTERVAL_SECONDS = max(300, int(os.environ.get("LEAD_WORKER_INTERVAL_SECONDS", "10800")))
SYNC_TIMEOUT_SECONDS = max(300, int(os.environ.get("LEAD_WORKER_SYNC_TIMEOUT_SECONDS", "900")))
SYNC_ATTEMPTS = max(1, int(os.environ.get("LEAD_WORKER_SYNC_ATTEMPTS", "2")))


def _run_once() -> int:
    env = os.environ.copy()
    env.setdefault("PYTHONUNBUFFERED", "1")
    env.setdefault("LEAD_FORCE_LOGIN_EACH_RUN", "1")
    for attempt in range(1, SYNC_ATTEMPTS + 1):
        print(f"[lead-worker] sync_attempt={attempt}/{SYNC_ATTEMPTS}", flush=True)
        try:
            completed = subprocess.run(
                [sys.executable, str(SYNC_SCRIPT)],
                cwd=str(PROJECT_DIR),
                env=env,
                check=False,
                timeout=SYNC_TIMEOUT_SECONDS,
            )
            code = int(completed.returncode)
        except subprocess.TimeoutExpired:
            code = 124
            print(f"[lead-worker] sync_timeout seconds={SYNC_TIMEOUT_SECONDS}", flush=True)
        if code == 0:
            return 0
        if attempt < SYNC_ATTEMPTS:
            time.sleep(10)
    return code


def _sleep_for_interval(interval_seconds: int) -> None:
    time.sleep(max(1, interval_seconds))


def main() -> int:
    interval_seconds = DEFAULT_INTERVAL_SECONDS
    print(f"[lead-worker] run_mode=immediate interval_seconds={interval_seconds}", flush=True)
    while True:
        started = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[lead-worker] sync_start={started}", flush=True)
        code = _run_once()
        finished = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[lead-worker] sync_end={finished} exit_code={code}", flush=True)
        next_run = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime(time.time() + interval_seconds))
        print(f"[lead-worker] next_sync={next_run}", flush=True)
        _sleep_for_interval(interval_seconds)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
