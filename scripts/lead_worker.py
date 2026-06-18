from __future__ import annotations

import os
import subprocess
import sys
import time
import datetime as dt
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parents[1]
SYNC_SCRIPT = PROJECT_DIR / "scripts" / "sync_from_lead.py"
DEFAULT_INTERVAL_SECONDS = max(300, int(os.environ.get("LEAD_WORKER_INTERVAL_SECONDS", "3600")))
DEFAULT_START_HOUR = int(os.environ.get("LEAD_WORKER_START_HOUR", "1"))
DEFAULT_START_MINUTE = int(os.environ.get("LEAD_WORKER_START_MINUTE", "30"))


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


def _sleep_until_start_window(start_hour: int, start_minute: int) -> None:
    now = dt.datetime.now()
    next_run = now.replace(hour=start_hour, minute=start_minute, second=0, microsecond=0)
    if next_run <= now:
        next_run = next_run + dt.timedelta(days=1)
    time.sleep(max(1, (next_run - now).total_seconds()))


def _sleep_for_interval(interval_seconds: int) -> None:
    time.sleep(max(1, interval_seconds))


def main() -> int:
    interval_seconds = DEFAULT_INTERVAL_SECONDS
    start_hour = DEFAULT_START_HOUR
    start_minute = DEFAULT_START_MINUTE
    print(
        f"[lead-worker] start_time={start_hour:02d}:{start_minute:02d} interval_seconds={interval_seconds}",
        flush=True,
    )
    print("[lead-worker] waiting for first scheduled run", flush=True)
    _sleep_until_start_window(start_hour, start_minute)
    while True:
        started = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[lead-worker] sync_start={started}", flush=True)
        code = _run_once()
        finished = time.strftime("%Y-%m-%d %H:%M:%S")
        print(f"[lead-worker] sync_end={finished} exit_code={code}", flush=True)
        _sleep_for_interval(interval_seconds)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
