"""Log append-only di ogni passaggio. Una riga JSON per evento."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path


def write_log(path: Path, step: str, status: str, **fields: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    record = {
        "ts": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "step": step,
        "status": status,
    }
    record.update(fields)
    with path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")
