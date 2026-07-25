#!/usr/bin/env python3
"""Check Yukari HTTP ports and write a JSON file for GitHub Pages."""

from __future__ import annotations

import concurrent.futures
import json
import socket
import time
from datetime import datetime, timezone
from pathlib import Path

HOST = "ykr.moe"
PORT_START = 11000
PORT_END = 11130
TIMEOUT_SECONDS = 3.5
MAX_WORKERS = 32
OUTPUT = Path("yukari-status.json")


def check_port(port: int) -> dict[str, object]:
    started = time.perf_counter()
    try:
        # A TCP handshake is sufficient to decide whether the service port is reachable.
        with socket.create_connection((HOST, port), timeout=TIMEOUT_SECONDS):
            elapsed_ms = round((time.perf_counter() - started) * 1000)
            return {"port": port, "online": True, "ms": elapsed_ms, "error": ""}
    except OSError as exc:
        elapsed_ms = round((time.perf_counter() - started) * 1000)
        return {
            "port": port,
            "online": False,
            "ms": elapsed_ms,
            "error": exc.__class__.__name__,
        }


def main() -> None:
    ports = list(range(PORT_START, PORT_END + 1))
    with concurrent.futures.ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
        results = list(pool.map(check_port, ports))

    results.sort(key=lambda item: int(item["port"]))
    payload = {
        "checked_at": datetime.now(timezone.utc).isoformat().replace("+00:00", "Z"),
        "host": HOST,
        "range": f"{PORT_START}-{PORT_END}",
        "online_count": sum(bool(item["online"]) for item in results),
        "ports": results,
    }
    OUTPUT.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


if __name__ == "__main__":
    main()
