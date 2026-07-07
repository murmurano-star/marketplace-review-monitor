from __future__ import annotations

import json
import mimetypes
import subprocess
import sys
from datetime import date
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse

HOST = "127.0.0.1"
PORT = 8765
REPOSITORY = "murmurano-star/marketplace-review-monitor"
WORKFLOW = "manual-monitor.yml"
ROOT = Path(__file__).resolve().parent


def validate_date(value: str) -> str:
    value = value.strip()
    if not value:
        return ""
    date.fromisoformat(value)
    return value


def run_gh(arguments: list[str]) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["gh", *arguments],
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=60,
        check=False,
    )


def trigger_workflow(date_from: str, date_to: str) -> dict:
    command = [
        "workflow",
        "run",
        WORKFLOW,
        "--repo",
        REPOSITORY,
        "--ref",
        "main",
        "-f",
        f"date_from={date_from}",
        "-f",
        f"date_to={date_to}",
    ]
    result = run_gh(command)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or result.stdout.strip() or "Не удалось запустить workflow")

    latest = run_gh([
        "run",
        "list",
        "--repo",
        REPOSITORY,
        "--workflow",
        WORKFLOW,
        "--event",
        "workflow_dispatch",
        "--limit",
        "1",
        "--json",
        "databaseId,url,status,conclusion,createdAt,displayTitle",
    ])
    runs = []
    if latest.returncode == 0 and latest.stdout.strip():
        try:
            runs = json.loads(latest.stdout)
        except json.JSONDecodeError:
            runs = []
    return {
        "ok": True,
        "message": "Сбор отзывов запущен",
        "run": runs[0] if runs else None,
    }


def latest_run() -> dict:
    result = run_gh([
        "run",
        "list",
        "--repo",
        REPOSITORY,
        "--workflow",
        WORKFLOW,
        "--limit",
        "1",
        "--json",
        "databaseId,url,status,conclusion,createdAt,updatedAt,displayTitle",
    ])
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "Не удалось получить состояние запуска")
    runs = json.loads(result.stdout or "[]")
    return {"ok": True, "run": runs[0] if runs else None}


class Handler(BaseHTTPRequestHandler):
    server_version = "ReviewMonitorLauncher/1.0"

    def log_message(self, format: str, *args) -> None:
        print(f"[launcher] {self.address_string()} - {format % args}")

    def send_json(self, payload: dict, status: int = HTTPStatus.OK) -> None:
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self) -> None:
        route = urlparse(self.path).path
        if route == "/api/status":
            try:
                self.send_json(latest_run())
            except Exception as exc:
                self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_GATEWAY)
            return

        target = ROOT / ("index.html" if route in {"", "/"} else route.lstrip("/"))
        try:
            resolved = target.resolve()
            if ROOT not in resolved.parents and resolved != ROOT:
                raise FileNotFoundError
            data = resolved.read_bytes()
        except (FileNotFoundError, IsADirectoryError):
            self.send_error(HTTPStatus.NOT_FOUND)
            return

        mime = mimetypes.guess_type(resolved.name)[0] or "application/octet-stream"
        self.send_response(HTTPStatus.OK)
        self.send_header("Content-Type", f"{mime}; charset=utf-8" if mime.startswith("text/") else mime)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_POST(self) -> None:
        if urlparse(self.path).path != "/api/run":
            self.send_error(HTTPStatus.NOT_FOUND)
            return
        try:
            length = int(self.headers.get("Content-Length", "0"))
            payload = json.loads(self.rfile.read(length) or b"{}")
            date_from = validate_date(str(payload.get("date_from", "")))
            date_to = validate_date(str(payload.get("date_to", "")))
            if date_from and date_to and date_from > date_to:
                raise ValueError("Дата начала не может быть позже даты окончания")
            self.send_json(trigger_workflow(date_from, date_to), HTTPStatus.ACCEPTED)
        except (ValueError, json.JSONDecodeError) as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except FileNotFoundError:
            self.send_json({"ok": False, "error": "GitHub CLI gh не установлен"}, HTTPStatus.FAILED_DEPENDENCY)
        except Exception as exc:
            self.send_json({"ok": False, "error": str(exc)}, HTTPStatus.BAD_GATEWAY)


def main() -> None:
    try:
        server = ThreadingHTTPServer((HOST, PORT), Handler)
    except OSError as exc:
        print(f"Не удалось запустить локальное приложение: {exc}", file=sys.stderr)
        raise SystemExit(1) from exc
    print(f"HTML-приложение: http://{HOST}:{PORT}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
