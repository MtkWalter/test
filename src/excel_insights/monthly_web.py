"""Loopback-only browser app. All input, computation and temporary files stay local."""
from __future__ import annotations

import argparse
import json
import re
import secrets
import shutil
import sys
import tempfile
import threading
import time
import uuid
import webbrowser
from concurrent.futures import ThreadPoolExecutor
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from importlib.resources import files
from pathlib import Path
from urllib.parse import parse_qs, quote, urlsplit
from urllib.request import ProxyHandler, Request, build_opener

from openpyxl import load_workbook

from .monthly_report import ReportRules, fiscal_label, generate_report, inspect_workbook

JSON_LIMIT = 512 * 1024
PREVIEW_ROWS = 200
STATIC = {"": "index.html", "app.js": "app.js", "style.css": "style.css"}


def _display(value, *, percent=False):
    if value is None:
        return ""
    amount = Decimal(str(value)) * (100 if percent else 1)
    return f"{amount.quantize(Decimal(1), rounding=ROUND_HALF_UP):,}" + ("%" if percent else "")


def preview_workbook(path: Path) -> dict:
    book = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = book["monthly pivot"]
        rows = []
        truncated = False
        for row in sheet.iter_rows(min_row=4, max_col=12):
            node = row[0].value
            if not node or str(node).startswith("金額："):
                continue
            if len(rows) >= PREVIEW_ROWS:
                truncated = True
                break
            cells = [str(node)]
            cells.extend(_display(c.value, percent=i in (8, 10)) for i, c in enumerate(row[1:11], 1))
            cells.append(str(row[11].value or ""))
            colors = []
            for i in (8, 10):
                color = row[i].fill.fgColor.rgb
                colors.append("high" if color in ("00FFF0AF", "FFFFF0AF") else
                              "low" if color in ("00C9EFF3", "FFC9EFF3") else "")
            rows.append({"cells": cells, "tsmc": colors[0], "non_t": colors[1]})
        return {"rows": rows, "truncated": truncated}
    finally:
        book.close()


@dataclass
class Upload:
    path: Path
    name: str
    sheets: list[str]


@dataclass
class Job:
    id: str
    filename: str
    status: str = "running"
    progress: str = "正在讀取來源資料…"
    result: dict | None = None
    preview: dict | None = None
    error: str | None = None
    path: Path | None = None


class LocalReportServer(ThreadingHTTPServer):
    allow_reuse_address = False
    daemon_threads = True

    def __init__(self, port: int = 0):
        self.token = secrets.token_urlsafe(32)
        self.prefix = f"/s/{self.token}/"
        self.workspace = tempfile.TemporaryDirectory(prefix="MonthlyReport-")
        self.root = Path(self.workspace.name)
        self.uploads: dict[str, Upload] = {}
        self.jobs: dict[str, Job] = {}
        self.lock = threading.RLock()
        self.executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="monthly-report")
        self.ready = threading.Event()
        self.closed = False
        self.transfers = 0
        try:
            super().__init__(("127.0.0.1", port), ReportHandler)
        except Exception:
            self.executor.shutdown(wait=False)
            self.workspace.cleanup()
            raise
        self.origin = f"http://127.0.0.1:{self.server_port}"
        self.url = self.origin + self.prefix

    def running(self) -> bool:
        return any(job.status == "running" for job in self.jobs.values())

    @contextmanager
    def transfer(self):
        with self.lock:
            self.transfers += 1
        try:
            yield
        finally:
            with self.lock:
                self.transfers -= 1

    def create_job(self, payload: dict) -> Job:
        required = ("upload_id", "sheet", "year", "month")
        if not isinstance(payload, dict) or any(name not in payload for name in required):
            raise ValueError("請先選擇 INPUT、工作表與報表年月。")
        with self.lock:
            if self.running():
                raise BlockingIOError("已有報表正在處理，請等待完成。")
            upload = self.uploads.get(payload["upload_id"])
            if upload is None:
                raise ValueError("找不到 INPUT，請重新選擇 Excel 檔案。")
            if payload["sheet"] not in upload.sheets:
                raise ValueError("請選擇有效的來源工作表。")
            for name in ("year", "month"):
                if isinstance(payload[name], bool) or not isinstance(payload[name], int):
                    raise ValueError("年份與月份必須是整數。")
            if not 1900 <= payload["year"] <= 9999 or not 1 <= payload["month"] <= 12:
                raise ValueError("年份或月份超出範圍。")
            rules = ReportRules.from_dict(payload.get("rules", {}))
            fy = payload.get("fiscal_year") or fiscal_label(payload["year"], payload["month"], rules.fiscal_start_month)
            if not isinstance(fy, str) or not re.fullmatch(r"FY\d{2}", fy.strip(), re.I):
                raise ValueError("年度標籤請使用 FY27 這類格式。")
            default_name = f"monthly-output-{payload['year']}{payload['month']:02d}.xlsx"
            filename = str(payload.get("filename") or default_name)
            filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename).strip(" .")
            if not filename.lower().endswith(".xlsx"):
                filename += ".xlsx"
            filename = filename[:115] + ".xlsx" if len(filename) > 120 else filename
            job_id = uuid.uuid4().hex
            job = Job(job_id, filename, path=self.root / f"result-{job_id}.xlsx")
            self.jobs[job_id] = job
            self.executor.submit(self._generate, job, upload, payload, fy, rules)
            return job

    def _generate(self, job, upload, payload, fiscal_year, rules):
        def progress(message):
            with self.lock:
                job.progress = message
        try:
            result = generate_report(upload.path, job.path, payload["year"], payload["month"],
                                     sheet=payload["sheet"], fiscal_year=fiscal_year,
                                     rules=rules, progress=progress)
            preview = preview_workbook(job.path)
            with self.lock:
                job.result = {k: v for k, v in result.items() if k != "output"}
                job.preview = preview
                job.progress = "OUTPUT Excel 已完成。"
                job.status = "success"
        except Exception as exc:
            with self.lock:
                job.error = str(exc)
                job.progress = "請依錯誤訊息修正來源或設定。"
                job.status = "error"

    def close(self):
        if self.closed:
            return
        self.closed = True
        self.executor.shutdown(wait=True, cancel_futures=True)
        self.server_close()
        self.workspace.cleanup()


class ReportHandler(BaseHTTPRequestHandler):
    server: LocalReportServer

    def log_message(self, *_args):
        pass  # Windowed EXE has no stderr; do not log filenames or session URLs.

    def _route(self) -> str | None:
        if self.headers.get("Host") != f"127.0.0.1:{self.server.server_port}":
            self._json({"error": "無效的本機連線。"}, HTTPStatus.FORBIDDEN)
            return None
        origin = self.headers.get("Origin")
        if origin and origin != self.server.origin:
            self._json({"error": "僅允許本機頁面操作。"}, HTTPStatus.FORBIDDEN)
            return None
        if not urlsplit(self.path).path.startswith(self.server.prefix):
            self._json({"error": "本機連線已失效，請重新啟動程式。"}, HTTPStatus.NOT_FOUND)
            return None
        return urlsplit(self.path).path[len(self.server.prefix):]

    def _headers(self, status, content_type, length, *, filename=None):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(length))
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header("Content-Security-Policy", "default-src 'self'; script-src 'self'; style-src 'self'; img-src 'self' data:; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'")
        if filename:
            self.send_header("Content-Disposition", "attachment; filename*=UTF-8''" + quote(filename))
        self.end_headers()

    def _json(self, value, status=HTTPStatus.OK):
        body = json.dumps(value, ensure_ascii=False).encode("utf-8")
        self._headers(status, "application/json; charset=utf-8", len(body))
        self.wfile.write(body)

    def do_GET(self):
        try:
            route = self._route()
            if route is None:
                return
            if route in STATIC:
                name = STATIC[route]
                body = files("excel_insights").joinpath("web", name).read_bytes()
                mime = {"index.html": "text/html", "app.js": "text/javascript", "style.css": "text/css"}[name]
                self._headers(HTTPStatus.OK, mime + "; charset=utf-8", len(body))
                self.wfile.write(body)
            elif route == "api/state":
                now = datetime.now()
                self.server.ready.set()
                self._json({"year": now.year, "month": now.month, "rules": ReportRules().to_dict(), "interface": "browser"})
            elif route.startswith("api/jobs/"):
                job_id = route.removeprefix("api/jobs/")
                with self.server.lock:
                    job = self.server.jobs.get(job_id)
                    if job is None:
                        self._json({"error": "找不到報表工作。"}, HTTPStatus.NOT_FOUND)
                    else:
                        self._json({"id": job.id, "status": job.status, "progress": job.progress,
                                    "result": job.result, "preview": job.preview, "error": job.error,
                                    "filename": job.filename, "download": f"api/download/{job.id}" if job.status == "success" else None})
            elif route.startswith("api/download/"):
                job_id = route.removeprefix("api/download/")
                with self.server.lock:
                    job = self.server.jobs.get(job_id)
                    if not job or job.status != "success" or not job.path.is_file():
                        self._json({"error": "報表尚未完成。"}, HTTPStatus.NOT_FOUND)
                        return
                    path, filename = job.path, job.filename
                with self.server.transfer():
                    self._headers(HTTPStatus.OK, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", path.stat().st_size, filename=filename)
                    with path.open("rb") as stream:
                        shutil.copyfileobj(stream, self.wfile, length=1024 * 1024)
            else:
                self._json({"error": "找不到這個頁面。"}, HTTPStatus.NOT_FOUND)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as exc:
            self._json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def _body(self, maximum=None) -> bytes:
        length = int(self.headers.get("Content-Length", "-1"))
        if length < 0:
            raise ValueError("請使用瀏覽器重新選擇檔案。")
        if maximum is not None and length > maximum:
            raise ValueError("分類設定內容過大。")
        body = self.rfile.read(length)
        if len(body) != length:
            raise ValueError("資料傳輸未完成，請重試。")
        return body

    def do_POST(self):
        try:
            route = self._route()
            if route is None:
                return
            if route == "api/upload":
                self._upload()
            elif route == "api/jobs":
                if self.headers.get_content_type() != "application/json":
                    raise ValueError("報表設定格式不正確。")
                payload = json.loads(self._body(JSON_LIMIT))
                job = self.server.create_job(payload)
                self._json({"id": job.id}, HTTPStatus.ACCEPTED)
            elif route == "api/shutdown":
                with self.server.lock:
                    if self.server.running() or self.server.transfers:
                        self._json({"error": "正在處理或傳送檔案，請等待完成後關閉。"}, HTTPStatus.CONFLICT)
                        return
                self._body(JSON_LIMIT)
                self._json({"stopped": True})
                threading.Thread(target=self.server.shutdown, daemon=True).start()
            else:
                self._json({"error": "找不到這項操作。"}, HTTPStatus.NOT_FOUND)
        except BlockingIOError as exc:
            self._json({"error": str(exc)}, HTTPStatus.CONFLICT)
        except (ValueError, TypeError, KeyError, json.JSONDecodeError) as exc:
            self._json({"error": str(exc)}, HTTPStatus.BAD_REQUEST)
        except (BrokenPipeError, ConnectionResetError):
            pass
        except Exception as exc:
            self._json({"error": str(exc)}, HTTPStatus.INTERNAL_SERVER_ERROR)

    def _upload(self):
        with self.server.transfer():
            self._receive_upload()

    def _receive_upload(self):
        query = parse_qs(urlsplit(self.path).query)
        filename = query.get("filename", [""])[0].replace("\\", "/").rsplit("/", 1)[-1]
        suffix = Path(filename).suffix.lower()
        if suffix not in {".xlsx", ".xlsm"}:
            raise ValueError("請選擇 .xlsx 或 .xlsm Excel 檔案。")
        length = int(self.headers.get("Content-Length", "-1"))
        if length <= 0:
            raise ValueError("Excel 檔案是空的或尚未選擇。")
        upload_id = uuid.uuid4().hex
        folder = self.server.root / upload_id
        folder.mkdir()
        safe_name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", filename)
        safe_name = "input-" + safe_name[:100 - len(suffix)] + suffix if len(safe_name) > 100 else "input-" + safe_name
        path = folder / safe_name
        try:
            remaining = length
            with path.open("wb") as stream:
                while remaining:
                    block = self.rfile.read(min(1024 * 1024, remaining))
                    if not block:
                        raise ValueError("Excel 傳輸未完成，請重新選擇檔案。")
                    stream.write(block)
                    remaining -= len(block)
            try:
                layouts = inspect_workbook(path)
            except Exception as exc:
                raise ValueError(f"無法辨識來源 Excel：{exc}") from exc
            with self.server.lock:
                self.server.uploads[upload_id] = Upload(path, filename, [s.name for s in layouts])
            self._json({"id": upload_id, "name": filename, "bytes": length, "sheets": [s.name for s in layouts]}, HTTPStatus.CREATED)
        except Exception:
            shutil.rmtree(folder, ignore_errors=True)
            raise


def smoke_browser_app(source: Path, output: Path, summary: Path, year: int, month: int) -> int:
    """Exercise the frozen browser server through the same HTTP routes as the UI."""
    server = LocalReportServer()
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    opener = build_opener(ProxyHandler({}))  # Only this fixed loopback server.
    def request(path, data=None, content_type="application/json"):
        req = Request(server.url + path, data=data, headers={"Content-Type": content_type})
        with opener.open(req, timeout=120) as response:
            return response.read()
    try:
        state = json.loads(request("api/state"))
        if b"app.js" not in request("") or b"initialize" not in request("app.js") or b".sidebar" not in request("style.css"):
            raise ValueError("瀏覽器介面檔案未完整打包。")
        upload = json.loads(request("api/upload?filename=" + quote(source.name), source.read_bytes(), "application/octet-stream"))
        job = json.loads(request("api/jobs", json.dumps({"upload_id": upload["id"], "sheet": upload["sheets"][0],
                                                        "year": year, "month": month}).encode()))
        deadline = time.monotonic() + 100
        while time.monotonic() < deadline:
            status = json.loads(request("api/jobs/" + job["id"]))
            if status["status"] != "running":
                break
            time.sleep(.05)
        else:
            raise TimeoutError("瀏覽器報表服務處理逾時。")
        if status["status"] != "success":
            raise ValueError(status["error"])
        output.write_bytes(request(status["download"]))
        result = dict(status["result"], output=str(output), interface=state["interface"])
        summary.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
        request("api/shutdown", b"{}")
        return 0
    except Exception as exc:
        summary.write_text(json.dumps({"error": str(exc)}, ensure_ascii=False), encoding="utf-8")
        server.shutdown()
        return 1
    finally:
        thread.join(timeout=5)
        server.close()


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Monthly Report 本機瀏覽器報表產生器")
    parser.add_argument("--port", type=int, default=0, help="預設自動選擇可用的本機連接埠")
    parser.add_argument("--no-browser", action="store_true", help="啟動後不自動開啟瀏覽器")
    parser.add_argument("--smoke-test", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--smoke-output", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--output", type=Path, help=argparse.SUPPRESS)
    parser.add_argument("--year", type=int, default=2026, help=argparse.SUPPRESS)
    parser.add_argument("--month", type=int, default=9, help=argparse.SUPPRESS)
    args = parser.parse_args(argv)
    if args.smoke_test:
        if not args.smoke_output or not args.output:
            parser.error("--smoke-test requires --smoke-output and --output")
        return smoke_browser_app(args.smoke_test, args.output, args.smoke_output, args.year, args.month)
    if not 0 <= args.port <= 65535:
        parser.error("port must be between 0 and 65535")
    server = LocalReportServer(args.port)
    if sys.stdout:
        print(f"本機瀏覽器介面：{server.url}", flush=True)
    if not args.no_browser:
        def launch():
            if not webbrowser.open(server.url) and sys.platform == "win32":
                import ctypes
                ctypes.windll.user32.MessageBoxW(None, f"請在瀏覽器開啟：\n{server.url}", "Monthly Report", 64)
        threading.Timer(.25, launch).start()
    try:
        server.serve_forever(poll_interval=.2)
    except KeyboardInterrupt:
        pass
    finally:
        server.close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
