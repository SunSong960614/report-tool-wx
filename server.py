from __future__ import annotations

import json
import os
import sys
import tempfile
import threading
from email.parser import BytesParser
from email.policy import default
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import quote, urlparse


APP_DIR = Path(__file__).resolve().parent
ROOT = APP_DIR.parents[1]
WORK_DIR = ROOT / "work"
sys.path.insert(0, str(WORK_DIR))
sys.path.insert(0, str(APP_DIR))

from report_engine import ReportError, build_report, parse_docx  # noqa: E402


MAX_JSON_BYTES = 10 * 1024 * 1024
MAX_UPLOAD_BYTES = 60 * 1024 * 1024
MAX_FILE_BYTES = 20 * 1024 * 1024
MAX_FILES = 8
REQUEST_SLOTS = threading.BoundedSemaphore(max(1, int(os.environ.get("MAX_CONCURRENT_REQUESTS", "2"))))


class Handler(SimpleHTTPRequestHandler):
    server_version = "SchoolReportTool/1.0"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(APP_DIR / "public"), **kwargs)

    def end_headers(self):
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("X-Frame-Options", "DENY")
        self.send_header("Referrer-Policy", "no-referrer")
        self.send_header(
            "Content-Security-Policy",
            "default-src 'self'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'",
        )
        super().end_headers()

    def _content_length(self, limit: int) -> int:
        try:
            length = int(self.headers.get("Content-Length", "0"))
        except ValueError as exc:
            raise ReportError("请求长度格式不正确") from exc
        if length <= 0:
            raise ReportError("请求内容为空")
        if length > limit:
            raise ReportError("请求内容超过大小限制")
        return length

    def _json(self, status: int, payload: dict):
        body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(body)

    def _read_json(self):
        length = self._content_length(MAX_JSON_BYTES)
        return json.loads(self.rfile.read(length).decode("utf-8"))

    def _multipart_files(self):
        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            raise ReportError("上传请求格式不正确")
        length = self._content_length(MAX_UPLOAD_BYTES)
        raw = self.rfile.read(length)
        message = BytesParser(policy=default).parsebytes(
            f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n".encode() + raw
        )
        files = []
        for part in message.iter_parts():
            if part.get_content_disposition() != "form-data":
                continue
            filename = part.get_filename()
            if filename:
                content = part.get_payload(decode=True) or b""
                if len(content) > MAX_FILE_BYTES:
                    raise ReportError(f"{Path(filename).name} 超过 20MB 限制")
                files.append((Path(filename).name, content))
                if len(files) > MAX_FILES:
                    raise ReportError(f"一次最多上传 {MAX_FILES} 份报告")
        return files

    def do_GET(self):
        path = urlparse(self.path).path
        if path == "/api/health":
            self._json(200, {"status": "ok"})
            return
        if path == "/":
            self.path = "/index.html"
        super().do_GET()

    def do_POST(self):
        if not REQUEST_SLOTS.acquire(blocking=False):
            self._json(503, {"error": "当前生成任务较多，请稍后重试"})
            return
        path = urlparse(self.path).path
        try:
            if path == "/api/parse":
                files = self._multipart_files()
                if not files:
                    raise ReportError("请选择至少一份 Word 报告")
                results = []
                for filename, content in files:
                    try:
                        data = parse_docx(content, filename)
                        results.append({"filename": filename, "ok": True, "data": data})
                    except ReportError as exc:
                        results.append({"filename": filename, "ok": False, "error": str(exc)})
                self._json(200, {"files": results})
                return

            if path == "/api/generate":
                payload = self._read_json()
                school = payload.get("school", "")
                year = payload.get("year")
                datasets = payload.get("datasets", [])
                safe_school = "".join(ch for ch in str(school) if ch not in '\\/:*?"<>|').strip()
                with tempfile.TemporaryDirectory(prefix="school-report-download-") as temp:
                    filename = f"{safe_school}学生素养测评数据分析报告.docx"
                    output = Path(temp) / filename
                    build_report(school, year, datasets, output)
                    body = output.read_bytes()
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                self.send_header("Content-Disposition", f"attachment; filename*=UTF-8''{quote(filename)}")
                self.send_header("Content-Length", str(len(body)))
                self.send_header("Cache-Control", "no-store")
                self.end_headers()
                self.wfile.write(body)
                return

            self._json(404, {"error": "接口不存在"})
        except ReportError as exc:
            self._json(400, {"error": str(exc)})
        except json.JSONDecodeError:
            self._json(400, {"error": "请求数据格式不正确"})
        except Exception as exc:
            print(f"Unhandled error: {exc}", file=sys.stderr)
            self._json(500, {"error": "处理失败，请检查报告格式后重试"})
        finally:
            REQUEST_SLOTS.release()

    def log_message(self, format, *args):
        print(f"[{self.log_date_time_string()}] {format % args}")


def main():
    host = os.environ.get("HOST", "127.0.0.1")
    port = int(os.environ.get("PORT", "18966"))
    ThreadingHTTPServer.daemon_threads = True
    server = ThreadingHTTPServer((host, port), Handler)
    print(f"学校测评报告合成工具已启动：http://{host}:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass
    finally:
        server.server_close()


if __name__ == "__main__":
    main()
