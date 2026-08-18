from __future__ import annotations

import tempfile
from io import BytesIO
from pathlib import Path
from urllib.parse import quote

from flask import Flask, jsonify, redirect, request, send_file
from werkzeug.exceptions import BadRequest, RequestEntityTooLarge

MAX_REQUEST_BYTES = 4 * 1024 * 1024
MAX_FILES = 8


class RequestError(ValueError):
    pass


app = Flask(__name__, static_folder=None)
app.config["MAX_CONTENT_LENGTH"] = MAX_REQUEST_BYTES


@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"] = "DENY"
    response.headers["Referrer-Policy"] = "no-referrer"
    response.headers["Content-Security-Policy"] = (
        "default-src 'self'; base-uri 'none'; form-action 'self'; frame-ancestors 'none'"
    )
    if request.path.startswith("/api/"):
        response.headers["Cache-Control"] = "no-store"
    return response


@app.get("/")
def home():
    return redirect("/index.html", code=307)


@app.get("/api/health")
def health():
    return jsonify(status="ok")


@app.post("/api/parse")
def parse_reports():
    import report_engine

    files = request.files.getlist("files")
    if not files:
        raise RequestError("请选择至少一份 Word 报告")
    if len(files) > MAX_FILES:
        raise RequestError(f"一次最多上传 {MAX_FILES} 份报告")

    results = []
    for uploaded in files:
        filename = Path(uploaded.filename or "").name
        content = uploaded.read()
        try:
            data = report_engine.parse_docx(content, filename)
            results.append({"filename": filename, "ok": True, "data": data})
        except report_engine.ReportError as exc:
            results.append({"filename": filename, "ok": False, "error": str(exc)})
    return jsonify(files=results)


@app.post("/api/generate")
def generate_report():
    import report_engine

    payload = request.get_json(force=False, silent=False)
    school = payload.get("school", "")
    year = payload.get("year")
    datasets = payload.get("datasets", [])
    safe_school = "".join(ch for ch in str(school) if ch not in '\\/:*?"<>|').strip()
    filename = f"{safe_school}学生素养测评数据分析报告.docx"

    with tempfile.TemporaryDirectory(prefix="school-report-download-") as temp:
        output = Path(temp) / filename
        try:
            report_engine.build_report(school, year, datasets, output)
        except report_engine.ReportError as exc:
            raise RequestError(str(exc)) from exc
        body = output.read_bytes()

    response = send_file(
        BytesIO(body),
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        as_attachment=True,
        download_name=filename,
    )
    response.headers["Content-Disposition"] = f"attachment; filename*=UTF-8''{quote(filename)}"
    return response


@app.errorhandler(RequestError)
def handle_report_error(error):
    return jsonify(error=str(error)), 400


@app.errorhandler(RequestEntityTooLarge)
def handle_too_large(_error):
    return jsonify(error="单次上传文件合计不能超过 4MB"), 413


@app.errorhandler(BadRequest)
def handle_bad_request(_error):
    return jsonify(error="请求数据格式不正确"), 400


@app.errorhandler(Exception)
def handle_unexpected_error(error):
    app.logger.exception("Unhandled report processing error", exc_info=error)
    return jsonify(error="处理失败，请检查报告格式后重试"), 500


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=18968)
