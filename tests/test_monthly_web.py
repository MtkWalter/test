import json
import threading
import time
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import ProxyHandler, Request, build_opener

import pytest
from openpyxl import load_workbook

from excel_insights.monthly_report import ReportRules
from excel_insights.monthly_web import LocalReportServer, smoke_browser_app
from excel_insights.report_sample import create_sample


@pytest.fixture
def web_server():
    server = LocalReportServer()
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    try:
        yield server
    finally:
        server.shutdown()
        thread.join(timeout=5)
        server.close()


def request(server, route, data=None, *, headers=None):
    if isinstance(data, dict):
        data = json.dumps(data).encode()
        headers = {"Content-Type": "application/json", **(headers or {})}
    req = Request(server.url + route, data=data, headers=headers or {})
    try:
        response = build_opener(ProxyHandler({})).open(req, timeout=30)
    except HTTPError as exc:
        response = exc
    with response:
        return response.status, response.headers, response.read()


def upload(server, path, name=None):
    status, _, body = request(server, "api/upload?filename=" + quote(name or path.name), path.read_bytes())
    assert status == 201, body
    return json.loads(body)


def wait_for_job(server, job_id):
    deadline = time.monotonic() + 10
    while time.monotonic() < deadline:
        status, _, body = request(server, "api/jobs/" + job_id)
        assert status == 200
        job = json.loads(body)
        if job["status"] != "running":
            return job
        time.sleep(.02)
    pytest.fail("HTTP report job did not complete")


def test_browser_http_upload_generate_preview_and_download(web_server, tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    before = source.read_bytes()
    status, headers, page = request(web_server, "")
    assert status == 200 and b"app.js" in page
    assert "frame-ancestors 'none'" in headers["Content-Security-Policy"]
    for route, expected in [("app.js", b"uploadFile"), ("style.css", b".sidebar")]:
        assert expected in request(web_server, route)[2]
    input_file = upload(web_server, source)
    payload = {"upload_id": input_file["id"], "sheet": "銷售明細", "year": 2026, "month": 9, "filename": "月報-九月.xlsx"}
    status, _, body = request(web_server, "api/jobs", payload)
    assert status == 202
    job = wait_for_job(web_server, json.loads(body)["id"])
    assert job["status"] == "success", job
    assert job["result"]["act_rows"] == job["result"]["ori_rows"] == 20
    assert "output" not in job["result"]  # No internal filesystem path in the page.
    rows = {row["cells"][0]: row for row in job["preview"]["rows"]}
    assert rows["N20"]["cells"][1:11] == ["120", "100", "80", "100", "200", "200", "20", "120%", "-20", "80%"]
    assert rows["N20"]["tsmc"] == "high" and rows["N20"]["non_t"] == "low"
    assert "END_A +40.0M" in rows["N20"]["cells"][11]
    assert rows["N7"]["cells"][8] == ""
    assert rows["N7"]["cells"][1] == "2"  # Excel's half-away-from-zero display.
    assert rows["N40"]["cells"][8] == "115%" and rows["N40"]["tsmc"] == ""
    status, headers, body = request(web_server, job["download"])
    assert status == 200 and body.startswith(b"PK")
    assert quote("月報-九月.xlsx") in headers["Content-Disposition"]
    destination = tmp_path / "download.xlsx"
    destination.write_bytes(body)
    book = load_workbook(destination, data_only=True)
    row = next(row for row in book.active if row[0].value == "N20")
    assert row[1].value == 120 and row[10].value == .8
    book.close()
    assert source.read_bytes() == before


def test_missing_actual_becomes_job_error_not_download(web_server, tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    book = load_workbook(source)
    book.active["AS3"] = None
    book.save(source)
    book.close()
    item = upload(web_server, source)
    status, _, body = request(web_server, "api/jobs", {"upload_id": item["id"], "sheet": item["sheets"][0], "year": 2026, "month": 9})
    job_id = json.loads(body)["id"]
    job = wait_for_job(web_server, job_id)
    assert job["status"] == "error" and "AS 實績金額空白" in job["error"]
    assert not job["download"]
    assert request(web_server, "api/download/" + job_id)[0] == 404


def test_loopback_host_origin_and_session_are_required(web_server):
    assert web_server.server_address[0] == "127.0.0.1"
    assert request(web_server, "api/state", headers={"Host": "untrusted.invalid"})[0] == 403
    assert request(web_server, "api/state", headers={"Origin": "https://untrusted.invalid"})[0] == 403
    assert request(web_server, "api/shutdown", {}, headers={"Origin": "null"})[0] == 403
    with pytest.raises(HTTPError) as exc:
        build_opener(ProxyHandler({})).open(web_server.origin + "/", timeout=10)
    assert exc.value.code == 404
    assert request(web_server, "api/download/../../anything")[0] == 404


def test_upload_checks_file_type_and_keeps_paths_in_local_temp(web_server, tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    item = upload(web_server, source, "../../outside.xlsx")
    path = web_server.uploads[item["id"]].path
    assert web_server.root in path.parents
    assert ".." not in path.relative_to(web_server.root).parts
    assert request(web_server, "api/upload?filename=test.exe", b"abc")[0] == 400
    status, _, body = request(web_server, "api/upload?filename=test.xlsx", b"broken")
    assert status == 400 and "無法辨識" in json.loads(body)["error"]
    assert len(web_server.uploads) == 1


def test_job_validates_classifications_and_period(web_server, tmp_path):
    item = upload(web_server, create_sample(tmp_path / "input.xlsx"))
    valid = {"upload_id": item["id"], "sheet": item["sheets"][0], "year": 2026, "month": 9}
    for update in ({"month": 13}, {"year": True}, {"sheet": "unknown"}, {"rules": {"fiscal_start_month": True}},
                   {"rules": {"tsmc_values": None}}, {"fiscal_year": "FYinvalid"}, {"upload_id": "unknown"}):
        assert request(web_server, "api/jobs", {**valid, **update})[0] == 400
    assert not web_server.jobs


def test_one_report_at_a_time_and_safe_shutdown(web_server, tmp_path, monkeypatch):
    import excel_insights.monthly_web as web
    release = threading.Event()
    original = web.generate_report
    def paused(*args, **kwargs):
        release.wait(5)
        return original(*args, **kwargs)
    monkeypatch.setattr(web, "generate_report", paused)
    item = upload(web_server, create_sample(tmp_path / "input.xlsx"))
    payload = {"upload_id": item["id"], "sheet": item["sheets"][0], "year": 2026, "month": 9}
    try:
        status, _, body = request(web_server, "api/jobs", payload)
        assert status == 202
        assert request(web_server, "api/jobs", payload)[0] == 409
        assert request(web_server, "api/shutdown", {})[0] == 409
    finally:
        release.set()
    assert wait_for_job(web_server, json.loads(body)["id"])["status"] == "success"
    assert request(web_server, "api/shutdown", {})[0] == 200


def test_smoke_uses_http_and_cleans_temporary_directory(tmp_path, monkeypatch):
    import excel_insights.monthly_web as web
    directories = []
    original = web.LocalReportServer
    def capture(*args, **kwargs):
        server = original(*args, **kwargs)
        directories.append(server.root)
        return server
    monkeypatch.setattr(web, "LocalReportServer", capture)
    source = create_sample(tmp_path / "input.xlsx")
    output, summary = tmp_path / "output.xlsx", tmp_path / "summary.json"
    assert smoke_browser_app(source, output, summary, 2026, 9) == 0
    assert json.loads(summary.read_text(encoding="utf-8"))["interface"] == "browser"
    assert output.exists() and directories and all(not path.exists() for path in directories)


def test_report_rules_dictionary_validation():
    with pytest.raises(ValueError):
        ReportRules.from_dict({"non_t_values": "not-a-list"})
    assert ReportRules.from_dict(ReportRules().to_dict()).to_dict() == ReportRules().to_dict()


def test_raw_export_job_includes_complete_errors_and_source_package(web_server, tmp_path):
    from tests.test_raw_export import assert_raw_preserved
    source = create_sample(tmp_path / "input.xlsx")
    book = load_workbook(source)
    book.active["AS3"] = "invalid actual"
    book.active["BF3"] = None
    book.active["BM3"] = None
    book.save(source)
    book.close()
    item = upload(web_server, source)
    payload = {"upload_id": item["id"], "sheet": item["report_sheets"][0], "year": 2026, "month": 9, "include_raw": True}
    status, _, body = request(web_server, "api/jobs", payload)
    assert status == 202
    job = wait_for_job(web_server, json.loads(body)["id"])
    assert job["status"] == "success", job
    assert job["result"]["report_generated"] is False
    assert len(job["result"]["issues"]) == 3
    assert job["preview"]["rows"] == []
    status, _, content = request(web_server, job["download"])
    output = tmp_path / "download.xlsx"
    output.write_bytes(content)
    assert_raw_preserved(source, output)
