"""Real browser tests against static files or a deployed GitHub Pages URL."""
import functools
import json
import os
import shutil
import threading
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

import pytest
from openpyxl import Workbook, load_workbook

playwright = pytest.importorskip("playwright.sync_api")
from excel_insights.report_sample import create_sample

ROOT = Path(__file__).resolve().parents[1]


@pytest.fixture
def static_page():
    deployed = os.environ.get("REPORT_SITE_URL")
    server = None
    if not deployed:
        if not (ROOT / "_site/index.html").exists():
            pytest.skip("Build the static site with npm ci and scripts/build-report-site.py first.")
        class Handler(SimpleHTTPRequestHandler):
            def log_message(self, *_args):
                pass
        handler = functools.partial(Handler, directory=str(ROOT / "_site"))
        server = ThreadingHTTPServer(("127.0.0.1", 0), handler)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        url = f"http://127.0.0.1:{server.server_port}/"
    else:
        url = deployed.rstrip("/") + "/"
    with playwright.sync_playwright() as driver:
        options = {"headless": True}
        channel = os.environ.get("REPORT_BROWSER_CHANNEL")
        executable = os.environ.get("REPORT_BROWSER_EXECUTABLE") or shutil.which("chromium")
        if channel:
            options["channel"] = channel
        elif executable:
            options["executable_path"] = executable
        if os.name != "nt":
            options["args"] = ["--no-sandbox", "--disable-dev-shm-usage"]
        browser = driver.chromium.launch(**options)
        context = browser.new_context(viewport={"width": 1440, "height": 1050})
        page = context.new_page()
        requests, errors = [], []
        context.on("request", lambda request: requests.append((request.method, request.url)))
        page.on("pageerror", lambda error: errors.append(str(error)))
        try:
            page.goto(url)
            playwright.expect(page.locator("#year")).not_to_have_value("", timeout=30_000)
            yield page, url, requests, errors
        finally:
            browser.close()
            if server:
                server.shutdown()
                thread.join(timeout=5)
                server.server_close()


def choose_input(page, source):
    page.locator("#input-file").set_input_files(source)
    playwright.expect(page.locator("#generate")).to_be_enabled(timeout=120_000)
    page.locator("#year").fill("2026")
    page.locator("#month").select_option("9")
    page.locator("#fiscal-year").fill("FY27")


def test_static_site_roundtrip_and_no_workbook_upload(static_page, tmp_path):
    page, url, requests, errors = static_page
    source = create_sample(tmp_path / "input.xlsx")
    original = source.read_bytes()
    choose_input(page, source)
    page.get_by_role("button", name="產生 OUTPUT").click()
    playwright.expect(page.locator("#results")).to_be_visible(timeout=120_000)
    playwright.expect(page.locator("#act-count")).to_have_text("20")
    row = page.locator("#preview-body tr").filter(has=page.locator("td:first-child", has_text="N20"))
    row.get_by_role("button", name="展開完整 Remark").click()
    remark = row.locator(".remark-text").inner_text()
    assert remark.index("END_A +40.0M") < remark.index("END_C +10.0M") < remark.index("END_B -30.0M")
    assert remark.index("NON_T_DOWN -50.0M") < remark.index("NON_T_UP +30.0M")
    with page.expect_download() as captured:
        page.locator("#download").click()
    output = tmp_path / "download.xlsx"
    captured.value.save_as(output)
    book = load_workbook(output, data_only=True)
    rows = {row[0].value: row for row in book.active if row[0].value}
    assert [c.value for c in rows["N20"][:11]] == ["N20", 120, 100, 80, 100, 200, 200, 20, 1.2, -20, .8]
    assert rows["N2/A16"][8].value == 1.15
    assert rows["N3/N3E"][8].value == .85
    assert rows["N7"][8].value is None and rows["N14"][10].value is None
    assert rows["N40"][11].value is None
    book.close()
    assert source.read_bytes() == original
    origin = f"{urlsplit(url).scheme}://{urlsplit(url).netloc}/"
    assert requests and all(method in {"GET", "HEAD"} and target.startswith(origin) for method, target in requests)
    assert not errors
    if os.environ.get("REPORT_EXPECTED_REVISION"):
        manifest = page.request.get(url + "site-manifest.json").json()
        assert manifest["revision"] == os.environ["REPORT_EXPECTED_REVISION"]
    page.on("dialog", lambda dialog: dialog.accept())
    page.get_by_role("button", name="清除資料").click()
    playwright.expect(page.locator("#results")).to_be_hidden()
    playwright.expect(page.locator("#file-summary")).to_be_hidden()
    playwright.expect(page.locator("#generate")).to_be_disabled()
    assert page.locator("#preview-body").inner_text() == ""


def test_static_site_large_workbook_over_old_limits(static_page, tmp_path):
    page, url, requests, errors = static_page
    source = tmp_path / "large.xlsx"
    book = Workbook(write_only=True)
    sheet = book.create_sheet("來源")
    sheet.append(["sales-NTD", "budget / act", "TSMC/non TSMC", "Tech Node", "END USER", "P/N",
                  "Year(calender)", "Month(calender)", "YM(calender)", "supplier", "客戶簡稱", "CUSTOMER"])
    for scenario, amount in (("act", 115), ("ori", 100)):
        for _ in range(50_001):
            sheet.append([amount, f"FY27 {scenario}", "tsmc", "N20", "Large Customer", "PN",
                          2026, 9, "202609", "FT products", "NOT_END_USER", "NOT_END_USER"])
    book.save(source)
    book.close()
    choose_input(page, source)
    page.locator("#export-raw").click()
    playwright.expect(page.locator("#results")).to_be_visible(timeout=300_000)
    playwright.expect(page.locator("#act-count")).to_have_text("50,001")
    playwright.expect(page.locator("#budget-count")).to_have_text("50,001")
    with page.expect_download() as captured:
        page.locator("#download").click()
    output = tmp_path / "large-output.xlsx"
    captured.value.save_as(output)
    book = load_workbook(output, data_only=True, read_only=True)
    assert "來源" in book.sheetnames
    row = next(row for row in book["monthly pivot"] if row[0].value == "N20")
    assert row[1].value == pytest.approx(5.750115)
    assert row[2].value == pytest.approx(5.0001)
    assert row[8].value == 1.15
    assert "Large Customer +0.8M" in row[11].value
    book.close()
    assert all(method in {"GET", "HEAD"} for method, _ in requests)
    assert not errors


def test_static_site_raw_errors_roundtrip_and_offline_html(static_page, tmp_path):
    page, url, requests, errors = static_page
    source = create_sample(tmp_path / "errors.xlsx")
    book = load_workbook(source)
    sheet = book.active
    for index in range(200):
        number = 100 + index
        for column, value in {"AS": f"bad-{index:03d} <img src=x onerror=alert(1)>", "BD": "FY27 act",
                              "BF": None, "BM": None, "BB": "ordinary", "BI": "Customer", "BC": "PN",
                              "BP": "202609", "BO": 2026, "BE": 9}.items():
            sheet[f"{column}{number}"] = value
    book.save(source)
    book.close()
    original = source.read_bytes()
    choose_input(page, source)
    page.get_by_role("button", name="產生 OUTPUT").click()
    playwright.expect(page.locator("#diagnostics")).to_be_visible(timeout=120_000)
    playwright.expect(page.locator("#diagnostic-count")).to_contain_text("600 項錯誤")
    playwright.expect(page.locator("#diagnostic-body tr")).to_have_count(500)
    assert page.locator("#diagnostic-scroll").evaluate("element => element.scrollHeight > element.clientHeight")
    page.locator("#errors-next").click()
    playwright.expect(page.locator("#diagnostic-body tr")).to_have_count(100)
    assert "bad-199" in page.locator("#diagnostic-body").inner_text()
    page.locator("#export-raw").click()
    playwright.expect(page.locator("#results")).to_be_visible(timeout=120_000)
    playwright.expect(page.locator("#result-title")).to_have_text("原始資料已完整保留")
    playwright.expect(page.locator("#report-preview")).to_be_hidden()
    playwright.expect(page.locator("#act-count")).to_have_text("—")
    with page.expect_download() as captured:
        page.locator("#download").click()
    output = tmp_path / "raw-output.xlsx"
    captured.value.save_as(output)
    from tests.test_raw_export import assert_raw_preserved
    assert_raw_preserved(source, output)
    book = load_workbook(output)
    assert len(list(book["完整錯誤清單"].values)) == 601
    assert "monthly pivot" not in book.sheetnames
    assert "bad-199" in book["銷售明細"]["AS299"].value
    book.close()
    with page.expect_download() as captured:
        page.locator("#download-result-html").click()
    html = tmp_path / "result.html"
    captured.value.save_as(html)
    assert "原始資料已完整保留" in html.read_text(encoding="utf-8")
    origin = f"{urlsplit(url).scheme}://{urlsplit(url).netloc}/"
    assert all(method in {"GET", "HEAD"} and target.startswith(origin) for method, target in requests)
    # Inspect the actual downloaded standalone HTML with networking disabled.
    offline = page.context.new_page()
    offline.route("http://**/*", lambda route: route.abort())
    offline.route("https://**/*", lambda route: route.abort())
    # set_content also works on managed browsers that forbid all file:// URLs.
    offline.set_content(html.read_text(encoding="utf-8"))
    assert offline.locator("tbody tr").count() == 600
    assert "bad-199" in offline.locator("tbody").inner_text()
    assert offline.locator("img, script").count() == 0
    assert offline.locator(".scroll").last.evaluate("element => element.scrollHeight > element.clientHeight")
    offline.close()
    page.set_viewport_size({"width": 390, "height": 844})
    assert page.evaluate("document.documentElement.scrollWidth <= window.innerWidth")
    if os.environ.get("REPORT_PREVIEW_OUTPUT"):
        destination = Path(os.environ["REPORT_PREVIEW_OUTPUT"])
        destination.parent.mkdir(parents=True, exist_ok=True)
        preview = html.read_text(encoding="utf-8").replace("<h1>", '<p><b>功能預覽：以下為 600 項合成測試錯誤，並非您的原始資料。</b></p><h1>', 1)
        destination.write_text(preview, encoding="utf-8")
    assert source.read_bytes() == original
    assert not errors

    if os.environ.get("REPORT_EXPECTED_REVISION"):
        example = page.request.get(url + "example-output.html")
        assert example.ok and "600 項合成測試錯誤" in example.text()


def test_static_site_missing_headers_raw_roundtrip(static_page, tmp_path):
    page, url, requests, errors = static_page
    source = tmp_path / "missing-headers.xlsx"
    book = Workbook()
    book.active.append(["Original", "No recognized report headers", "=1+1"])
    book.save(source)
    book.close()
    page.locator("#input-file").set_input_files(source)
    playwright.expect(page.locator("#export-raw")).to_be_enabled(timeout=120_000)
    playwright.expect(page.locator("#generate")).to_be_disabled()
    playwright.expect(page.locator("#form-error")).to_contain_text("必要欄名")
    page.locator("#export-raw").click()
    playwright.expect(page.locator("#results")).to_be_visible(timeout=120_000)
    playwright.expect(page.locator("#result-title")).to_have_text("原始資料已完整保留")
    with page.expect_download() as captured:
        page.locator("#download").click()
    output = tmp_path / "raw.xlsx"
    captured.value.save_as(output)
    from tests.test_raw_export import assert_raw_preserved
    assert_raw_preserved(source, output)
    assert all(method in {"GET", "HEAD"} for method, _ in requests)
    assert not errors
