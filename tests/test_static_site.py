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
    page.get_by_role("button", name="產生 OUTPUT").click()
    playwright.expect(page.locator("#results")).to_be_visible(timeout=300_000)
    playwright.expect(page.locator("#act-count")).to_have_text("50,001")
    playwright.expect(page.locator("#budget-count")).to_have_text("50,001")
    with page.expect_download() as captured:
        page.locator("#download").click()
    output = tmp_path / "large-output.xlsx"
    captured.value.save_as(output)
    book = load_workbook(output, data_only=True)
    row = next(row for row in book.active if row[0].value == "N20")
    assert row[1].value == pytest.approx(5.750115)
    assert row[2].value == pytest.approx(5.0001)
    assert row[8].value == 1.15
    assert "Large Customer +0.8M" in row[11].value
    book.close()
    assert all(method in {"GET", "HEAD"} for method, _ in requests)
    assert not errors
