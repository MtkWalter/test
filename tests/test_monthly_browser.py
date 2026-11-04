"""Optional real-browser tests: install the browser-test extra to run these."""
import json
import os
import shutil
import threading

import pytest
from openpyxl import load_workbook

playwright = pytest.importorskip("playwright.sync_api")

from excel_insights.monthly_web import LocalReportServer
from excel_insights.report_sample import create_sample


@pytest.fixture
def browser_page():
    server = LocalReportServer()
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
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
        page = browser.new_page(viewport={"width": 1440, "height": 1050})
        errors, requests = [], []
        page.on("pageerror", lambda error: errors.append(str(error)))
        page.on("request", lambda request: requests.append(request.url))
        try:
            page.goto(server.url)
            playwright.expect(page.locator("#year")).not_to_have_value("")
            yield page, server, errors, requests
        finally:
            browser.close()
            server.shutdown()
            thread.join(timeout=5)
            server.close()


def select_sample(page, tmp_path):
    source = create_sample(tmp_path / "input.xlsx")
    page.locator("#input-file").set_input_files(source)
    playwright.expect(page.locator("#file-name")).to_have_text("input.xlsx")
    playwright.expect(page.locator("#generate")).to_be_enabled()
    page.locator("#year").fill("2026")
    page.locator("#month").select_option("9")
    page.locator("#fiscal-year").fill("FY27")
    return source


def test_browser_upload_preview_expand_and_download(browser_page, tmp_path):
    page, server, errors, requests = browser_page
    select_sample(page, tmp_path)
    page.get_by_role("button", name="產生 OUTPUT").click()
    playwright.expect(page.locator("#results")).to_be_visible(timeout=20_000)
    playwright.expect(page.locator("#act-count")).to_have_text("20")
    row = page.locator("#preview-body tr").filter(has=page.locator("td:first-child", has_text="N20"))
    playwright.expect(row).to_have_count(1)
    playwright.expect(row.locator("td").nth(8)).to_have_text("120%")
    assert "tone-high" in row.locator("td").nth(8).get_attribute("class")
    row.get_by_role("button", name="展開完整 Remark").click()
    remark = row.locator(".remark-text").inner_text()
    assert remark.index("END_A +40.0M") < remark.index("END_C +10.0M") < remark.index("END_B -30.0M")
    assert "Non-T\nNON_T_DOWN -50.0M" in remark
    with page.expect_download() as captured:
        page.get_by_role("link", name="下載 OUTPUT Excel", exact=True).click()
    destination = tmp_path / "download.xlsx"
    captured.value.save_as(destination)
    book = load_workbook(destination, data_only=True)
    saved = next(row for row in book.active if row[0].value == "N20")
    assert saved[1].value == 120 and saved[3].value == 80
    book.close()
    assert not errors
    assert requests and all(url.startswith(server.origin + "/") for url in requests)


def test_browser_rule_files_and_mobile_layout(browser_page, tmp_path):
    page, server, errors, requests = browser_page
    page.set_viewport_size({"width": 390, "height": 844})
    assert page.evaluate("document.documentElement.scrollWidth <= window.innerWidth")
    page.locator("#classification summary").click()
    rules = {"tsmc_values": ["tsmc"], "non_t_values": ["non-t", "FI"],
             "node_aliases": {"N2": "N2+A16", "A16": "N2+A16"}, "fiscal_start_month": 4}
    config = tmp_path / "rules.json"
    config.write_text(json.dumps(rules), encoding="utf-8")
    page.locator("#rules-file").set_input_files(config)
    playwright.expect(page.locator("#non-t-values")).to_have_value("non-t、FI")
    with page.expect_download() as captured:
        page.get_by_role("button", name="儲存分類設定").click()
    saved = tmp_path / "saved-rules.json"
    captured.value.save_as(saved)
    assert json.loads(saved.read_text(encoding="utf-8")) == rules
    select_sample(page, tmp_path)
    page.get_by_role("button", name="產生 OUTPUT").click()
    playwright.expect(page.locator("#results")).to_be_visible(timeout=20_000)
    playwright.expect(page.locator("#preview-body tr").filter(has=page.locator("td:first-child", has_text="N2+A16"))).to_have_count(1)
    assert page.evaluate("document.documentElement.scrollWidth <= window.innerWidth")
    assert not errors


def test_browser_displays_bad_source_and_can_close(browser_page, tmp_path):
    page, server, errors, requests = browser_page
    source = select_sample(page, tmp_path)
    book = load_workbook(source)
    book.active["AS3"] = None
    book.save(source)
    book.close()
    page.locator("#input-file").set_input_files(source)
    playwright.expect(page.locator("#generate")).to_be_enabled()
    page.get_by_role("button", name="產生 OUTPUT").click()
    playwright.expect(page.locator("#form-error")).to_contain_text("AS 實績金額空白", timeout=20_000)
    playwright.expect(page.locator("#results")).to_be_hidden()
    page.on("dialog", lambda dialog: dialog.accept())
    page.get_by_role("button", name="關閉程式").click()
    playwright.expect(page.locator("#connection-error")).to_contain_text("本機程式正在結束")
    playwright.expect(page.locator("#generate")).to_be_disabled()
    assert not errors
