/* Python runs in this browser worker, with a memory-only filesystem. */
"use strict";
let pythonReady;
let py;
const inputs = new Map();
const version = new URL(self.location.href).search;
const assetURL = (name) => new URL(name + version, self.location.href).href;
const notify = (message) => self.postMessage({event: "progress", message: String(message)});

async function fetchBytes(name, expected) {
  const response = await fetch(assetURL(name));
  if (!response.ok) throw new Error(`無法載入運算資源：${name}。請重新整理網頁。`);
  const buffer = await response.arrayBuffer();
  if (expected) {
    const digest = await crypto.subtle.digest("SHA-256", buffer);
    const hex = Array.from(new Uint8Array(digest), (n) => n.toString(16).padStart(2, "0")).join("");
    if (hex !== expected) throw new Error("網頁資源版本不一致，請重新整理後再試。");
  }
  return new Uint8Array(buffer);
}

async function initialize() {
  notify("正在載入瀏覽器運算核心，首次開啟需要一些時間…");
  const response = await fetch(assetURL("site-manifest.json"));
  if (!response.ok) throw new Error("無法載入網站設定，請重新整理。");
  const manifest = await response.json();
  importScripts(new URL("runtime/pyodide.js", self.location.href).href);
  py = await loadPyodide({indexURL: new URL("runtime/", self.location.href).href,
                         stdout: () => {}, stderr: () => {}});
  notify("正在準備 Excel 讀寫與既有報表規則…");
  const [dependencies, engine] = await Promise.all([
    fetchBytes("python-deps.zip", manifest.dependencies_sha256),
    fetchBytes("report-engine.zip", manifest.engine_sha256),
  ]);
  py.unpackArchive(dependencies, "zip", {extractDir: "/app/vendor"});
  py.unpackArchive(engine, "zip", {extractDir: "/app"});
  await py.runPythonAsync(`
import sys, json, gc
sys.path[:0] = ['/app', '/app/vendor']
from pathlib import Path
from excel_insights.monthly_report import generate_report, ReportRules, ReportValidationError
from excel_insights.raw_export import inspect_raw_workbook, generate_report_with_raw
from excel_insights.report_preview import preview_workbook

def _inspect_browser_input(filename):
    return inspect_raw_workbook(filename)

def _generate_browser_report(payload_json, progress):
    payload = json.loads(payload_json)
    try:
        if payload.get('include_raw'):
            result = generate_report_with_raw(payload['source'], payload['output'], payload['year'], payload['month'],
                sheet=payload['sheet'], fiscal_year=payload.get('fiscal_year'), rules=payload.get('rules', {}),
                settings_error=payload.get('settings_error'), progress=progress)
        else:
            result = generate_report(payload['source'], payload['output'], payload['year'], payload['month'],
                sheet=payload['sheet'], fiscal_year=payload.get('fiscal_year'),
                rules=ReportRules.from_dict(payload.get('rules', {})), progress=progress)
    except ReportValidationError as exc:
        return json.dumps({'error': f'來源資料共有 {len(exc.issues):,} 項錯誤，請查看完整清單或匯出 Raw Data。',
                           'issues': exc.issues}, ensure_ascii=False)
    output = result.pop('output')
    preview = preview_workbook(Path(output), result.get('preview_sheet') or 'monthly pivot') if result.get('report_generated', True) else {'rows': [], 'truncated': False}
    return json.dumps({'result': result, 'preview': preview, 'output': output}, ensure_ascii=False)
`);
  py.FS.mkdirTree("/inputs");
  py.FS.mkdirTree("/reports");
}

function ready() {
  if (!pythonReady) pythonReady = initialize();
  return pythonReady;
}

function errorMessage(error) {
  const message = String(error.message || error);
  const match = message.match(/(?:ValueError|FileNotFoundError|FileExistsError|RuntimeError|MemoryError|BadZipFile): ([\s\S]*)$/u);
  return match?.[1].trim() || message;
}

self.onmessage = async (event) => {
  const {id, action, payload} = event.data;
  try {
    await ready();
    if (action === "upload") {
      const folder = `/inputs/${payload.id}`;
      const name = payload.name.replace(/[\\/\u0000]/gu, "_");
      const path = `${folder}/${name}`;
      py.FS.mkdirTree(folder);
      py.FS.writeFile(path, new Uint8Array(payload.bytes));
      try {
        notify("正在瀏覽器中辨識 Excel 工作表…");
        py.globals.set("_browser_input", path);
        const catalog = JSON.parse(await py.runPythonAsync("json.dumps(_inspect_browser_input(_browser_input), ensure_ascii=False)"));
        for (const old of inputs.values()) {
          py.FS.unlink(old);
          py.FS.rmdir(old.slice(0, old.lastIndexOf("/")));
        }
        inputs.clear(); inputs.set(payload.id, path);
        self.postMessage({id, value: {id: payload.id, name: payload.name, bytes: payload.size, ...catalog}});
      } catch (error) {
        py.FS.unlink(path); py.FS.rmdir(folder); throw error;
      }
    } else if (action === "generate") {
      const path = inputs.get(payload.upload_id);
      if (!path) throw new Error("請重新選擇 INPUT Excel。");
      let output = `/reports/${payload.id}.xlsx`;
      py.globals.set("_browser_job", JSON.stringify({...payload, source: path, output}));
      py.globals.set("_browser_progress", notify);
      try {
        const result = JSON.parse(await py.runPythonAsync("_generate_browser_report(_browser_job, _browser_progress)"));
        if (result.error) {self.postMessage({id, error: result.error, issues: result.issues}); return;}
        output = result.output; delete result.output;
        const bytes = py.FS.readFile(output).slice();
        self.postMessage({id, value: result, bytes: bytes.buffer}, [bytes.buffer]);
      } finally {
        try {py.FS.unlink(output);} catch { /* no output after validation error */ }
        await py.runPythonAsync("gc.collect()");
      }
    } else throw new Error("不支援的瀏覽器操作。");
  } catch (error) {
    self.postMessage({id, error: errorMessage(error)});
  }
};
