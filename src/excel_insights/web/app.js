"use strict";
const $ = (id) => document.getElementById(id);
const form = $("report-form");
let upload = null;
let busy = false;
let closed = false;
let autoFilename = "";
let defaultRules = {tsmc_values: ["tsmc"], non_t_values: null, node_aliases: {N2: "N2/A16", A16: "N2/A16", N3: "N3/N3E", N3E: "N3/N3E"}, fiscal_start_month: 4};

function showError(message) {
  $("form-error").textContent = message;
  $("form-error").hidden = !message;
}
function resetResult() {
  $("results").hidden = true;
  $("step-results").classList.remove("active", "done");
}
function setBusy(value) {
  busy = value;
  for (const input of form.querySelectorAll("input, select, textarea, button")) input.disabled = value || closed;
  $("sheet").disabled = value || !upload || closed;
  $("generate").disabled = value || !upload || closed;
  $("stop-app").disabled = value || closed;
  $("generate").textContent = value ? "正在處理…" : "產生 OUTPUT →";
}
function progress(message, percent = null) {
  $("progress-card").hidden = false;
  $("progress-message").textContent = message;
  if (percent === null) $("progress-bar").removeAttribute("value");
  else $("progress-bar").value = percent;
}
async function api(path, data) {
  if (window.ReportTransport) return window.ReportTransport.api(path, data);
  if (document.body.dataset.reportMode === "static") throw new Error("網頁資源尚未載入，請重新整理頁面。");
  const options = data === undefined ? {} : {method: "POST", headers: {"Content-Type": "application/json"}, body: JSON.stringify(data)};
  let response;
  try { response = await fetch(path, options); }
  catch { throw new Error("無法連接本機程式，請重新開啟 MonthlyReport.exe。"); }
  let value;
  try { value = await response.json(); }
  catch { throw new Error("本機連線已失效，請重新開啟 MonthlyReport.exe。"); }
  if (!response.ok) throw new Error(value.error || "操作未完成，請重試。");
  return value;
}
function fillMonths(select) {
  for (let month = 1; month <= 12; month++) {
    const option = document.createElement("option");
    option.value = String(month);
    option.textContent = `${String(month).padStart(2, "0")} 月`;
    select.append(option);
  }
}
function updatePeriod() {
  const year = Number($("year").value);
  const month = Number($("month").value);
  const start = Number($("fiscal-start").value);
  if (year && month && start) {
    const end = start === 1 ? year : year + (month >= start ? 1 : 0);
    $("fiscal-year").value = `FY${String(end % 100).padStart(2, "0")}`;
    const name = `monthly-output-${year}${String(month).padStart(2, "0")}.xlsx`;
    if (!$("output-name").value || $("output-name").value === autoFilename) $("output-name").value = name;
    autoFilename = name;
  }
}
function addAlias(source = "", target = "") {
  const row = document.createElement("div");
  row.className = "alias-row";
  const a = document.createElement("input");
  a.value = source; a.placeholder = "例如 N2"; a.setAttribute("aria-label", "INPUT Tech Node");
  const b = document.createElement("input");
  b.value = target; b.placeholder = "例如 N2/A16"; b.setAttribute("aria-label", "OUTPUT Tech Node");
  const remove = document.createElement("button");
  remove.type = "button"; remove.className = "alias-remove"; remove.textContent = "×"; remove.setAttribute("aria-label", "移除此對照");
  remove.addEventListener("click", () => {row.remove(); resetResult();});
  row.append(a, b, remove);
  $("alias-list").append(row);
}
function applyRules(rules) {
  if (!rules || typeof rules !== "object" || Array.isArray(rules)) throw new Error("分類設定必須是有效的 JSON 物件。");
  const allowed = ["tsmc_values", "non_t_values", "node_aliases", "fiscal_start_month"];
  if (Object.keys(rules).some((name) => !allowed.includes(name))) throw new Error("分類設定包含未知的欄位名稱。");
  const get = (name) => Object.prototype.hasOwnProperty.call(rules, name) ? rules[name] : defaultRules[name];
  const tsmc = get("tsmc_values");
  const nonT = get("non_t_values");
  if (!Array.isArray(tsmc) || !tsmc.length || tsmc.some((v) => typeof v !== "string" || !v.trim())) throw new Error("TSMC 分類不得空白。");
  if (nonT !== null && (!Array.isArray(nonT) || nonT.some((v) => typeof v !== "string"))) throw new Error("Non-T 分類格式不正確。");
  const aliases = get("node_aliases");
  if (!aliases || typeof aliases !== "object" || Array.isArray(aliases) || Object.entries(aliases).some(([a, b]) => !a.trim() || typeof b !== "string" || !b.trim())) throw new Error("Tech Node 對照格式不正確。");
  const start = get("fiscal_start_month");
  if (!Number.isInteger(start) || start < 1 || start > 12) throw new Error("會計年度起始月份必須是 1–12。");
  $("tsmc-values").value = tsmc.join("、");
  $("non-t-values").value = (nonT || []).join("、");
  $("fiscal-start").value = String(start);
  $("alias-list").replaceChildren();
  for (const [a, b] of Object.entries(aliases)) addAlias(a, b);
  updatePeriod();
}
function readRules() {
  const split = (text) => text.split(/[,，、;；\n]/u).map((s) => s.trim()).filter(Boolean);
  const aliases = Object.create(null);
  const used = new Set();
  for (const row of $("alias-list").children) {
    const inputs = row.querySelectorAll("input");
    const a = inputs[0].value.trim();
    const b = inputs[1].value.trim();
    if (!a && !b) continue;
    if (!a || !b) throw new Error("Tech Node 的 INPUT／OUTPUT 名稱都必須填寫。");
    const k = a.normalize("NFKC").replace(/\s/gu, "").toLowerCase();
    if (used.has(k)) throw new Error(`Tech Node 對照重複：${a}`);
    used.add(k); aliases[a] = b;
  }
  const tsmc = split($("tsmc-values").value);
  if (!tsmc.length) throw new Error("請填寫 TSMC 的 BF 值。");
  const nonT = split($("non-t-values").value);
  return {tsmc_values: tsmc, non_t_values: nonT.length ? nonT : null, node_aliases: aliases, fiscal_start_month: Number($("fiscal-start").value)};
}
function uploadFile(file) {
  if (window.ReportTransport) return window.ReportTransport.upload(file, (message) => progress(message));
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("POST", `api/upload?filename=${encodeURIComponent(file.name)}`);
    xhr.setRequestHeader("Content-Type", "application/octet-stream");
    xhr.upload.onprogress = (event) => {
      if (event.lengthComputable) {
        const pct = Math.round(event.loaded / event.total * 100);
        progress(pct < 100 ? `正在傳送至本機程式… ${pct}%` : "正在辨識 Excel 工作表…", pct < 100 ? pct : null);
      }
    };
    xhr.onload = () => {
      try {
        const data = JSON.parse(xhr.responseText);
        if (xhr.status >= 200 && xhr.status < 300) resolve(data);
        else reject(new Error(data.error || "無法讀取 Excel。"));
      } catch { reject(new Error("本機程式未回應，請重新啟動。")); }
    };
    xhr.onerror = () => reject(new Error("無法連接本機程式，請重新開啟 MonthlyReport.exe。"));
    xhr.send(file);
  });
}
async function chooseInput(file) {
  if (!file || busy || closed) return;
  showError(""); resetResult();
  if (!/\.(xlsx|xlsm)$/iu.test(file.name)) {
    showError("請選擇 .xlsx 或 .xlsm Excel 檔案。"); return;
  }
  setBusy(true); progress("正在讀取 Excel…");
  try {
    upload = await uploadFile(file);
    $("file-name").textContent = upload.name;
    const size = upload.bytes < 1024 * 1024 ? `${(upload.bytes / 1024).toFixed(1)} KB` : `${(upload.bytes / 1024 / 1024).toFixed(1)} MB`;
    $("file-size").textContent = `${size} · ${upload.sheets.length} 張可用工作表`;
    $("sheet").replaceChildren(...upload.sheets.map((name) => {const option = document.createElement("option"); option.value = name; option.textContent = name; return option;}));
    $("file-summary").hidden = false; $("drop-zone").hidden = true;
    $("step-source").classList.replace("active", "done");
    $("step-settings").classList.add("active");
  } catch (error) { showError(error.message); }
  finally {setBusy(false); $("progress-card").hidden = true; $("input-file").value = "";}
}
function renderResult(job) {
  const result = job.result;
  $("result-period").textContent = `${result.year} 年 ${String(result.month).padStart(2, "0")} 月 · ${result.fiscal_year} · ${result.sheet}`;
  $("act-count").textContent = result.act_rows.toLocaleString("zh-TW");
  $("budget-count").textContent = result.ori_rows.toLocaleString("zh-TW");
  $("remark-count").textContent = result.remark_nodes.toLocaleString("zh-TW");
  $("warnings").hidden = !result.warnings.length;
  $("warnings").textContent = result.warnings.join("\n");
  $("download").href = job.download;
  $("download").setAttribute("download", job.filename);
  const fragment = document.createDocumentFragment();
  for (const row of job.preview.rows) {
    const tr = document.createElement("tr");
    row.cells.forEach((value, i) => {
      const td = document.createElement("td");
      if (i === 5 || i === 6) td.classList.add("total-cell");
      if ((i === 7 || i === 8) && row.tsmc) td.classList.add(`tone-${row.tsmc}`);
      if ((i === 9 || i === 10) && row.non_t) td.classList.add(`tone-${row.non_t}`);
      if (i === 11) {
        td.classList.add("remark-cell");
        const body = document.createElement("div"); body.className = "remark-text"; body.textContent = value; td.append(body);
        if (value.split("\n").length > 3 || value.length > 170) {
          const toggle = document.createElement("button"); toggle.type = "button"; toggle.className = "button-link remark-toggle"; toggle.textContent = "展開完整 Remark"; toggle.setAttribute("aria-expanded", "false");
          toggle.addEventListener("click", () => {const expanded = body.classList.toggle("expanded"); toggle.setAttribute("aria-expanded", String(expanded)); toggle.textContent = expanded ? "收合 Remark" : "展開完整 Remark";});
          td.append(toggle);
        }
      } else td.textContent = value;
      tr.append(td);
    });
    fragment.append(tr);
  }
  $("preview-body").replaceChildren(fragment);
  $("preview-truncated").hidden = !job.preview.truncated;
  $("results").hidden = false;
  $("step-settings").classList.remove("active"); $("step-settings").classList.add("done"); $("step-results").classList.add("active");
  $("results").scrollIntoView({block: "start"});
}
async function generate(event) {
  event.preventDefault();
  if (!upload || busy || closed || !form.reportValidity()) return;
  showError(""); resetResult();
  let rules;
  try {rules = readRules();} catch (error) {showError(error.message); return;}
  const payload = {upload_id: upload.id, sheet: $("sheet").value, year: Number($("year").value), month: Number($("month").value), fiscal_year: $("fiscal-year").value.toUpperCase(), filename: $("output-name").value, rules};
  setBusy(true); progress("正在建立單月報表…");
  try {
    const created = await api("api/jobs", payload);
    for (;;) {
      const job = await api(`api/jobs/${created.id}`);
      progress(job.progress);
      if (job.status === "success") {renderResult(job); break;}
      if (job.status === "error") throw new Error(job.error);
      await new Promise((resolve) => setTimeout(resolve, 500));
    }
  } catch (error) {showError(error.message); $("form-error").scrollIntoView({block: "center"});}
  finally {setBusy(false); $("progress-card").hidden = true;}
}

$("choose-file").addEventListener("click", () => $("input-file").click());
$("replace-file").addEventListener("click", () => $("input-file").click());
$("input-file").addEventListener("change", (event) => chooseInput(event.target.files[0]));
for (const name of ["dragenter", "dragover"]) $("drop-zone").addEventListener(name, (event) => {event.preventDefault(); if (!busy) $("drop-zone").classList.add("dragover");});
for (const name of ["dragleave", "drop"]) $("drop-zone").addEventListener(name, (event) => {event.preventDefault(); $("drop-zone").classList.remove("dragover");});
$("drop-zone").addEventListener("drop", (event) => chooseInput(event.dataTransfer.files[0]));
form.addEventListener("submit", generate);
form.addEventListener("input", resetResult);
for (const name of ["year", "month", "fiscal-start"]) $(name).addEventListener("change", updatePeriod);
$("add-alias").addEventListener("click", () => {addAlias(); resetResult();});
$("load-rules").addEventListener("click", () => $("rules-file").click());
$("rules-file").addEventListener("change", async (event) => {
  const file = event.target.files[0]; if (!file) return;
  try {applyRules(JSON.parse((await file.text()).replace(/^\uFEFF/u, ""))); resetResult(); showError("");}
  catch (error) {showError(`無法載入分類設定：${error.message}`);}
  event.target.value = "";
});
$("save-rules").addEventListener("click", () => {
  try {
    const blob = new Blob([JSON.stringify(readRules(), null, 2)], {type: "application/json"});
    const url = URL.createObjectURL(blob); const a = document.createElement("a");
    a.href = url; a.download = "monthly-report-rules.json"; a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000); showError("");
  } catch (error) {showError(error.message);}
});
$("stop-app").addEventListener("click", async () => {
  const isStatic = window.ReportTransport?.mode === "static";
  const question = isStatic ? "清除瀏覽器中的 Excel 與運算資料？已下載的 OUTPUT 會保留。" : "關閉本機程式並清除暫存？已下載的 Excel 會保留。";
  if (busy || !window.confirm(question)) return;
  try {
    await api("api/shutdown", {});
    if (isStatic) {
      upload = null; resetResult(); setBusy(false);
      $("file-summary").hidden = true; $("drop-zone").hidden = false;
      $("file-name").textContent = ""; $("file-size").textContent = "";
      $("sheet").replaceChildren(new Option("先選擇 INPUT Excel", ""));
      $("preview-body").replaceChildren(); $("warnings").textContent = "";
      $("result-period").textContent = "";
      for (const id of ["act-count", "budget-count", "remark-count"]) $(id).textContent = "0";
      $("step-source").classList.remove("done"); $("step-source").classList.add("active");
      $("step-settings").classList.remove("active", "done");
      $("download").removeAttribute("href"); showError("");
      return;
    }
    closed = true; setBusy(false);
    $("connection-error").textContent = "本機程式正在結束並清除暫存。若要繼續，請重新開啟 MonthlyReport.exe。";
    $("connection-error").hidden = false; $("download").removeAttribute("href");
  } catch (error) {showError(error.message);}
});
async function initialize() {
  fillMonths($("month")); fillMonths($("fiscal-start"));
  setBusy(true);
  try {
    const state = await api("api/state");
    defaultRules = state.rules;
    $("year").value = state.year; $("month").value = String(state.month);
    applyRules(state.rules); setBusy(false);
  } catch (error) {
    closed = true; setBusy(false); $("connection-error").textContent = error.message; $("connection-error").hidden = false;
  }
}
initialize();
