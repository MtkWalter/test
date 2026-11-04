/* Static-site transport: no workbook data is sent through HTTP. */
"use strict";
(() => {
  const revision = new URL(document.currentScript.src).search;
  const assetURL = (name) => new URL(name + revision, document.baseURI).href;
  let worker;
  let nextID = 0;
  let loadingProgress;
  let activeJob;
  let currentUpload;
  let manifestPromise;
  const pending = new Map();
  const jobs = new Map();
  const urls = new Set();

  function getWorker() {
    if (!worker) {
      worker = new Worker(assetURL("report-worker.js"));
      worker.onmessage = (event) => {
        const message = event.data;
        if (message.event === "progress") {
          if (loadingProgress) loadingProgress(message.message);
          if (activeJob) activeJob.progress = message.message;
          return;
        }
        const request = pending.get(message.id);
        if (!request) return;
        pending.delete(message.id);
        if (message.error) request.reject(new Error(message.error));
        else request.resolve(message);
      };
      worker.onerror = (event) => {
        const error = new Error(event.message || "運算核心中斷，請重新整理並重新選擇 Excel。");
        for (const request of pending.values()) request.reject(error);
        pending.clear(); worker.terminate(); worker = null;
      };
    }
    return worker;
  }

  function rpc(action, payload, transfer = []) {
    return new Promise((resolve, reject) => {
      const id = ++nextID;
      pending.set(id, {resolve, reject});
      try {getWorker().postMessage({id, action, payload}, transfer);}
      catch (error) {pending.delete(id); reject(error);}
    });
  }

  function clear() {
    if (worker) worker.terminate();
    worker = null; currentUpload = null; activeJob = null; loadingProgress = null;
    for (const request of pending.values()) request.reject(new Error("工作資料已清除，請重新選擇 Excel。"));
    pending.clear(); jobs.clear();
    for (const url of urls) URL.revokeObjectURL(url);
    urls.clear();
  }

  window.ReportTransport = {
    mode: "static",
    async upload(file, onProgress) {
      loadingProgress = onProgress;
      try {
        onProgress("正在讀取電腦上的 Excel…");
        const bytes = await file.arrayBuffer();
        const response = await rpc("upload", {id: crypto.randomUUID(), name: file.name, size: file.size, bytes}, [bytes]);
        currentUpload = response.value;
        return currentUpload;
      } finally {loadingProgress = null;}
    },
    async api(path, payload) {
      if (path === "api/state") {
        if (!manifestPromise) manifestPromise = fetch(assetURL("site-manifest.json")).then((response) => {
          if (!response.ok) throw new Error("網站設定無法載入，請重新整理頁面。");
          return response.json();
        });
        const manifest = await manifestPromise;
        const now = new Date();
        return {year: now.getFullYear(), month: now.getMonth() + 1, rules: manifest.rules, interface: "static", revision: manifest.revision};
      }
      if (path === "api/jobs") {
        if (!currentUpload || payload.upload_id !== currentUpload.id) throw new Error("請先選擇 INPUT Excel。");
        if (activeJob?.status === "running") throw new Error("正在產生報表，請稍候。");
        for (const url of urls) URL.revokeObjectURL(url);
        urls.clear(); jobs.clear();
        const id = crypto.randomUUID();
        const filename = payload.filename.toLowerCase().endsWith(".xlsx") ? payload.filename : payload.filename + ".xlsx";
        const job = {id, filename, status: "running", progress: "正在瀏覽器內計算月報…"};
        activeJob = job; jobs.set(id, job);
        rpc("generate", {...payload, id}).then((message) => {
          const blob = new Blob([message.bytes], {type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});
          job.download = URL.createObjectURL(blob); urls.add(job.download);
          job.result = message.value.result; job.preview = message.value.preview;
          job.progress = "OUTPUT Excel 已完成。"; job.status = "success";
        }).catch((error) => {job.error = error.message; job.status = "error";});
        return {id};
      }
      if (path.startsWith("api/jobs/")) {
        const job = jobs.get(path.slice("api/jobs/".length));
        if (!job) throw new Error("找不到報表工作，請重新產生。");
        return job;
      }
      if (path === "api/shutdown") {
        clear(); return {cleared: true};
      }
      throw new Error("不支援的網頁操作。");
    },
  };
})();
