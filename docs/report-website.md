# 直接開網址使用的報表網站

這個版本由 GitHub Pages 提供靜態網頁，**不需下載或執行 EXE，也不需安裝 Python**。報表運算使用 Pyodide／WebAssembly，在瀏覽器背景工作執行相同的 Python `monthly_report.py`，保留已確認的計算與 Remark 規則。

## 使用方式

1. 開啟已部署的網站網址。
2. 選擇或拖曳 INPUT `.xlsx`／`.xlsm`，選取來源工作表。
3. 指定西元年、單月、FY 年度標籤；需要時展開分類設定。
4. 點「產生 OUTPUT」，查看進度、數值與完整 Remark。
5. 點「下載 OUTPUT Excel」。完成後可用「清除資料」釋放瀏覽器中的工作資料，已下載的結果會保留。

首次選檔會下載約 13 MB 的瀏覽器 Python 運算核心，請稍候；後續可使用瀏覽器的資源快取。建議使用電腦版的新版 Chrome、Edge 或 Firefox，處理大型資料時保留足夠記憶體。

## Excel 會上傳嗎？

**不會。** 網路只用來下載網頁、Python 核心及程式資源。選擇的 Excel 透過瀏覽器 File API 讀進 Web Worker 的記憶體檔案系統，再在同一個瀏覽器內計算、產生下載檔。沒有 Excel 上傳 API、外部計算伺服器、資料庫、分析追蹤或外部字型。

原始檔不會被網頁修改。Excel 與報表資料不會存入 localStorage、IndexedDB 或 GitHub；「清除資料」會終止運算 Worker、清除工作資料並撤銷舊的下載連結。分類設定可另外下載成 JSON。關閉分頁會結束該分頁的運算。

瀏覽器驗證包含「所有請求都只向網站同來源發出 GET／HEAD」檢查，確保選檔及產生報表不透過 HTTP 傳出 Excel。此版本也測試超過 100,000 列、1,000,000 儲存格的資料；實際速度與可處理大小仍取決於電腦和瀏覽器記憶體。

## 已保留的報表規則

- 單月、AS 最終金額來源、BD act／ori、BF TSMC／Non-T、BM 節點分類。
- BI END USER 作為 Remark 客戶，BC P/N 作為產品。
- F=B+D、G=C+E，Service／ADJ 相同。
- 達成率≥115%／≤85%：分客群標色、產生對應 Remark；兩邊觸發則分別說明。
- 全部客戶與品號列出，同方向優先、差額絕對值遞減。
- 主表百萬元整數；Remark 百萬元一位小數。
- 預算零、缺值、負數時達成率空白；先用完整精度計算，最後格式化。
- 過長 Remark 用接續列保存全文，附完整差異明細與報表資訊。

完整輸入要求與分類預設見 [報表規則](monthly-report.md)。網站與 EXE 直接共用同一份 Python 計算引擎，不另外維護一份 JavaScript 金額算法。

## 開發、測試與發布

在儲存庫根目錄執行：

```sh
uv sync --locked --extra dev --extra browser-test
npm ci --ignore-scripts
uv run --no-sync python scripts/build-report-site.py
```

產物在 `_site/`，包含 HTML、JavaScript、CSS、運算核心與固定版本套件。`python-deps.zip`、`report-engine.zip` 以確定的檔案順序與時間戳建立，並由瀏覽器核對 SHA-256。運算核心版本由 `package-lock.json` 固定，不從第三方 CDN 動態載入使用者程式。

開發預覽可用 `python -m http.server --directory _site`；這只提供靜態檔案，不是 Excel 後端。網站需透過 HTTP／HTTPS 開啟，正式使用者直接開 GitHub Pages 的 HTTPS 網址。

```sh
uv run --no-sync python -m pytest -q tests/test_static_site.py
```

需有 Playwright 瀏覽器；可執行 `playwright install chromium`，或設定 `REPORT_BROWSER_CHANNEL=chrome`／`msedge` 使用已安裝的 Chrome／Edge。`REPORT_SITE_URL` 可改為已部署的網址；`REPORT_EXPECTED_REVISION` 可驗證部署版本。

GitHub Actions **Publish report website** 會建置靜態網站、執行真實瀏覽器與大型 Excel 測試、上傳 `MonthlyReport-Website`，再部署 GitHub Pages，最後對正式網址做選檔、產生及下載驗證。

第一次發布須由儲存庫管理者在 **Settings → Pages → Build and deployment → Source** 選擇 **GitHub Actions**。一般工作流程的 `GITHUB_TOKEN` 可發布已啟用的網站，但不能替管理者第一次啟用 Pages；若尚未啟用，Configure Pages 步驟會明確失敗。啟用後重跑該次工作流程即可，不需提供新的個人權杖。

## 格式錯誤、完整 Raw Data 與 HTML 結果

新增「**完整保留 Raw Data 匯出**」按鈕。此模式會將輸入的**所有工作表、所有月份、所有原始列與欄**保留在 OUTPUT，原有工作表不改名、不篩列、不改值。公式、已儲存的公式結果、樣式、合併儲存格、隱藏工作表與附屬檔案保留。若來源為巨集格式，下載副檔名為 `.xlsm` 以保留原內容。

驗證通過時，另加入正常單月報表、差異明細及報表資訊。驗證失敗時，仍可下載原始資料，另附「匯出說明」及「完整錯誤清單」，明確標示尚未產生統計，不會以 0 取代無效金額。來源已存在相同名稱的工作表時，新增工作表會自動編號，原始頁籤不覆蓋。缺少必要報表欄名的有效活頁簿也能使用此模式。

錯誤清單列出工作表、來源列號、欄位位置、欄名、**完整原始值及原因**，不再限制為前 15 筆；同一列的金額、客群及節點錯誤會分別列出。畫面清單支援上下、左右捲動，每頁最多 500 項，可切頁查看所有項目；下載的 HTML 與 Excel 包含全部項目。Excel 單一儲存格過長時使用接續列，錯誤超過 Excel 單表列數時使用接續工作表。

「**下載結果 HTML**」及「**下載完整清單 HTML**」會產生可直接雙擊、離線開啟的獨立 HTML，不依賴網站、Python 或 EXE。結果 HTML 含畫面報表預覽及全部錯誤資訊；完整原始活頁簿仍在下載的 OUTPUT Excel。報表產生網站需透過 HTTPS 開啟；下載的結果 HTML 則可以直接開檔查看。

為避免 Excel 套件重新存檔造成原始工作表資訊遺失，保留模式直接複製原活頁簿的 ZIP 零件，僅擴充加入報表所需的中繼資料與樣式。輸出檔內另含**與輸入檔位元組完全相同的備份**；取回路徑記錄在「匯出說明」，可使用 ZIP 工具取出。此模式針對可讀取的 Excel 活頁簿；加密、非 Excel 或損毀至無法讀取封裝的檔案仍會明確顯示讀取失敗。

## OUTPUT Excel 與 HTML 版型

主表、網頁預覽及「下載結果 HTML」統一採使用者提供的 OUTPUT 版型：A–L 共 12 欄，保留月份／FY 標題、門檻圖例、客群分組及 act／ori 三層表頭，含 Total 綠色分組標題、藍色與米色交錯列、TSMC／Non-T 各自的黃／藍色差異標示、負數紅字與右側 Remark。欄寬依範例比例排列；窄螢幕可橫向捲動。

HTML 直接帶入 Excel 的表頭與註記，不重新計算預覽數值。下載的 HTML 內嵌版型樣式與完整 Remark，不需要另外下載 CSS；可以離線開啟或列印。網站的 `example-output.html` 使用合成交易產生月報版型預覽，`example-errors.html` 提供完整錯誤清單範例；使用者上傳的原檔與真實數字不會發布到這些公開範例。
