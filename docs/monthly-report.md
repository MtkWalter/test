# Monthly Report — 本機瀏覽器報表工作台

這是獨立的 Python 本機程式，Windows 執行檔名稱為 **MonthlyReport.exe**，啟動後以預設瀏覽器操作。原有的 ExcelInsights 程式仍可照常使用。計算規則沿用先前確認的單月、AS 金額來源、BI 客戶與 Remark 排序，不需要 Microsoft Excel。

## 使用 Windows EXE

1. 解開 `MonthlyReport-Windows-x64.zip`，執行 `MonthlyReport.exe`，會自動開啟本機瀏覽器頁面。
2. 在頁面選擇或拖曳 INPUT `.xlsx`／`.xlsm`，再選擇來源工作表。
3. 選擇**西元年、月份（單月）**與年度標籤，例如 `2026 / 09 / FY27`。
4. 視需要展開「客群與 Tech Node 分類」調整、載入或儲存設定。
5. 輸入下載檔名，點「產生 OUTPUT」。網頁會顯示進度、來源筆數、資料提醒及報表預覽；可展開完整 Remark。
6. 點「下載 OUTPUT Excel」，儲存位置由瀏覽器下載設定決定。
7. 完成後點右上角「關閉程式」，結束本機服務並清除暫存；已下載的結果會保留。僅關閉分頁不會結束背景程式。

Windows EXE 包含 Python、必要套件與全部網頁資源，不需另外安裝 Python、啟動 Node 或連線下載網頁資源。請用 Edge／Chrome／Firefox 等現代瀏覽器。輸入檔不會被覆寫；修改來源或分類後請重新產生 OUTPUT。結果是當次計算的快照，不依賴 Excel 重新計算公式。

### Excel 在哪裡處理？

EXE 只在 `127.0.0.1` 啟動本機服務，自動選擇可用連接埠。頁面中的選檔／上傳，是把檔案傳給**同一台電腦上的 Python 程式**，不是外部網站。原始檔的工作副本、計算與暫存 OUTPUT 都在本機暫存資料夾；正常以「關閉程式」結束時清除。GitHub 只負責下載程式與更新，不接收使用者的 Excel。頁面與程式不需要網際網路即可運作。

輸出包含三張工作表：

- `monthly pivot`：依範例排列的 A–L 報表、門檻顏色與 Remark。
- `差異明細`：所有 END USER／P/N 的 ACT、BUDGET 與差額，**單位為原始元**，便於對帳。
- `報表資訊`：月份、年度、來源筆數、使用的分類規則與資料提醒。

範例資料 `samples/report-demo-input.xlsx` 是**合成測試資料**，不是使用者真實銷售數字。選擇 2026 年 9 月、FY27、工作表「銷售明細」。對應範例 OUTPUT 的 N20：TSMC 實績 120M／預算 100M，Non-T 實績 80M／預算 100M，兩邊都需要 Remark。

## 輸入欄位

依欄名辨識，位置可移動；下列欄字母對應先前提供的 INPUT。標題須在前 20 列內。截圖重建檔的第 2 列欄字母會略過。有多張符合欄位的工作表時，使用者必須選擇其中一張，避免重複加總。

| 欄位 | 欄名 | 用途 |
| --- | --- | --- |
| AS | sales-NTD | **最終金額來源，原始元** |
| BD | budget / act | `FY27 act` 實績、`FY27 ori` 原始預算；不採用 `rev` |
| BF | TSMC/non TSMC | TSMC／Non-T 客群 |
| BM | Tech Node | 報表節點分類 |
| BI | END USER | **Remark 唯一的客戶辨識欄位** |
| BC | P/N | 品號明細 |
| BB | supplier | 辨識 Service；可選欄位 |
| BP | YM(calender) | 年月，例如 `202609` |
| BO、BE | Year(calender)、Month(calender) | 年份、月份；可代替 BP |

BP 與 BO／BE 同時有值時必須一致；兩者皆無時可讀取「日期」欄。年度標籤預設依 4 月至隔年 3 月、以結束年命名：2026/09 → FY27。可手動修改年度標籤，分類設定也可更改會計年度起始月份。

大型資料採串流彙總，不受原統計工作台的 100,000 列／1,000,000 儲存格／512 欄限制；欄名可位於第 512 欄之後。只讀取選定來源工作表的必要欄位，其餘像 `monthly pivot(CS)` 的報表工作表不會納入明細計算。僅有格式的遠端欄位不會造成「超過 512 欄」錯誤。實際速度仍取決於檔案大小與電腦資源。

**資料完整性：**所選月份的實績 AS 缺值或含無效數字時停止產生，並顯示來源列號。真正零值請填 `0`。原始預算 AS 缺值時，該客群／節點的預算與達成率留白並列出提醒，避免把不完整預算當成已知金額；沒有分配預算明細的個別客戶／品號，在差異分析中按零預算比較。AS 來源公式若無已儲存計算結果，會要求先用 Excel／LibreOffice 重新計算並儲存。BF 空白、非服務資料的 BM 空白或年月衝突也會指出錯誤。

BI／BC 空白時以「未填 END USER／P/N」保留分組並提示補齊，不以 C 客戶簡稱或 BJ CUSTOMER 代替。沒有該月份的任何 act／ori 資料時停止產生；只找到其中一個版本則在報表資訊提醒核對來源完整性。

## 已確認的計算規則

所有版本先篩選指定**單月**，依報表節點、客群彙總 AS 原始金額。

| OUTPUT | 邏輯 |
| --- | --- |
| A Tech Node | 依 BM 與分類對照 |
| B TSMC act | TSMC 的當月實績 |
| C TSMC ori | TSMC 的當月原始預算 |
| D Non-T act | Non-T 的當月實績 |
| E Non-T ori | Non-T 的當月原始預算 |
| F Total act | B＋D；**Service／ADJ 相同** |
| G Total ori | C＋E；**Service／ADJ 相同** |
| H TSMC 差額 | B－C，使用原始金額 |
| I TSMC 達成率 | B÷C；預算為零、缺值或負數時空白 |
| J Non-T 差額 | D－E，使用原始金額 |
| K Non-T 達成率 | D÷E；預算為零、缺值或負數時空白 |
| L Remark | 依下列門檻、客戶與品號排序產生 |

主表所有金額換算成 **AS 加總÷1,000,000，以百萬元顯示整數**；只有 **Remark 金額保留一位小數，附 M**。I／K 顯示整數百分比。**先使用完整精度計算，最後才套用顯示格式**；正負差額及達成率不從畫面已四捨五入的金額重算。

原始預算為正數時，即使主表顯示成 0，仍以原始值計算達成率。空白達成率不視為 0%，不觸發警示或必填 Remark。

## 警示與 Remark

- I／K **≥115%**：對應 H／I 或 J／K 黃色。
- I／K **≤85%**：對應 H／I 或 J／K 藍色。
- 包含 115%、85% 本身，以**未四捨五入的比率**判斷。
- 只有 I 觸發：只需解釋 TSMC；只有 K 觸發：只需解釋 Non-T；兩邊都觸發：分別解釋。
- 所需客群內，**BI END USER 的所有客戶全部列出**。
- 各客戶差額＝ACT－BUDGET。與**該 Tech Node、該客群總差額同方向**的客戶優先，各方向內按差額絕對值遞減。反方向客戶也全部保留。
- **BC P/N 的所有品號全部列出**，按**所屬客戶差額方向**優先，再按差額絕對值遞減。
- 零差額列最後；同差額按名稱排序，使重跑結果一致。無需 Remark 的列留白。
- Remark 太長時用接續列保存全文，避免 Excel 單格長度限制截斷客戶或品號；完整數字另見「差異明細」。

例如：

```text
TSMC
END_A +40.0M：PL_A1 +35.0M；PL_A2 +5.0M
END_C +10.0M：PL_C +10.0M
END_B -30.0M：PL_B_DOWN -50.0M；PL_B_UP +20.0M
Non-T
NON_T_DOWN -50.0M：PL_D -50.0M
NON_T_UP +30.0M：PL_E +30.0M
```

## 尚未由使用者確認的分類，採可調整預設

- BF `tsmc`（忽略大小寫與空白）歸 TSMC；**其餘非空 BF 值歸 Non-T，包含 FI**。可改為明確的 Non-T 值清單；不在清單的值會要求修正，不會悄悄略過。
- BM `N2`、`A16` → `N2/A16`；`N3`、`N3E` → `N3/N3E`。
- **其他 BM 名稱保留原名**，未預設未知的 Reclaim／AP 等分類對照。
- BB 為 `Service`、BM 空白時獨立列 `Service`，不自行分攤至製程分類；BF 仍決定 TSMC／Non-T。
- 一般範例節點依模板順序排列；新節點也會列出，沒有資料的模板列留白。

可在網頁的「客群與 Tech Node 分類」修改、儲存與載入。隨包附的 `monthly-report-rules.json` 是可載入的預設；請在頁面載入需要的版本。CLI 用 `--rules` 指定。這些分類會逐次寫入 OUTPUT 的「報表資訊」。

## Python 開發與命令列

在儲存庫根目錄，以 Python 3.12 與 uv 0.12.23 執行：

```sh
uv sync --locked --extra dev
uv run --no-sync python -m excel_insights.monthly_web
```

Windows 也可先執行既有 `setup.bat`，再執行新增的 `run-report.bat`。CLI 不需要顯示器：

```sh
uv run --no-sync python -m excel_insights.monthly_report \
  --input input.xlsx --output output.xlsx --year 2026 --month 9 \
  --fiscal-year FY27 --sheet 重建資料
```

`--sheet` 在只有一張符合來源工作表時可省略。CLI 已有 OUTPUT 時預設拒絕覆寫，須明確加 `--overwrite`；瀏覽器下載的同名檔案則依瀏覽器設定處理。`--summary-json result.json` 可儲存當次結果或錯誤，失敗時 exit code 為 1。獨立 EXE 亦可透過上述參數執行批次模式，以 `--summary-json` 讀取結果。

開發時可加 `--no-browser` 只啟動服務、不自動開啟瀏覽器，終端會列印本機位址；`--port` 可指定本機連接埠。服務固定綁定 `127.0.0.1`，不提供外網監聽選項。

雲端家目錄唯讀時，命令前設定 `XDG_CACHE_HOME=/workspace/.cache`、`UV_CACHE_DIR=/workspace/.cache/uv`；沒有顯示器的測試加 `QT_QPA_PLATFORM=offscreen`。

```sh
python -m excel_insights.report_sample artifacts/report-demo-input.xlsx
QT_QPA_PLATFORM=offscreen python -m pytest -q
```

實際瀏覽器測試另使用可選的 `browser-test` extra：`uv sync --locked --extra dev --extra browser-test`，再以 `playwright install chromium` 安裝測試瀏覽器。Linux 若已有 `/usr/bin/chromium` 可直接使用；Windows CI 使用已安裝的 Edge（`REPORT_BROWSER_CHANNEL=msedge`）。這些測試工具不會包入使用者的 EXE。

## 產生 Windows EXE

**Windows 上**安裝 Python 3.12 後，雙擊 `build-report.bat`。它在 `build/report-tools` 建立工具環境，安裝固定 uv，依 lockfile 準備 `.venv`，執行測試，再打包並啟動真正的 EXE 做範例轉換。若組織禁止 PowerShell 腳本，請使用符合組織政策的方式或 GitHub Actions。

已安裝依賴時也可在 Windows PowerShell 執行：

```powershell
./scripts/publish-report-windows.ps1
```

成功才會產生 `artifacts/MonthlyReport-Windows-x64.zip` 與 `Report-SHA256SUMS.txt`；內含 `MonthlyReport.exe`、使用說明、分類設定、合成 INPUT／OUTPUT。打包驗證會啟動真正的 EXE，檢查打包網頁資源，再透過本機 HTTP 選檔、產生、下載 Excel，檢查金額、比率、雙客群 Remark、門檻與排序。CI 另用真正的 Edge 驗證完整網頁操作。瀏覽器版 EXE 不需 Qt 桌面執行環境。程式尚未簽章。

現有 GitHub Actions **Python Windows build and test** 已加入瀏覽器版 EXE 建置；成功後下載 **MonthlyReport-windows-x64** artifact。Linux 可執行相同本機網頁及產生 Excel；**Windows EXE 的建置與驗證在 Windows 執行**。
