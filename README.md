# Excel Insights — Python Excel 統計工作台

這是原 C# / WPF 初版的 Python 替代實作（v0.3.0）。使用 **PySide6 + pandas**，提供本機自動判讀、可修改的報表預覽與範本 Excel 匯出，也保留多檔匯入、分組統計與 CSV。資料全程在本機處理，不上傳、不修改來源檔，也不需要安裝 Microsoft Excel。

## 本機自動判讀＋預覽確認

1. 在主畫面匯入本機 Excel。程式在前 50 列尋找熟悉的欄名，例如 `Tech Node`、`sales-NTD`、`budget / act`，可辨識第 4 列才是標題的版面。若判讀不符，在左側指定標題列，清除該檔後重新匯入。
2. 按 **「本機判讀＋報表預覽」**。第 1 頁勾選要合併的工作表，檢查來源欄位對應；相同欄位結構共用設定。無法對應必要欄位的表，預設不勾選。
3. 第 2 頁用下拉選單及點選值，修改 act／ori、TSMC／Non-T／FI 的判斷。可改用客戶分類、END USER、產品、供應商或料號，選擇「等於」或「包含」，不用寫程式。
4. 可調整曆年／會計年度、FY 起始月、曆月、供應商／料號／產品篩選、金額倍率與顯示單位、小數位、差額方向、百分比算法、顏色門檻、Total 範圍與總計列。
5. 第 3 頁查看數值、黃／藍色、來源筆數及略過原因。**結果不是你要的，就回前兩頁修改。** 每次修改都會取消上一次的確認；重新勾選「我已檢查本次預覽與設定」後才能匯出。
6. 按「匯出範本 Excel」產生 XLSX，附「設定與檢查」工作表，不覆寫來源。輸出是快照，在 Excel 手動改數字不會重算；請在程式調整後重新匯出。
7. 「儲存設定」產生本機 JSON，下次匯入後開啟即可重用。設定含欄名、選取的分類／篩選值及算法，不包含交易列；每次仍可勾選要納入的來源。

**FI 可在第 2 頁自行改判斷欄位、等於／包含及選取值。Remark 可在第 3 頁雙擊該格修改，或選取列後按「編輯選取列 Remark」輸入多行文字。** 手動備註會隨 Excel 匯出，不修改來源；可按「還原選取列 Remark」回到原始／系統產生的備註。修改後必須重新確認才能匯出。手動 Remark 依 Tech Node 在本次視窗保留，改篩選後請重新核對；不存入設定 JSON，關閉前請匯出。切換設定或示範資料前會詢問是否清除手動備註。

**自動判讀是可修改的建議，不會自動理解企業內部的 FI 定義或會計年度。** 無法確定的版本值保持未設定；沒有指定或未匹配的分類留白並提示。預設不限制年度。會計年度採結束年命名，例如 FY27、起始月 4，代表 2026/04/01 至 2027/03/31。

預設差額 `act − ori`、百分比 `act ÷ ori`、黃色 ≥115%、藍色 ≤85%；切換成「差額 ÷ ori」會建議 +15%／−15%，仍可修改。ori 為 0 或缺值時比率留白。Total 預設包含有效明細中的 Service／未分類，也可只取三分類聯集；重疊列不重複計入 Total。總計列比率由總金額重新計算。無效金額、空白 Node、無法判讀期間及其他版本會計數並提示。

這版使用交易明細（長表）轉固定報表，各輸入須含 Node、金額及版本欄位。跨檔 lookup/join、多條件 AND 分類、自由拖曳版型尚未包含。分類為單一欄位的多值「等於／包含」。分類與篩選先列出最多 2,000 個值；篩選欄可直接輸入其他精確值。報表最多 5,000 列。

### 不用真實資料也能試用

未匯入檔案也可開啟報表視窗，按 **「載入示範資料」**。虛構資料只影響此預覽視窗，主畫面的檔案會保留。`samples/` 也有：

- `report-demo-input.xlsx`：第 4 列標題、48 筆虛構明細。
- `report-demo-output.xlsx`：6 個 Tech Node，保留分組表頭、百萬顯示、比率及色彩。
- `report-demo-settings.json`：可開啟的設定，示範 FY27 起始月 4、曆月 9。

範例 2nm（以百萬顯示）：TSMC act=120、ori=100、差額=20、比例=120%；Non-T act=40、ori=50、比例=80%；FI act=21、ori=20；Total act=191、ori=180（含 Service）。**這些不是使用者的實際業績。**

## Windows 開發：先執行，再打包

1. 安裝 **Python 3.12 x64**，勾選「Add python.exe to PATH」。支援 Python 3.11–3.13。
2. 下載或 clone 本儲存庫原始碼，解壓縮到例如 `C:\ExcelInsights`。
3. 第一次雙擊 `setup.bat`，建立隔離環境並安裝套件。
4. 雙擊 `run.bat` 開啟桌面程式。
5. 用 VS Code 或其他編輯器修改 `src/excel_insights/` 的 Python 檔案；關閉程式後重新執行 `run.bat` 即可看到修改。這個流程不必每次重新打包 EXE，也不會自動熱更新正在執行的視窗。

也可在儲存庫根目錄使用 PowerShell：

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -e ".[dev]"
.\.venv\Scripts\python.exe -m excel_insights
```

若要使用完整鎖定的相依版本，可安裝 uv 後執行 `uv sync --locked --extra dev`，再用 `uv run --no-sync python -m excel_insights`。`uv.lock` 也用於雲端與 Windows CI。

## 使用方式

1. 點「匯入 Excel 檔案」一次多選，或把檔案拖到視窗。
2. 左側選取工作表。此一般分析模式一次分析一張表；跨檔合併及指定範本請使用新的報表流程。
3. 選數值欄位，例如「金額」；需要時選分組欄位，例如「地區」。
4. 選篩選欄位並輸入包含文字（不分大小寫）；清除文字以恢復全部。
5. 查看資料筆數、有效數值、空白、非數值、總和、平均、最小值與最大值。
6. 點「匯出統計 CSV」。結果包含檔名、工作表、數值欄位、分組及篩選條件，使用 UTF-8 BOM，可用 Excel 開啟。文字中的公式前綴會加上保護。

匯入在背景執行；損壞檔案會個別顯示錯誤，其他檔案可繼續匯入。匯入期間請等待完成後再關閉視窗。資料不會在關閉程式後自動保留。

### 範例

同時匯入 `samples/` 的兩個檔案，左側應有 3 張工作表。

一月「銷售明細」選金額、不分組：6 筆、4 個有效數值、1 個空白、1 個非數值，總和 **11,100**、平均 **2,775**、最小 **1,200**、最大 **6,000**。

依地區分組，台北為 3 筆、2 個有效數值、1 個空白，總和 **8,400**、平均 **4,200**。地區篩選「台北」時只計算這 3 列。

## 打包 Windows EXE

在 **Windows** 完成上述安裝後，在 PowerShell 執行：

```powershell
./scripts/publish-windows.ps1
```

腳本會先執行測試，使用 PyInstaller 建立 `ExcelInsights.exe`，再實際啟動打包後的程式，以範例 Excel 驗證 Qt 初始化、檔案讀取與統計結果。成功後產生：

- `artifacts/ExcelInsights-Python-Windows-x64.zip`：EXE、使用說明及範例檔。
- `artifacts/Python-SHA256SUMS.txt`：壓縮包校驗碼。
- `artifacts/python-smoke.json`：打包後程式的檢查結果。

使用者只需解壓縮並執行 `ExcelInsights.exe`，不必安裝 Python。程式尚未簽章，首次啟動可能因單檔解壓縮而較慢。

若 PowerShell 的組織政策不允許執行腳本，請使用符合組織政策的方式執行，或透過 GitHub Actions 取得打包檔案。

推送後，GitHub Actions 的 **Python Windows build and test** 會在 Windows runner 測試與打包。執行成功後，登入 GitHub，在該次執行頁面下方下載 **ExcelInsights-Python-windows-x64** artifact，再解開裡面的程式壓縮包。這與舊版的 `ExcelInsights-windows-x64` artifact 不同。

## 格式與計算規則

- `.xlsx` 使用 openpyxl、`.xls` 使用 xlrd、`.xlsb` 使用 pyxlsb。不支援加密檔。
- 標題列自動判讀或手動指定；陌生版面預設第一列。空白欄名自動命名，重複欄名加編號。全空白資料列略過；字串 `NA` 等仍保留為文字。
- 數值與可解析的數字文字納入計算，例如 `1,234.50`、`-5`、`2e3`；使用小數點 `.` 與千分位 `,`，不依作業系統地區改變。
- 日期、布林、一般文字、貨幣與百分比文字不列入有效數值。原生百分比依儲存比例計算，例如 50% 為 0.5。**pyxlsb 無法可靠辨識日期格式，XLSB 日期可能作為 Excel 序號數值呈現；需要日期辨識時請先存成 XLSX。**
- 空白與非數值分開計數；沒有有效數值時總和、平均及極值留空。畫面顯示 2 位小數，CSV 保留計算精度。採浮點運算，不適合要求精確小數的財務結算。
- 公式只讀取檔案已儲存的結果，不重新計算。公式未儲存結果時可能為空白。
- 預覽最多 5,000 列，統計使用所有篩選後資料。每次最多選取 20 檔、每表 100,000 資料列 / 512 欄、每檔掃描及保留最多 1,000,000 儲存格、總載入最多 2,000,000 儲存格。Excel 的空白或格式範圍也可能觸及限制；超限會整檔拒絕，不以截斷資料統計。
- 尚未包含圖表、任意公式編輯、自由版型、安裝精靈與自動更新。

## 測試與專案結構

```powershell
.\.venv\Scripts\python.exe -m pytest -q
```

測試包含真正的 XLSX / XLS 讀取、原始範例計算、型別保留、資源上限、CSV 回讀，以及 Qt 視窗中的背景匯入、篩選、分組、匯出與移除操作。XLSB 引擎有包含在打包中，仍需以實際 XLSB 檔案驗收。

- `src/excel_insights/core.py`：Excel 讀取、篩選、統計與 CSV。
- `src/excel_insights/app.py`：PySide6 視窗、背景匯入與表格模型。
- `tests/`：資料與 Qt 整合測試。
- `scripts/publish-windows.ps1`：Windows 打包與 EXE 自我檢查。
- `scripts/create-samples.py`：僅用 Python 標準函式庫產生範例 Excel。

Linux 可直接開發並執行 PySide6；沒有顯示器的雲端環境可用 `QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -q` 驗證真正的 Qt 介面。Windows EXE 必須在 Windows 打包；離屏測試不代表已完成 Windows 實際桌面視覺驗收。
