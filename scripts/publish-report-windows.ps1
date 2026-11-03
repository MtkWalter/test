param([string]$Python = "", [switch]$SkipTests)
$ErrorActionPreference = 'Stop'
if ($env:OS -ne 'Windows_NT') { throw 'Build MonthlyReport.exe on Windows.' }
$repo = Split-Path $PSScriptRoot -Parent
Push-Location $repo
try {
    if (-not $Python) { $Python = Join-Path $repo '.venv/Scripts/python.exe' }
    if (-not (Test-Path $Python)) { throw 'Run build-report.bat or uv sync --locked --extra dev first.' }
    if (-not $SkipTests) {
        & $Python -m pytest -q
        if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
    }
    & $Python -m PyInstaller --noconfirm --clean --onefile --windowed --name MonthlyReport --paths src --hidden-import openpyxl --exclude-module pandas --exclude-module numpy --exclude-module matplotlib --exclude-module tkinter --distpath artifacts/monthly-report-windows --workpath build/monthly-report --specpath build report_main.py
    if ($LASTEXITCODE -ne 0) { throw 'MonthlyReport PyInstaller build failed.' }
    $output = Join-Path $repo 'artifacts/monthly-report-windows'
    $samples = Join-Path $output 'samples'
    New-Item -ItemType Directory -Path $samples -Force | Out-Null
    $sample = Join-Path $samples 'report-demo-input.xlsx'
    $workbook = Join-Path $samples 'report-demo-output.xlsx'
    $report = Join-Path $repo 'artifacts/report-smoke.json'
    foreach ($path in @($report, $workbook)) {
        if (Test-Path $path) { Remove-Item $path }
    }
    & $Python -m excel_insights.report_sample $sample
    if ($LASTEXITCODE -ne 0) { throw 'Could not create report sample.' }
    $arguments = '--smoke-test "{0}" --output "{1}" --smoke-output "{2}" --year 2026 --month 9' -f $sample, $workbook, $report
    $hadPlatform = Test-Path Env:QT_QPA_PLATFORM
    $previousPlatform = $env:QT_QPA_PLATFORM
    try {
        $env:QT_QPA_PLATFORM = 'offscreen'
        $process = Start-Process -FilePath (Join-Path $output 'MonthlyReport.exe') -ArgumentList $arguments -PassThru
        $null = $process.Handle
        if (-not $process.WaitForExit(120000)) {
            $process.Kill()
            throw 'MonthlyReport executable smoke test timed out.'
        }
        if ($process.ExitCode -ne 0) { throw 'MonthlyReport executable smoke test failed.' }
    } finally {
        if ($hadPlatform) { $env:QT_QPA_PLATFORM = $previousPlatform }
        else { Remove-Item Env:QT_QPA_PLATFORM -ErrorAction SilentlyContinue }
    }
    if (-not (Test-Path $report)) { throw 'Missing executable smoke report.' }
    $result = Get-Content $report -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($result.act_rows -ne 20 -or $result.ori_rows -ne 20 -or $result.error) {
        throw 'Frozen report did not read the expected sample rows.'
    }
    & $Python scripts/verify-monthly-report.py $workbook
    if ($LASTEXITCODE -ne 0) { throw 'Frozen workbook validation failed.' }
    Copy-Item docs/monthly-report.md (Join-Path $output 'README.md') -Force
    Copy-Item config/monthly-report-rules.json (Join-Path $output 'monthly-report-rules.json') -Force
    $archive = Join-Path $repo 'artifacts/MonthlyReport-Windows-x64.zip'
    Compress-Archive -Path "$output/*" -DestinationPath $archive -Force
    $checksum = (Get-FileHash $archive -Algorithm SHA256).Hash.ToLowerInvariant()
    "$checksum  MonthlyReport-Windows-x64.zip" | Set-Content artifacts/Report-SHA256SUMS.txt -Encoding ascii
    Write-Host "Ready: $archive"
} catch {
    $details = $_.Exception.Message
    $smoke = Join-Path $repo 'artifacts/report-smoke.json'
    if (Test-Path $smoke) { $details += "`n" + (Get-Content $smoke -Raw -Encoding UTF8) }
    $escaped = $details.Replace('%', '%25').Replace("`r", '%0D').Replace("`n", '%0A')
    Write-Output "::error::$escaped"
    throw
} finally { Pop-Location }
