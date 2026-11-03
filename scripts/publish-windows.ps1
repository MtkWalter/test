param([string]$Python = "")
$ErrorActionPreference = 'Stop'
$repo = Split-Path $PSScriptRoot -Parent
Push-Location $repo
try {
    if (-not $Python) { $Python = Join-Path $repo '.venv/Scripts/python.exe' }
    if (-not (Test-Path $Python)) { throw 'Run setup.bat or uv sync --extra dev first.' }
    & $Python -m pytest -q
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
    & $Python -m PyInstaller --noconfirm --clean --onefile --windowed --name ExcelInsights --paths src --hidden-import openpyxl --hidden-import xlrd --hidden-import pyxlsb --distpath artifacts/python-windows --workpath build/pyinstaller --specpath build main.py
    if ($LASTEXITCODE -ne 0) { throw 'PyInstaller build failed.' }

    $output = Join-Path $repo 'artifacts/python-windows'
    $report = Join-Path $repo 'artifacts/python-smoke.json'
    if (Test-Path $report) { Remove-Item $report }
    $sample = Join-Path $repo 'samples/銷售範例-一月.xlsx'
    $arguments = '--smoke-test "{0}" --smoke-output "{1}"' -f $sample, $report
    $hadPlatform = Test-Path Env:QT_QPA_PLATFORM
    $previousPlatform = $env:QT_QPA_PLATFORM
    try {
        $env:QT_QPA_PLATFORM = 'offscreen'
        $process = Start-Process -FilePath (Join-Path $output 'ExcelInsights.exe') -ArgumentList $arguments -PassThru
        $null = $process.Handle
        if (-not $process.WaitForExit(120000)) {
            $process.Kill()
            throw 'Frozen executable smoke test timed out.'
        }
        if ($process.ExitCode -ne 0) { throw 'Frozen executable smoke test failed.' }
    } finally {
        if ($hadPlatform) { $env:QT_QPA_PLATFORM = $previousPlatform }
        else { Remove-Item Env:QT_QPA_PLATFORM -ErrorAction SilentlyContinue }
    }
    if (-not (Test-Path $report)) { throw 'Frozen executable produced no smoke report.' }
    $result = Get-Content $report -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($result.sheets -ne 2 -or $result.preview_rows -ne 6 -or $result.summary[0].'總和' -ne 11100) {
        throw 'Frozen executable statistics do not match the sample.'
    }
    Copy-Item README.md $output -Force
    Copy-Item samples (Join-Path $output 'samples') -Recurse -Force
    $archive = Join-Path $repo 'artifacts/ExcelInsights-Python-Windows-x64.zip'
    Compress-Archive -Path "$output/*" -DestinationPath $archive -Force
    $checksum = (Get-FileHash $archive -Algorithm SHA256).Hash.ToLowerInvariant()
    "$checksum  ExcelInsights-Python-Windows-x64.zip" | Set-Content artifacts/Python-SHA256SUMS.txt -Encoding ascii
    Write-Host "Ready: $archive"
} catch {
    $details = $_.Exception.Message
    $smoke = Join-Path $repo 'artifacts/python-smoke.json'
    if (Test-Path $smoke) { $details += "`n" + (Get-Content $smoke -Raw -Encoding UTF8) }
    $escaped = $details.Replace('%', '%25').Replace("`r", '%0D').Replace("`n", '%0A')
    Write-Output "::error::$escaped"
    throw
} finally { Pop-Location }
