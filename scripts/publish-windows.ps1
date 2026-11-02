$ErrorActionPreference = 'Stop'
Push-Location (Join-Path $PSScriptRoot '..')
try {
    dotnet test tests/ExcelInsights.Tests/ExcelInsights.Tests.csproj -c Release
    if ($LASTEXITCODE -ne 0) { throw 'Tests failed.' }
    dotnet publish src/ExcelInsights.App/ExcelInsights.App.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o artifacts/windows
    if ($LASTEXITCODE -ne 0) { throw 'Publish failed.' }
    Write-Host 'Ready: artifacts/windows/ExcelInsights.App.exe'
} finally { Pop-Location }
