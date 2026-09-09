# Rebuild the Cut off / Open gate dashboard from a new daily CUT file and push it live.
#
#   Right-click this file  ->  "Run with PowerShell"
#   or:  powershell -ExecutionPolicy Bypass -File update.ps1 [path\to\new-CUT.xls]
#
# With no argument it picks the newest *CUT*.xls in this folder.

$ErrorActionPreference = "Stop"
Set-Location -Path $PSScriptRoot

# 1. find the source .xls
if ($args.Count -ge 1) {
    $xls = $args[0]
} else {
    $xls = Get-ChildItem -File -Filter *.xls |
           Where-Object { $_.Name -match 'CUT' } |
           Sort-Object LastWriteTime -Descending |
           Select-Object -First 1 -ExpandProperty FullName
}
if (-not $xls -or -not (Test-Path $xls)) {
    Write-Host "No CUT .xls file found. Pass one as an argument." -ForegroundColor Red
    exit 1
}
Write-Host "Source file : $xls" -ForegroundColor Cyan

# 2. rebuild both HTML files
python build_dashboard.py "$xls"
if ($LASTEXITCODE -ne 0) { Write-Host "Build failed." -ForegroundColor Red; exit 1 }

# 3. commit + push
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm"
git add -A
git commit -m "Update dashboard from $([System.IO.Path]::GetFileName($xls)) ($stamp)"
if ($LASTEXITCODE -ne 0) { Write-Host "Nothing to commit (no changes)." -ForegroundColor Yellow; exit 0 }
git push origin main

Write-Host ""
Write-Host "Pushed. GitHub Pages will refresh in ~1 minute:" -ForegroundColor Green
Write-Host "  https://sirichai1265.github.io/heung-a-cutoff/" -ForegroundColor Green
