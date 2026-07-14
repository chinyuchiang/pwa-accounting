cd "$PSScriptRoot"

Write-Host "Pushing code to Apps Script..." -ForegroundColor Cyan
clasp push --force
if (-not $?) { Write-Host "Push failed." -ForegroundColor Red; exit 1 }

Write-Host ""
Write-Host "Code updated!" -ForegroundColor Green
Write-Host ""
Write-Host "如需更新線上版本，請到 Apps Script 執行：" -ForegroundColor Yellow
Write-Host "  部署 -> 管理部署 -> 鉛筆 -> 版本選「建立新版本」-> 部署" -ForegroundColor Yellow
