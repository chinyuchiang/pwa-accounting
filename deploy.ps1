$DEPLOYMENT_ID = "AKfycbyxb7p_WQHddMlU-S5czzhH8wJEjQ1ISXmwZkKjDxAJLJXbxqAfAb1-TiEAvLxYyXo-Jw"

Write-Host "Pushing code..." -ForegroundColor Cyan
clasp push --force
if (-not $?) { Write-Host "Push failed." -ForegroundColor Red; exit 1 }

Write-Host "Deploying new version..." -ForegroundColor Cyan
clasp deploy --deploymentId $DEPLOYMENT_ID
if (-not $?) { Write-Host "Deploy failed." -ForegroundColor Red; exit 1 }

Write-Host "Done." -ForegroundColor Green
