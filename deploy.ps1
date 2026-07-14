$DEPLOYMENT_ID = "AKfycbxqg2xBlm77Tim8o9ZzVGCwCdTaSmErZwjUSS6hUt4CCK1DhECwNz0ulARcYovHJm08xA"

Write-Host "Pushing code..." -ForegroundColor Cyan
clasp push --force
if (-not $?) { Write-Host "Push failed." -ForegroundColor Red; exit 1 }

Write-Host "Deploying new version..." -ForegroundColor Cyan
clasp deploy --deploymentId $DEPLOYMENT_ID
if (-not $?) { Write-Host "Deploy failed." -ForegroundColor Red; exit 1 }

Write-Host "Done." -ForegroundColor Green
