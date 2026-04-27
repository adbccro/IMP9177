# cleanup_duplicate_dco.ps1
# Removes the older duplicate DCO-0002 record from QMS_DCOs.
# Keeps the item with the HIGHER ID (most recent).
# Run after connecting: Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177"
$ListName = "QMS_DCOs"
$TargetTitle = "DCO-0002"

Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan

try {
    Connect-PnPOnline -Url $SiteUrl -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
} catch {
    Write-Host "Already connected or connection failed: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`nQuerying $ListName for items with Title = '$TargetTitle'..." -ForegroundColor Cyan

$items = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$TargetTitle</Value></Eq></Where></Query></View>"

if ($null -eq $items -or $items.Count -eq 0) {
    Write-Host "No items found with Title = '$TargetTitle'. Nothing to do." -ForegroundColor Green
    exit 0
}

Write-Host "Found $($items.Count) item(s) with Title = '$TargetTitle':" -ForegroundColor Yellow

foreach ($item in $items) {
    Write-Host "  ID=$($item.Id)  Phase=$($item['DCO_Phase'])  Title=$($item['Title'])" -ForegroundColor White
}

if ($items.Count -eq 1) {
    Write-Host "`nOnly one DCO-0002 exists — no duplicate to remove." -ForegroundColor Green
    exit 0
}

# Sort by ID descending — keep the highest ID, delete the rest
$sorted = $items | Sort-Object { $_.Id } -Descending
$keepItem = $sorted[0]
$deleteItems = $sorted | Select-Object -Skip 1

Write-Host "`nKeeping:  ID=$($keepItem.Id)" -ForegroundColor Green
foreach ($del in $deleteItems) {
    Write-Host "Deleting: ID=$($del.Id)" -ForegroundColor Red
}

$confirm = Read-Host "`nProceed with deletion? (yes/no)"
if ($confirm -ne "yes") {
    Write-Host "Aborted." -ForegroundColor Yellow
    exit 0
}

foreach ($del in $deleteItems) {
    try {
        Remove-PnPListItem -List $ListName -Identity $del.Id -Force
        Write-Host "Hard-deleted item ID=$($del.Id)" -ForegroundColor Green
    } catch {
        Write-Host "ERROR deleting ID=$($del.Id): $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Verify
Write-Host "`nVerifying..." -ForegroundColor Cyan
$remaining = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$TargetTitle</Value></Eq></Where></Query></View>"
Write-Host "Remaining '$TargetTitle' items: $($remaining.Count)" -ForegroundColor $(if ($remaining.Count -eq 1) { "Green" } else { "Red" })
foreach ($r in $remaining) {
    Write-Host "  ID=$($r.Id)  Phase=$($r['DCO_Phase'])" -ForegroundColor White
}
