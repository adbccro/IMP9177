# cleanup_duplicate_approvers.ps1
# Reads QMS_Approvers, finds duplicate rows (same Appr_Name + Appr_DocType),
# keeps the active one (or lower ID if both active), deletes duplicates.
# Run after connecting: Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

$SiteUrl  = "https://adbccro.sharepoint.com/sites/IMP9177"
$ListName = "QMS_Approvers"

Write-Host "Connecting to SharePoint..." -ForegroundColor Cyan

try {
    Connect-PnPOnline -Url $SiteUrl -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
} catch {
    Write-Host "Already connected or connection failed: $($_.Exception.Message)" -ForegroundColor Yellow
}

Write-Host "`nLoading all $ListName items..." -ForegroundColor Cyan

$all = Get-PnPListItem -List $ListName -Fields "Id","Title","Appr_Name","Appr_DocType","Appr_Active","Appr_Role","Approver_Email"

Write-Host "Total items loaded: $($all.Count)" -ForegroundColor White

# Group by Appr_Name + Appr_DocType
$groups = @{}
foreach ($item in $all) {
    $name    = $item["Appr_Name"]
    $doctype = $item["Appr_DocType"]
    if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($doctype)) { continue }
    $key = "$name|$doctype"
    if (-not $groups.ContainsKey($key)) { $groups[$key] = @() }
    $groups[$key] += $item
}

$duplicateGroups = $groups.GetEnumerator() | Where-Object { $_.Value.Count -gt 1 }

if (-not $duplicateGroups) {
    Write-Host "`nNo duplicates found. QMS_Approvers is clean." -ForegroundColor Green
    exit 0
}

Write-Host "`nFound $(@($duplicateGroups).Count) duplicate group(s):`n" -ForegroundColor Yellow

$toDelete = @()

foreach ($group in $duplicateGroups) {
    $key   = $group.Key
    $items = $group.Value | Sort-Object { $_.Id }

    Write-Host "Group: $key" -ForegroundColor Cyan
    foreach ($item in $items) {
        $active = $item["Appr_Active"]
        Write-Host "  ID=$($item.Id)  Active=$active  Role=$($item['Appr_Role'])" -ForegroundColor White
    }

    # Determine which to keep:
    # 1. Prefer Appr_Active = true
    # 2. If both active (or both inactive), keep lower ID
    $activeItems   = $items | Where-Object { $_["Appr_Active"] -eq $true }
    $inactiveItems = $items | Where-Object { $_["Appr_Active"] -ne $true }

    $keep   = $null
    $others = @()

    if ($activeItems.Count -ge 1) {
        # Keep the active one with lowest ID
        $keep   = ($activeItems | Sort-Object { $_.Id })[0]
        $others = $items | Where-Object { $_.Id -ne $keep.Id }
    } else {
        # All inactive — keep lowest ID
        $keep   = ($items | Sort-Object { $_.Id })[0]
        $others = $items | Where-Object { $_.Id -ne $keep.Id }
    }

    Write-Host "  --> KEEP ID=$($keep.Id)" -ForegroundColor Green
    foreach ($o in $others) {
        Write-Host "  --> DELETE ID=$($o.Id)" -ForegroundColor Red
        $toDelete += $o
    }
    Write-Host ""
}

if ($toDelete.Count -eq 0) {
    Write-Host "Nothing to delete after analysis." -ForegroundColor Green
    exit 0
}

Write-Host "Items to delete: $($toDelete.Count)" -ForegroundColor Yellow
$confirm = Read-Host "Proceed with deletion? (yes/no)"
if ($confirm -ne "yes") {
    Write-Host "Aborted." -ForegroundColor Yellow
    exit 0
}

$deleted = 0
foreach ($item in $toDelete) {
    try {
        $name    = $item["Appr_Name"]
        $doctype = $item["Appr_DocType"]
        Remove-PnPListItem -List $ListName -Identity $item.Id -Force
        Write-Host "Deleted ID=$($item.Id)  ($name | $doctype)" -ForegroundColor Green
        $deleted++
    } catch {
        Write-Host "ERROR deleting ID=$($item.Id): $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`nDone. Deleted $deleted item(s)." -ForegroundColor Cyan

# Final report
Write-Host "`nFinal state of $ListName (active approvers only):" -ForegroundColor Cyan
$finalActive = Get-PnPListItem -List $ListName -Query "<View><Query><Where><Eq><FieldRef Name='Appr_Active'/><Value Type='Boolean'>1</Value></Eq></Where></Query></View>"
foreach ($f in $finalActive | Sort-Object { $_["Appr_Name"] }) {
    Write-Host "  ID=$($f.Id)  $($f['Appr_Name'])  DocType=$($f['Appr_DocType'])  Role=$($f['Appr_Role'])" -ForegroundColor White
}
