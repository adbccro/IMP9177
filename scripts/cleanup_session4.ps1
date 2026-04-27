# cleanup_session4.ps1
# Run after: Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177"

Write-Host "=== Task 1: Reset DCO-0001 to Draft ===" -ForegroundColor Cyan

# Find DCO-0001 item
$dco = Get-PnPListItem -List "QMS_DCOs" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>DCO-0001</Value></Eq></Where></Query></View>" | Select-Object -First 1
if ($dco) {
    Set-PnPListItem -List "QMS_DCOs" -Identity $dco.Id -Values @{
        "DCO_Phase" = "Draft"
        "DCO_SubmittedDate" = $null
        "DCO_EffectiveDate" = $null
    }
    Write-Host "  [OK] DCO-0001 reset to Draft" -ForegroundColor Green
} else {
    Write-Host "  [WARN] DCO-0001 not found" -ForegroundColor Yellow
}

# Delete all QMS_DCOApprovals for DCO-0001
$approvals = Get-PnPListItem -List "QMS_DCOApprovals" -Query "<View><Query><Where><Eq><FieldRef Name='Appr_DCOID'/><Value Type='Text'>DCO-0001</Value></Eq></Where></Query></View>"
foreach ($item in $approvals) {
    Remove-PnPListItem -List "QMS_DCOApprovals" -Identity $item.Id -Force
    Write-Host "  [DEL] QMS_DCOApprovals item $($item.Id)" -ForegroundColor Yellow
}
Write-Host "  [OK] Deleted $($approvals.Count) approval rows for DCO-0001" -ForegroundColor Green

# Delete all QMS_RoutingHistory for DCO-0001
$history = Get-PnPListItem -List "QMS_RoutingHistory" -Query "<View><Query><Where><Eq><FieldRef Name='RH_DCOID'/><Value Type='Text'>DCO-0001</Value></Eq></Where></Query></View>"
foreach ($item in $history) {
    Remove-PnPListItem -List "QMS_RoutingHistory" -Identity $item.Id -Force
    Write-Host "  [DEL] QMS_RoutingHistory item $($item.Id)" -ForegroundColor Yellow
}
Write-Host "  [OK] Deleted $($history.Count) routing history rows for DCO-0001" -ForegroundColor Green

Write-Host "`n=== Task 2: Remove duplicate employees ===" -ForegroundColor Cyan

$employees = Get-PnPListItem -List "QMS_Employees" -Fields "Id","Title","Emp_FullName"
$groups = $employees | Group-Object { $_["Emp_FullName"] }

foreach ($group in $groups) {
    if ($group.Count -gt 1) {
        $sorted = $group.Group | Sort-Object { $_.Id }
        $keep = $sorted[0]
        $dupes = $sorted | Select-Object -Skip 1
        Write-Host "  Name: $($group.Name) — keeping Id $($keep.Id), deleting $($dupes.Count) duplicate(s)" -ForegroundColor Yellow
        foreach ($dupe in $dupes) {
            Remove-PnPListItem -List "QMS_Employees" -Identity $dupe.Id -Force
            Write-Host "    [DEL] Employee Id $($dupe.Id)" -ForegroundColor Red
        }
    }
}

Write-Host "`nCleanup complete." -ForegroundColor Cyan
