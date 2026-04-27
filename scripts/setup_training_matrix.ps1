# setup_training_matrix.ps1
# Seeds QMS_TrainingMatrix with one row per (employee, document) pair.
# Run after: Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

$DocIds = @(
    "QM-001","SOP-QMS-001","SOP-QMS-002","SOP-QMS-003",
    "SOP-SUP-001","SOP-SUP-002","SOP-FS-001","SOP-FS-002",
    "SOP-FS-003","SOP-FS-004","SOP-PC-001","SOP-PRD-108",
    "SOP-PRD-432","SOP-FRS-549","SOP-RCL-321","FPS-001",
    "FM-001","FM-002","FM-003","FM-004","FM-005","FM-006",
    "FM-007","FM-027","FM-030","FM-ALG"
)

Write-Host "Loading employees from QMS_Employees..." -ForegroundColor Cyan
$employees = Get-PnPListItem -List "QMS_Employees" -Fields "Id","Title","Emp_Email","Emp_Status" |
    Where-Object { $_["Emp_Status"] -ne "Inactive" }

Write-Host "Loading existing matrix rows..." -ForegroundColor Cyan
$existing = Get-PnPListItem -List "QMS_TrainingMatrix" -Fields "Id","TM_EmpID","TM_DocID"
$existingSet = [System.Collections.Generic.HashSet[string]]::new()
foreach ($row in $existing) {
    $existingSet.Add("$($row["TM_EmpID"])|$($row["TM_DocID"])") | Out-Null
}

$added = 0
$skipped = 0

foreach ($emp in $employees) {
    $empId = $emp["Title"]
    foreach ($docId in $DocIds) {
        $key = "$empId|$docId"
        if ($existingSet.Contains($key)) {
            $skipped++
            continue
        }
        try {
            Add-PnPListItem -List "QMS_TrainingMatrix" -Values @{
                "Title"       = "$empId-$docId"
                "TM_EmpID"    = $empId
                "TM_DocID"    = $docId
                "TM_Required" = $true
            } | Out-Null
            $existingSet.Add($key) | Out-Null
            $added++
            Write-Host "  [ADD] $empId — $docId" -ForegroundColor Green
        } catch {
            Write-Host "  [ERR] $empId — $docId : $($_.Exception.Message)" -ForegroundColor Red
        }
    }
}

Write-Host "`nDone. Added: $added  Skipped (already exist): $skipped" -ForegroundColor Cyan
