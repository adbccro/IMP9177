# ============================================================
# IMP9177 PM_DocumentDeliverables Seed
# Run while connected: Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -UseWebLogin
# ============================================================

$ErrorActionPreference = "Continue"
Write-Host "Seeding PM_DocumentDeliverables..." -ForegroundColor Cyan

$docs = @(
    # W1 -- 15 documents
    @{Title="PM-QM-001";PMDocID="PM-QM-001";PMDocNumber="QM-001";PMDocWeek="W1";PMDocTitle="Quality Manual";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-SOP-QMS-001";PMDocID="PM-SOP-QMS-001";PMDocNumber="SOP-QMS-001";PMDocWeek="W1";PMDocTitle="Management Responsibility SOP";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-SOP-DC-001";PMDocID="PM-SOP-DC-001";PMDocNumber="SOP-DC-001";PMDocWeek="W1";PMDocTitle="Document Control SOP";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-SOP-TR-001";PMDocID="PM-SOP-TR-001";PMDocNumber="SOP-TR-001";PMDocWeek="W1";PMDocTitle="Training and Competency SOP";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-SOP-CAPA-001";PMDocID="PM-SOP-CAPA-001";PMDocNumber="SOP-CAPA-001";PMDocWeek="W1";PMDocTitle="CAPA SOP";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-SOP-INT-001";PMDocID="PM-SOP-INT-001";PMDocNumber="SOP-INT-001";PMDocWeek="W1";PMDocTitle="Internal Audit SOP";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-SOP-CHG-001";PMDocID="PM-SOP-CHG-001";PMDocNumber="SOP-CHG-001";PMDocWeek="W1";PMDocTitle="Change Control SOP";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-SOP-COMP-001";PMDocID="PM-SOP-COMP-001";PMDocNumber="SOP-COMP-001";PMDocWeek="W1";PMDocTitle="Complaint Handling SOP";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-SOP-REC-001";PMDocID="PM-SOP-REC-001";PMDocNumber="SOP-REC-001";PMDocWeek="W1";PMDocTitle="Records Management SOP";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-FM-001";PMDocID="PM-FM-001";PMDocNumber="FM-001";PMDocWeek="W1";PMDocTitle="Document Change Order Form";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-FM-002";PMDocID="PM-FM-002";PMDocNumber="FM-002";PMDocWeek="W1";PMDocTitle="Training Record Form";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-FM-003";PMDocID="PM-FM-003";PMDocNumber="FM-003";PMDocWeek="W1";PMDocTitle="CAPA Form";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-FM-004";PMDocID="PM-FM-004";PMDocNumber="FM-004";PMDocWeek="W1";PMDocTitle="Internal Audit Checklist";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-FM-027";PMDocID="PM-FM-027";PMDocNumber="FM-027";PMDocWeek="W1";PMDocTitle="Management Review Minutes";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    @{Title="PM-DCO-0001";PMDocID="PM-DCO-0001";PMDocNumber="DCO-0001";PMDocWeek="W1";PMDocTitle="Document Change Order -- W1 Package";PMDocStatus="PAST DUE -- Sign-Off Overdue"},
    # W2 -- 12 documents
    @{Title="PM-SOP-SUP-001";PMDocID="PM-SOP-SUP-001";PMDocNumber="SOP-SUP-001";PMDocWeek="W2";PMDocTitle="Supplier Qualification SOP";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-SOP-PC-001";PMDocID="PM-SOP-PC-001";PMDocNumber="SOP-PC-001";PMDocWeek="W2";PMDocTitle="Pest Control SOP";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-SOP-ALG-001";PMDocID="PM-SOP-ALG-001";PMDocNumber="SOP-ALG-001";PMDocWeek="W2";PMDocTitle="Allergen Control SOP";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-SOP-SAN-001";PMDocID="PM-SOP-SAN-001";PMDocNumber="SOP-SAN-001";PMDocWeek="W2";PMDocTitle="Sanitation and Hygiene SOP";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-SOP-LAB-001";PMDocID="PM-SOP-LAB-001";PMDocNumber="SOP-LAB-001";PMDocWeek="W2";PMDocTitle="Laboratory Controls SOP";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-SOP-REJ-001";PMDocID="PM-SOP-REJ-001";PMDocNumber="SOP-REJ-001";PMDocWeek="W2";PMDocTitle="Nonconforming Material SOP";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-SOP-HARPC-001";PMDocID="PM-SOP-HARPC-001";PMDocNumber="SOP-HARPC-001";PMDocWeek="W2";PMDocTitle="HARPC / Food Safety Plan SOP";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-FM-005";PMDocID="PM-FM-005";PMDocNumber="FM-005";PMDocWeek="W2";PMDocTitle="Supplier Qualification Checklist";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-FM-007";PMDocID="PM-FM-007";PMDocNumber="FM-007";PMDocWeek="W2";PMDocTitle="Allergen Risk Assessment Form";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-FM-008";PMDocID="PM-FM-008";PMDocNumber="FM-008";PMDocWeek="W2";PMDocTitle="Supplier CoA Requirements Checklist";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-FM-009";PMDocID="PM-FM-009";PMDocNumber="FM-009";PMDocWeek="W2";PMDocTitle="Nonconforming Material Report";PMDocStatus="Delivered -- Pending Feedback"},
    @{Title="PM-DCO-0002";PMDocID="PM-DCO-0002";PMDocNumber="DCO-0002";PMDocWeek="W2";PMDocTitle="Document Change Order -- W2 Package";PMDocStatus="Delivered -- Pending Feedback"}
)

$added = 0; $skipped = 0
foreach ($d in $docs) {
    $existing = Get-PnPListItem -List "PM_DocumentDeliverables" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($d.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "PM_DocumentDeliverables" -Values $d | Out-Null
        Write-Host "  + $($d.PMDocNumber) -- $($d.PMDocTitle)" -ForegroundColor DarkGreen
        $added++
    } else {
        $skipped++
    }
}

Write-Host "`n  Added: $added  |  Skipped (exists): $skipped" -ForegroundColor Cyan
Write-Host "  PM dashboard will now show 27/27 docs delivered counter." -ForegroundColor Green
