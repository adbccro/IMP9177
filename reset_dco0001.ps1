# reset_dco0001.ps1
# Full reset of DCO-0001 to Draft state for clean lifecycle testing
# Removes: signatures, routing history, training completions, Official zone files,
#          metadata stamps, DCO phase reset
# Run from SPFx root after connecting to SharePoint

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

$SITE = "https://adbccro.sharepoint.com/sites/IMP9177"
Write-Host "`n[1/8] Connecting..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SITE -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
Write-Host "Connected." -ForegroundColor Green

# ── Step 1: Reset DCO-0001 phase to Draft ────────────────────────────────────
Write-Host "`n[2/8] Resetting DCO-0001 to Draft..." -ForegroundColor Cyan
$dco = Get-PnPListItem -List "QMS_DCOs" -Id 1
Set-PnPListItem -List "QMS_DCOs" -Identity 1 -Values @{
    DCO_Phase         = "Draft"
    DCO_EffectiveDate = $null
} | Out-Null
Write-Host "  DCO-0001 → Draft" -ForegroundColor Green

# ── Step 2: Reset approval records ───────────────────────────────────────────
Write-Host "`n[3/8] Resetting approval records..." -ForegroundColor Cyan
$apprs = Get-PnPListItem -List "QMS_DCOApprovals" -PageSize 500 | Where-Object { $_["Appr_DCOID"] -eq "DCO-0001" }
foreach ($a in $apprs) {
    Set-PnPListItem -List "QMS_DCOApprovals" -Identity ([int]$a.Id) -Values @{
        Appr_Status     = "Waiting"
        Appr_SigID      = ""
        Appr_SignedDate = $null
    } | Out-Null
    Write-Host "  Reset: $($a["Appr_Name"]) → Waiting" -ForegroundColor Gray
}

# ── Step 3: Delete all routing history for DCO-0001 ──────────────────────────
Write-Host "`n[4/8] Deleting routing history..." -ForegroundColor Cyan
$hist = Get-PnPListItem -List "QMS_RoutingHistory" -PageSize 500 | Where-Object { $_["RH_DCOID"] -eq "DCO-0001" }
foreach ($h in $hist) {
    Remove-PnPListItem -List "QMS_RoutingHistory" -Identity ([int]$h.Id) -Force
    Write-Host "  Deleted: $($h["Title"])" -ForegroundColor Gray
}
Write-Host "  $($hist.Count) routing history records deleted" -ForegroundColor Green

# ── Step 4: Delete training completions ──────────────────────────────────────
Write-Host "`n[5/8] Deleting training completions (Cindy Dong)..." -ForegroundColor Cyan
$training = Get-PnPListItem -List "QMS_TrainingCompletions" -PageSize 500 | Where-Object { $_["TC_EmpID"] -eq "Cindy Dong" }
foreach ($t in $training) {
    Remove-PnPListItem -List "QMS_TrainingCompletions" -Identity ([int]$t.Id) -Force
    Write-Host "  Deleted: $($t["Title"])" -ForegroundColor Gray
}
Write-Host "  $($training.Count) training records deleted" -ForegroundColor Green

# ── Step 5: Delete Official zone documents (DCO-0001 set) ────────────────────
Write-Host "`n[6/8] Removing Official zone documents..." -ForegroundColor Cyan
$officialFiles = @(
    "Shared Documents/Official/QMS/Documents/QM-001_Quality_Manual_RevA.docx",
    "Shared Documents/Official/QMS/Documents/SOP-QMS-001_RevA_Management_Responsibility.docx",
    "Shared Documents/Official/QMS/Documents/SOP-QMS-002_RevA_Document_Control.docx",
    "Shared Documents/Official/QMS/Documents/SOP-QMS-003_RevA_Change_Control.docx",
    "Shared Documents/Official/QMS/Documents/SOP-PRD-108_RevA.docx",
    "Shared Documents/Official/QMS/Documents/SOP-PRD-432_RevA.docx",
    "Shared Documents/Official/QMS/Documents/SOP-FRS-549_RevA.docx",
    "Shared Documents/Official/QMS/Documents/DCO-0001_Completion_Report_RevA.pdf",
    "Shared Documents/Official/QMS/Forms/FM-001_Master_Document_Log_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-002_Change_Request_Form_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-003_Document_Change_Order_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-027_QU_QS_Designation_Record_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-030_Finished_Product_Spec_Sheet_RevA.docx",
    "Shared Documents/Official/QMS/Change Orders/DCO-0001_Completion_Report_RevA.pdf"
)
foreach ($f in $officialFiles) {
    try {
        Remove-PnPFile -ServerRelativeUrl "/sites/IMP9177/$f" -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed: $($f.Split('/')[-1])" -ForegroundColor Gray
    } catch {
        Write-Host "  Not found (ok): $($f.Split('/')[-1])" -ForegroundColor DarkGray
    }
}

# ── Step 6: Remove any generated PDFs from Official zones ────────────────────
Write-Host "`n[7/8] Removing generated PDFs from Official zones..." -ForegroundColor Cyan
$pdfDocs = @("QM-001","SOP-QMS-001","SOP-QMS-002","SOP-QMS-003","FM-001","FM-002","FM-003","FM-027","FM-030")
foreach ($docId in $pdfDocs) {
    $pdfPath = "Shared Documents/Official/QMS/Documents/$docId`_RevA.pdf"
    $pdfPathForms = "Shared Documents/Official/QMS/Forms/$docId`_RevA.pdf"
    try { Remove-PnPFile -ServerRelativeUrl "/sites/IMP9177/$pdfPath" -Force -ErrorAction SilentlyContinue } catch {}
    try { Remove-PnPFile -ServerRelativeUrl "/sites/IMP9177/$pdfPathForms" -Force -ErrorAction SilentlyContinue } catch {}
}
Write-Host "  PDF cleanup done" -ForegroundColor Green

# ── Step 7: Clear QMS metadata from Published zone files ─────────────────────
Write-Host "`n[8/8] Clearing QMS metadata stamps from Published zone..." -ForegroundColor Cyan
$publishedFiles = @(
    "Shared Documents/Published/QMS/Documents/SOP-QMS-001_RevA_Management_Responsibility.docx",
    "Shared Documents/Published/QMS/Documents/SOP-QMS-002_RevA_Document_Control.docx",
    "Shared Documents/Published/QMS/Documents/SOP-QMS-003_RevA_Change_Control.docx",
    "Shared Documents/Published/QMS/Forms/FM-001_Master_Document_Log_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-002_Change_Request_Form_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-003_Document_Change_Order_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-027_QU_QS_Designation_Record_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-030_Finished_Product_Spec_Sheet_RevA.docx",
    "Shared Documents/Published/QMS/Quality Manual/QM-001_Quality_Manual_RevA.docx"
)
foreach ($f in $publishedFiles) {
    try {
        $item = Get-PnPFile -Url "/sites/IMP9177/$f" -AsListItem -ErrorAction SilentlyContinue
        if ($item) {
            Set-PnPListItem -List "Shared Documents" -Identity $item.Id -Values @{
                QMS_Revision      = ""
                QMS_DocID         = ""
                QMS_Status        = "Published"
                QMS_EffectiveDate = $null
            } | Out-Null
            Write-Host "  Cleared metadata: $($f.Split('/')[-1])" -ForegroundColor Gray
        }
    } catch { Write-Host "  Skip: $($f.Split('/')[-1])" -ForegroundColor DarkGray }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "`n════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " DCO-0001 RESET COMPLETE" -ForegroundColor Green
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " DCO-0001 Phase    : Draft"
Write-Host " Approvals         : All reset to Waiting"
Write-Host " Routing History   : Cleared"
Write-Host " Training Records  : Cleared"
Write-Host " Official Zone     : Files removed"
Write-Host " Metadata Stamps   : Cleared from Published"
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " Ready for clean lifecycle test." -ForegroundColor Green
