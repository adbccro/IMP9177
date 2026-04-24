# reset_dco0002.ps1
# Full reset of DCO-0002 to Draft state for clean lifecycle testing
# DCO-0002 covers: SOP-SUP-001/002, SOP-FS-001/002/003/004, SOP-PC-001,
#                  FM-004/005/006/007/008, FM-ALG
# Run from SPFx root after connecting to SharePoint

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

$SITE = "https://adbccro.sharepoint.com/sites/IMP9177"
Write-Host "`n[1/8] Connecting..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SITE -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
Write-Host "Connected." -ForegroundColor Green

# ── Step 1: Reset DCO-0002 phase to Draft ────────────────────────────────────
Write-Host "`n[2/8] Resetting DCO-0002 to Draft..." -ForegroundColor Cyan
$dco = Get-PnPListItem -List "QMS_DCOs" -PageSize 500 | Where-Object { $_["Title"] -eq "DCO-0002" }
if ($dco) {
    Set-PnPListItem -List "QMS_DCOs" -Identity ([int]$dco.Id) -Values @{
        DCO_Phase         = "Draft"
        DCO_EffectiveDate = $null
    } | Out-Null
    Write-Host "  DCO-0002 → Draft (item $($dco.Id))" -ForegroundColor Green
} else {
    Write-Host "  DCO-0002 not found in QMS_DCOs" -ForegroundColor Yellow
}

# ── Step 2: Reset DCO-0002 approval records ───────────────────────────────────
Write-Host "`n[3/8] Resetting DCO-0002 approval records..." -ForegroundColor Cyan
$apprs = Get-PnPListItem -List "QMS_DCOApprovals" -PageSize 500 | Where-Object { $_["Appr_DCOID"] -eq "DCO-0002" }
foreach ($a in $apprs) {
    Set-PnPListItem -List "QMS_DCOApprovals" -Identity ([int]$a.Id) -Values @{
        Appr_Status     = "Waiting"
        Appr_SigID      = ""
        Appr_SignedDate = $null
    } | Out-Null
    Write-Host "  Reset: $($a["Appr_Name"]) → Waiting" -ForegroundColor Gray
}
Write-Host "  $($apprs.Count) approval records reset" -ForegroundColor Green

# ── Step 3: Delete routing history for DCO-0002 ──────────────────────────────
Write-Host "`n[4/8] Deleting routing history..." -ForegroundColor Cyan
$hist = Get-PnPListItem -List "QMS_RoutingHistory" -PageSize 500 | Where-Object { $_["RH_DCOID"] -eq "DCO-0002" }
foreach ($h in $hist) {
    Remove-PnPListItem -List "QMS_RoutingHistory" -Identity ([int]$h.Id) -Force
    Write-Host "  Deleted: $($h["Title"])" -ForegroundColor Gray
}
Write-Host "  $($hist.Count) routing history records deleted" -ForegroundColor Green

# ── Step 4: Delete training completions tied to DCO-0002 SOPs ────────────────
Write-Host "`n[5/8] Deleting training completions for DCO-0002 SOPs..." -ForegroundColor Cyan
$dco2Sops = @("SOP-SUP-001","SOP-SUP-002","SOP-FS-001","SOP-FS-002","SOP-FS-003","SOP-FS-004","SOP-PC-001")
$training = Get-PnPListItem -List "QMS_TrainingCompletions" -PageSize 500 | Where-Object { 
    $dco2Sops -contains $_["TC_DocID"]
}
foreach ($t in $training) {
    Remove-PnPListItem -List "QMS_TrainingCompletions" -Identity ([int]$t.Id) -Force
    Write-Host "  Deleted: $($t["Title"])" -ForegroundColor Gray
}
Write-Host "  $($training.Count) training records deleted" -ForegroundColor Green

# ── Step 5: Remove Official zone documents (DCO-0002 set) ────────────────────
Write-Host "`n[6/8] Removing Official zone documents..." -ForegroundColor Cyan
$officialFiles = @(
    # Documents
    "Shared Documents/Official/QMS/Documents/SOP-SUP-001_RevA_Supplier_Qualification_FINAL.docx",
    "Shared Documents/Official/QMS/Documents/SOP-SUP-002_RevA_Receiving_Inspection_FINAL.docx",
    "Shared Documents/Official/QMS/Documents/SOP-FS-001_RevA_Allergen_Control_FINAL.docx",
    "Shared Documents/Official/QMS/Documents/SOP-FS-002_RevA_Equipment_Cleaning_FINAL.docx",
    "Shared Documents/Official/QMS/Documents/SOP-FS-003_RevA_Facility_Sanitation_FINAL.docx",
    "Shared Documents/Official/QMS/Documents/SOP-FS-004_RevA_Environmental_Monitoring_FINAL.docx",
    "Shared Documents/Official/QMS/Documents/SOP-PC-001_RevA_Pest_Sighting_Response.docx",
    # Forms
    "Shared Documents/Official/QMS/Forms/FM-004_Supplier_Evaluation_Form_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-005_Ingredient_Approval_Form_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-006_Material_Receipt_Log_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-007_CoA_Review_Checklist_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-008_Supplier_CoA_Requirements_Checklist_RevA.docx",
    "Shared Documents/Official/QMS/Forms/FM-ALG_Allergen_Status_Record_RevA.docx",
    # DCO report
    "Shared Documents/Official/QMS/Change Orders/DCO-0002_Completion_Report_RevA.pdf"
)
foreach ($f in $officialFiles) {
    try {
        Remove-PnPFile -ServerRelativeUrl "/sites/IMP9177/$f" -Force -ErrorAction SilentlyContinue
        Write-Host "  Removed: $($f.Split('/')[-1])" -ForegroundColor Gray
    } catch {
        Write-Host "  Not found (ok): $($f.Split('/')[-1])" -ForegroundColor DarkGray
    }
}

# ── Step 6: Remove generated PDFs from Official zones ────────────────────────
Write-Host "`n[7/8] Removing generated PDFs..." -ForegroundColor Cyan
$pdfDocs = @(
    "SOP-SUP-001","SOP-SUP-002","SOP-FS-001","SOP-FS-002",
    "SOP-FS-003","SOP-FS-004","SOP-PC-001",
    "FM-004","FM-005","FM-006","FM-007","FM-008","FM-ALG"
)
foreach ($docId in $pdfDocs) {
    try { Remove-PnPFile -ServerRelativeUrl "/sites/IMP9177/Shared Documents/Official/QMS/Documents/$docId`_RevA.pdf" -Force -ErrorAction SilentlyContinue } catch {}
    try { Remove-PnPFile -ServerRelativeUrl "/sites/IMP9177/Shared Documents/Official/QMS/Forms/$docId`_RevA.pdf" -Force -ErrorAction SilentlyContinue } catch {}
}
Write-Host "  PDF cleanup done" -ForegroundColor Green

# ── Step 7: Clear QMS metadata from Published zone files ─────────────────────
Write-Host "`n[8/8] Clearing QMS metadata from Published zone..." -ForegroundColor Cyan
$publishedFiles = @(
    "Shared Documents/Published/QMS/Documents/SOP-SUP-001_RevA_Supplier_Qualification_FINAL.docx",
    "Shared Documents/Published/QMS/Documents/SOP-SUP-002_RevA_Receiving_Inspection_FINAL.docx",
    "Shared Documents/Published/QMS/Documents/SOP-FS-001_RevA_Allergen_Control_FINAL.docx",
    "Shared Documents/Published/QMS/Documents/SOP-FS-002_RevA_Equipment_Cleaning_FINAL.docx",
    "Shared Documents/Published/QMS/Documents/SOP-FS-003_RevA_Facility_Sanitation_FINAL.docx",
    "Shared Documents/Published/QMS/Documents/SOP-FS-004_RevA_Environmental_Monitoring_FINAL.docx",
    "Shared Documents/Published/QMS/Documents/SOP-PC-001_RevA_Pest_Sighting_Response.docx",
    "Shared Documents/Published/QMS/Forms/FM-004_Supplier_Evaluation_Form_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-005_Ingredient_Approval_Form_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-006_Material_Receipt_Log_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-007_CoA_Review_Checklist_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-008_Supplier_CoA_Requirements_Checklist_RevA.docx",
    "Shared Documents/Published/QMS/Forms/FM-ALG_Allergen_Status_Record_RevA.docx"
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
            Write-Host "  Cleared: $($f.Split('/')[-1])" -ForegroundColor Gray
        }
    } catch { Write-Host "  Skip: $($f.Split('/')[-1])" -ForegroundColor DarkGray }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "`n════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " DCO-0002 RESET COMPLETE" -ForegroundColor Green
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " DCO-0002 Phase    : Draft"
Write-Host " Approvals         : All reset to Waiting"
Write-Host " Routing History   : Cleared"
Write-Host " Training Records  : Cleared (DCO-0002 SOPs)"
Write-Host " Official Zone     : Files removed"
Write-Host " Metadata Stamps   : Cleared from Published"
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " Ready for clean lifecycle test." -ForegroundColor Green
