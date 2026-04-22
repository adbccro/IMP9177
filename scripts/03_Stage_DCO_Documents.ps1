# =============================================================================
# IMP9177 — QMS PORTAL SHAREPOINT SETUP
# Script 03: Stage DCO-0001 & DCO-0002 Documents for Publish Workflow
# Run AFTER 02_Cleanup_Misplaced_Files.ps1 and all manual steps are done
#
# What this script does:
#   - Verifies all DCO-0001 and DCO-0002 documents exist in their correct
#     Drafts locations before the application routes them
#   - Reports any missing documents so you can upload them
#   - Does NOT move Published documents — the application handles that
# =============================================================================

param(
    [string]$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177",
    [switch]$Upload  = $false   # Future: set true to auto-upload missing stubs
)

$ErrorActionPreference = "Continue"

Write-Host "`n[IMP9177] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
Write-Host "[OK] Connected`n" -ForegroundColor Green

$lib = "Shared Documents"
$pass = 0; $fail = 0; $warn = 0

function Check-File {
    param([string]$FolderPath, [string]$FileName, [string]$DocId, [string]$DCO)
    $fullPath = "$lib/$FolderPath/$FileName"
    try {
        $file = Get-PnPFile -Url $fullPath -AsListItem -ErrorAction Stop
        Write-Host "  [OK]     $DocId — $FileName" -ForegroundColor Green
        $script:pass++
    } catch {
        Write-Host "  [MISSING] $DocId — $FileName" -ForegroundColor Red
        Write-Host "            Expected at: $FolderPath/" -ForegroundColor DarkRed
        Write-Host "            DCO: $DCO — Upload this file before loading the DCO." -ForegroundColor DarkRed
        $script:fail++
    }
}

function Check-Optional {
    param([string]$FolderPath, [string]$FileName, [string]$Note)
    $fullPath = "$lib/$FolderPath/$FileName"
    try {
        $file = Get-PnPFile -Url $fullPath -AsListItem -ErrorAction Stop
        Write-Host "  [OK]     $FileName" -ForegroundColor Green
        $script:pass++
    } catch {
        Write-Host "  [NOTE]   $FileName — NOT FOUND ($Note)" -ForegroundColor DarkYellow
        $script:warn++
    }
}

# =============================================================================
# DCO-0001 DOCUMENT LIST (Week 1 — 15 documents)
# Source: DCO-0001_Document_Change_Order_V1_March30_2026.docx
# All documents should be in QMS/Documents/Drafts or QMS/Forms/Drafts
# =============================================================================
Write-Host "=== DCO-0001 — Week 1 QMS Document Package (15 docs) ===" -ForegroundColor Yellow
Write-Host "Verifying files in QMS/Documents/Drafts and QMS/Forms/Drafts...`n"

# Quality Manual
Check-File "QMS/Documents/Drafts" "QM-001_Quality_Manual_RevA.docx"          "QM-001"  "DCO-0001"

# Core SOPs
Check-File "QMS/Documents/Drafts" "SOP-001_RevA_Document_Control.docx"        "SOP-001" "DCO-0001"
Check-File "QMS/Documents/Drafts" "SOP-002_RevA_CAPA_Procedure.docx"          "SOP-002" "DCO-0001"
Check-File "QMS/Documents/Drafts" "SOP-003_RevA_Supplier_Qualification.docx"  "SOP-003" "DCO-0001"
Check-File "QMS/Documents/Drafts" "SOP-004_RevA_Training_Control.docx"        "SOP-004" "DCO-0001"
Check-File "QMS/Documents/Drafts" "SOP-005_RevA_Management_Responsibility.docx" "SOP-005" "DCO-0001"
Check-File "QMS/Documents/Drafts" "SOP-006_RevA_Internal_Audit.docx"          "SOP-006" "DCO-0001"
Check-File "QMS/Documents/Drafts" "SOP-007_RevA_Nonconforming_Product.docx"   "SOP-007" "DCO-0001"

# Forms (DCO-0001 batch)
Check-File "QMS/Forms/Drafts" "FM-001_RevA_Master_Document_Log.docx"          "FM-001"  "DCO-0001"
Check-File "QMS/Forms/Drafts" "FM-002_RevA_Change_Request_Form.docx"          "FM-002"  "DCO-0001"
Check-File "QMS/Forms/Drafts" "FM-003_RevA_Document_Change_Order.docx"        "FM-003"  "DCO-0001"
Check-File "QMS/Forms/Drafts" "FM-008_RevA_Supplier_CoA_Checklist.docx"       "FM-008"  "DCO-0001"
Check-File "QMS/Forms/Drafts" "FM-025_RevA_RCD_Form.docx"                     "FM-025"  "DCO-0001"
Check-File "QMS/Forms/Drafts" "FM-027_RevA_QU_QS_Designation_Record.docx"    "FM-027"  "DCO-0001"
Check-File "QMS/Forms/Drafts" "FM-030_RevA_Finished_Product_Spec_Sheet.docx"  "FM-030"  "DCO-0001"

Write-Host ""

# =============================================================================
# DCO-0002 DOCUMENT LIST (Week 2 — 12 documents)
# Source: DCO-0002_Phase2A_W2_Document_Change_Order.docx
# =============================================================================
Write-Host "=== DCO-0002 — Week 2 QMS Document Package (12 docs) ===" -ForegroundColor Yellow
Write-Host "Verifying files in QMS/Documents/Drafts and QMS/Forms/Drafts...`n"

# Food Safety SOPs
Check-File "QMS/Documents/Drafts" "SOP-FS-001_RevA_Allergen_Control.docx"           "SOP-FS-001" "DCO-0002"
Check-File "QMS/Documents/Drafts" "SOP-FS-002_RevA_Equipment_Cleaning.docx"         "SOP-FS-002" "DCO-0002"
Check-File "QMS/Documents/Drafts" "SOP-FS-003_RevA_Facility_Sanitation.docx"        "SOP-FS-003" "DCO-0002"
Check-File "QMS/Documents/Drafts" "SOP-FS-004_RevA_Environmental_Monitoring.docx"   "SOP-FS-004" "DCO-0002"
Check-File "QMS/Documents/Drafts" "SOP-PC-001_RevA_Pest_Sighting_Response.docx"     "SOP-PC-001" "DCO-0002"
Check-File "QMS/Documents/Drafts" "SOP-SUP-001_RevA_Supplier_Qualification.docx"    "SOP-SUP-001" "DCO-0002"
Check-File "QMS/Documents/Drafts" "SOP-SUP-002_RevA_Receiving_Inspection.docx"      "SOP-SUP-002" "DCO-0002"

# Forms (DCO-0002 batch)
Check-File "QMS/Forms/Drafts" "FM-004_RevA_Approved_Supplier_List.docx"             "FM-004" "DCO-0002"
Check-File "QMS/Forms/Drafts" "FM-005_RevA_Receiving_Log.docx"                      "FM-005" "DCO-0002"
Check-File "QMS/Forms/Drafts" "FM-006_RevA_RawMaterial_Spec_Sheet.docx"             "FM-006" "DCO-0002"
Check-File "QMS/Forms/Drafts" "FM-007_RevA_Material_Hold_Label.docx"                "FM-007" "DCO-0002"
Check-File "QMS/Forms/Drafts" "FM-ALG_RevA_Allergen_Status_Record.docx"             "FM-ALG" "DCO-0002"

Write-Host ""

# =============================================================================
# DCO FORMS (the actual DCO and CR documents)
# =============================================================================
Write-Host "=== Change Order & Change Request Documents ===" -ForegroundColor Yellow

Check-File "QMS/Change Orders"   "DCO-0001_Document_Change_Order_V1_March30_2026.docx"  "DCO-0001" "NA"
Check-File "QMS/Change Orders"   "DCO-0002_Phase2A_W2_Document_Change_Order.docx"        "DCO-0002" "NA"
Check-File "QMS/Change Requests" "CR-0001_Change_Request_Form_V1_March30_2026.docx"      "CR-0001"  "NA"
Check-File "QMS/Change Requests" "CR-0002_Phase2A_W2_Change_Request.docx"                "CR-0002"  "NA"

Write-Host ""

# =============================================================================
# PUBLISHED ZONE CHECK — What's already been promoted
# =============================================================================
Write-Host "=== Published Zone — Currently Promoted Documents ===" -ForegroundColor Yellow
Write-Host "Checking Published/QMS for already-published files...`n"

Check-Optional "Published/QMS/Quality Manual" "QM-001_Quality_Manual_RevA.docx"           "Already published — good"
Check-Optional "Published/QMS/Forms"          "FM-001_Master_Document_Log_RevA.docx"       "Check if should be in Forms subfolder"
Check-Optional "Published/QMS/Forms"          "FM-002_Change_Request_Form_RevA.docx"       "Already in Published/Forms"
Check-Optional "Published/QMS/Forms"          "FM-003_Document_Change_Order_RevA.docx"     "Already in Published/Forms"
Check-Optional "Published/QMS/Forms"          "FM-027_QU_QS_Designation_Record_RevA.docx"  "Already in Published/Forms"
Check-Optional "Published/QMS/Documents"      "SOP-001_RevA_Document_Control.docx"         "Should be here after PM publish"

Write-Host ""

# =============================================================================
# SUMMARY
# =============================================================================
$total = $pass + $fail + $warn
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " DCO STAGING VERIFICATION SUMMARY" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " Total checked : $total"
Write-Host " Present (OK)  : $pass" -ForegroundColor Green
Write-Host " Missing       : $fail" -ForegroundColor $(if ($fail -gt 0) {"Red"} else {"Green"})
Write-Host " Notes         : $warn" -ForegroundColor DarkYellow
Write-Host ""

if ($fail -gt 0) {
    Write-Host " ACTION REQUIRED: $fail document(s) are missing from their expected Drafts location." -ForegroundColor Red
    Write-Host " Upload each missing file to its expected folder before loading DCOs into the application." -ForegroundColor Red
} else {
    Write-Host " All required documents are present. You may proceed to application deployment." -ForegroundColor Green
}
Write-Host ""
Write-Host " Next script: 04_Validate_Published_Zone.ps1"
Write-Host ""
