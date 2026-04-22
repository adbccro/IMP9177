# =============================================================================
# IMP9177 — QMS PORTAL SHAREPOINT SETUP
# Script 02: Clean Up Misplaced Files & Reconcile Legacy Documents
# Run AFTER 01_Create_Folder_Structure.ps1
#
# What this script does:
#   A. Moves CR-0001/DCO-0001 out of QMS/Records (wrong location)
#   B. Moves FPS-001 product spec out of QMS/Records
#   C. Moves supplier CoA letters out of QMS/Records/CAPA
#   D. Moves QM-001 from Documents root to Documents/Drafts
#   E. Flags legacy SOP files with non-standard naming for review
# =============================================================================

param(
    [string]$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177",
    [switch]$WhatIf  = $false   # Run with -WhatIf to preview without moving
)

$ErrorActionPreference = "Stop"

Write-Host "`n[IMP9177] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
$web = Get-PnPWeb
Write-Host "[OK] Connected to: $($web.Title)`n" -ForegroundColor Green

$lib = "Shared Documents"

function Move-SPFile {
    param([string]$SourcePath, [string]$DestFolder, [string]$Reason)
    $fileName = $SourcePath -split "/" | Select-Object -Last 1
    Write-Host "  MOVE: $fileName" -ForegroundColor Yellow
    Write-Host "    From: $SourcePath"
    Write-Host "    To:   $DestFolder/"
    Write-Host "    Why:  $Reason"
    if (-not $WhatIf) {
        try {
            Move-PnPFile -SourceUrl "$lib/$SourcePath" `
                         -TargetUrl "$lib/$DestFolder/$fileName" `
                         -Force -ErrorAction Stop
            Write-Host "    [DONE]" -ForegroundColor Green
        } catch {
            Write-Warning "    [FAILED] $_"
        }
    } else {
        Write-Host "    [WHATIF - no action taken]" -ForegroundColor DarkGray
    }
    Write-Host ""
}

function Flag-File {
    param([string]$FilePath, [string]$Issue, [string]$Action)
    Write-Host "  FLAG: $FilePath" -ForegroundColor Magenta
    Write-Host "    Issue:  $Issue"
    Write-Host "    Action: $Action`n"
}

# =============================================================================
# A. CR and DCO documents wrongly placed in QMS/Records
#    Records folder should contain completed record instances only
# =============================================================================
Write-Host "=== A. Remove CR/DCO documents from QMS/Records ===" -ForegroundColor Cyan

Move-SPFile `
    -SourcePath "QMS/Records/CR-0001_Change_Request_Form_V1_March30_2026.docx" `
    -DestFolder "QMS/Change Requests" `
    -Reason "Change Requests belong in QMS/Change Requests, not QMS/Records"

Move-SPFile `
    -SourcePath "QMS/Records/DCO-0001_Document_Change_Order_V1_March30_2026.docx" `
    -DestFolder "QMS/Change Orders" `
    -Reason "DCO documents belong in QMS/Change Orders, not QMS/Records (duplicate of file already there)"

# =============================================================================
# B. Product Specification wrongly placed in QMS/Records
# =============================================================================
Write-Host "=== B. Move product spec out of QMS/Records ===" -ForegroundColor Cyan

Move-SPFile `
    -SourcePath "QMS/Records/FPS-001_Lychee_VD3_Gummy_Spec_RevA.docx" `
    -DestFolder "QMS/Documents" `
    -Reason "Product specifications are controlled documents, not records. Place in QMS/Documents until a dedicated Product Specs folder is created."

# =============================================================================
# C. Supplier CoA request letters wrongly placed in QMS/Records/CAPA
#    These are supplier correspondence, not CAPA records
# =============================================================================
Write-Host "=== C. Move supplier CoA letters out of CAPA ===" -ForegroundColor Cyan

Move-SPFile `
    -SourcePath "QMS/Records/CAPA/3H_Letter_AIC01_Ingredion_CornSyrup_CoA_Request.docx" `
    -DestFolder "QMS/Records/Supplier" `
    -Reason "Supplier CoA request letters are supplier correspondence, not CAPA records"

Move-SPFile `
    -SourcePath "QMS/Records/CAPA/3H_Letter_AIC02_PacificPectin_HM100_CoA_Request.docx" `
    -DestFolder "QMS/Records/Supplier" `
    -Reason "Supplier CoA request letters are supplier correspondence, not CAPA records"

Move-SPFile `
    -SourcePath "QMS/Records/CAPA/3H_Letter_AIC03_JeanNiel_Flavor_CoA_Program.docx" `
    -DestFolder "QMS/Records/Supplier" `
    -Reason "Supplier CoA request letters are supplier correspondence, not CAPA records"

# =============================================================================
# D. QM-001 is in Documents root — move to Documents/Drafts
#    It needs to flow through Publish action to Published/QMS/Quality Manual
#    NOTE: A copy also exists in Documents/Drafts/Archive — review which is current
# =============================================================================
Write-Host "=== D. Stage QM-001 correctly in Draft zone ===" -ForegroundColor Cyan

Flag-File `
    -FilePath "QMS/Documents/QM-001_Quality_Manual_RevA.docx" `
    -Issue "QM-001 sits in the Documents root. It should be in Documents/Drafts to be properly staged for publish. A copy also exists in Documents/Arhive (note typo in archive folder name)." `
    -Action "MANUAL: Confirm which copy is current. Move the current version to QMS/Documents/Drafts. Archive the older copy. DO NOT simply run a move — verify the document content first."

# =============================================================================
# E. Flag legacy SOP files with non-IMP9177 naming convention
#    These use old internal naming (SOP-FRS, SOP-PRD, SOP-QMS, SOP-RCL)
#    rather than the IMP9177 convention (SOP-001, SOP-002...)
# =============================================================================
Write-Host "=== E. Legacy SOP naming — requires manual review ===" -ForegroundColor Cyan

$legacySOPs = @(
    @{File="QMS/Documents/SOP-FRS-549_RevA.docx";    Note="Food Safety series — likely maps to SOP-FS-001 through SOP-FS-004. Confirm numbering with document register."},
    @{File="QMS/Documents/SOP-PRD-108_RevA.docx";    Note="Production series — no IMP9177 number assigned yet. Check document register."},
    @{File="QMS/Documents/SOP-PRD-432_RevA.docx";    Note="Production series — no IMP9177 number assigned yet. Check document register."},
    @{File="QMS/Documents/SOP-QMS-001_RevA_Management_Responsibility.docx"; Note="Likely maps to SOP-001 or SOP-002 series. Confirm against DCO-0001 document list."},
    @{File="QMS/Documents/SOP-QMS-002_RevA_Document_Control.docx";          Note="Likely maps to SOP-002 Document Control. Confirm against DCO-0001 document list."},
    @{File="QMS/Documents/SOP-QMS-003_RevA_Change_Control.docx";            Note="Likely maps to SOP-003 or Change Control SOP. Confirm against DCO-0001 document list."},
    @{File="QMS/Documents/SOP-RCL-321_RevB.docx";    Note="Already at Rev B — no Rev A visible. Pre-existing 3H document? Needs Rev A history review. Unusual to issue Rev B as first formal revision."}
)

Write-Host ""
Write-Host "  The following files in QMS/Documents use legacy naming and need manual review:" -ForegroundColor Yellow
foreach ($sop in $legacySOPs) {
    Write-Host ""
    Write-Host "  FILE:  $($sop.File)" -ForegroundColor White
    Write-Host "  NOTE:  $($sop.Note)" -ForegroundColor Gray
}

Write-Host ""
Write-Host "  RECOMMENDED ACTION:" -ForegroundColor Yellow
Write-Host "  1. Open the Master Record Sheet (V12) to see the official IMP9177 document register"
Write-Host "  2. Cross-reference each legacy SOP to an IMP9177 number"
Write-Host "  3. Rename files to match IMP9177 convention (e.g., SOP-QMS-002 → SOP-002_RevA_Document_Control)"
Write-Host "  4. Move renamed files to QMS/Documents/Drafts"
Write-Host ""

# =============================================================================
# F. Fix typo in QMS/Documents/Arhive folder name (missing 'c')
# =============================================================================
Write-Host "=== F. Rename typo folder QMS/Documents/Arhive → Archive ===" -ForegroundColor Cyan

Write-Host "  The folder 'QMS/Documents/Arhive' has a typo (missing 'c')."
Write-Host "  SharePoint does not support folder rename via PnP Move in all versions."
Write-Host "  MANUAL ACTION: Rename 'Arhive' to 'Archive' via SharePoint UI."
Write-Host ""

# =============================================================================
# SUMMARY
# =============================================================================
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " SCRIPT 02 COMPLETE" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host " Items requiring MANUAL action before continuing:"
Write-Host "   1. Confirm which QM-001 copy is current, move to Drafts"
Write-Host "   2. Rename legacy SOP files to IMP9177 naming convention"
Write-Host "   3. Rename 'Arhive' folder to 'Archive' in QMS/Documents"
Write-Host "   4. Review SOP-RCL-321 RevB — confirm Rev A history"
Write-Host ""
Write-Host " Next script: 03_Stage_DCO0001_Documents.ps1"
Write-Host ""
