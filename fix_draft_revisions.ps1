# fix_draft_revisions.ps1
# Downloads all Drafts zone DOCX files, applies revision fix (RevA → DRAFT),
# renames files, and re-uploads to SharePoint
# Run from SPFx root after connecting to SharePoint

Set-StrictMode -Off
$ErrorActionPreference = "Continue"

$SITE = "https://adbccro.sharepoint.com/sites/IMP9177"
$PYTHON = "python"  # or "python3" depending on your system
$SCRIPT = "$PSScriptRoot\fix_draft_revision.py"
$TMPDIR = "$env:TEMP\imp9177_draft_fix"

Write-Host "`n[1] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SITE -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
Write-Host "Connected." -ForegroundColor Green

# Create temp working directory
if (Test-Path $TMPDIR) { Remove-Item $TMPDIR -Recurse -Force }
New-Item -ItemType Directory -Path $TMPDIR | Out-Null
New-Item -ItemType Directory -Path "$TMPDIR\output" | Out-Null

# ── File inventory ──────────────────────────────────────────────────────────
$docFiles = @(
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "QM-001_Quality_Manual_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-QMS-001_RevA_Management_Responsibility.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-QMS-002_RevA_Document_Control.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-QMS-003_RevA_Change_Control.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-SUP-001_RevA_Supplier_Qualification_FINAL.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-SUP-002_RevA_Receiving_Inspection_FINAL.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-FS-001_RevA_Allergen_Control_FINAL.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-FS-002_RevA_Equipment_Cleaning_FINAL.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-FS-003_RevA_Facility_Sanitation_FINAL.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-FS-004_RevA_Environmental_Monitoring_FINAL.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-PC-001_RevA_Pest_Sighting_Response.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-PRD-108_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-PRD-432_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-FRS-549_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "SOP-RCL-321_RevB.docx" },
    @{ folder = "Shared Documents/QMS/Documents/Drafts"; name = "FPS-001_Lychee_VD3_Gummy_Spec_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-001_Master_Document_Log_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-002_Change_Request_Form_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-003_Document_Change_Order_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-004_RevA_Approved_Supplier_List_TEMPLATE.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-005_RevA_Receiving_Log_TEMPLATE.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-006_RevA_RawMaterial_Spec_Sheet_TEMPLATE.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-007_RevA_Material_Hold_Label_TEMPLATE.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-008_Supplier_CoA_Requirements_Checklist_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-027_QU_QS_Designation_Record_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-030_Finished_Product_Spec_Sheet_RevA.docx" },
    @{ folder = "Shared Documents/QMS/Forms/Drafts"; name = "FM-ALG_RevA_Allergen_Status_Record_TEMPLATE.docx" }
)

$results = @()
$total = $docFiles.Count
$i = 0

foreach ($f in $docFiles) {
    $i++
    $srcUrl = "/sites/IMP9177/$($f.folder)/$($f.name)"
    $localIn  = "$TMPDIR\$($f.name)"
    $outDir   = "$TMPDIR\output"
    
    Write-Host "`n[$i/$total] $($f.name)" -ForegroundColor Cyan
    
    # Step 1 — Download
    try {
        Get-PnPFile -Url $srcUrl -Path $TMPDIR -Filename $f.name -AsFile -Force | Out-Null
        Write-Host "  ✓ Downloaded" -ForegroundColor Gray
    } catch {
        Write-Host "  ✗ Download failed: $_" -ForegroundColor Red
        continue
    }
    
    # Step 2 — Run Python fix script
    try {
        $pyOut = & $PYTHON $SCRIPT $localIn $outDir 2>&1
        $newName = ($pyOut | Where-Object { $_ -match "^OUTPUT:" }) -replace "OUTPUT:", ""
        if (!$newName) {
            # Fallback: derive name manually
            $newName = $f.name -replace "_RevA_", "_DRAFT_" -replace "_RevB_", "_DRAFT_" -replace "_RevA\.docx$", "_DRAFT.docx" -replace "_RevB\.docx$", "_DRAFT.docx"
        }
        Write-Host "  ✓ Fixed: $newName" -ForegroundColor Gray
    } catch {
        Write-Host "  ✗ Python fix failed: $_" -ForegroundColor Red
        continue
    }
    
    $localOut = "$outDir\$newName"
    if (!(Test-Path $localOut)) {
        Write-Host "  ✗ Output file not found: $localOut" -ForegroundColor Red
        continue
    }
    
    # Step 3 — Upload new file
    try {
        Add-PnPFile -Path $localOut -Folder $f.folder | Out-Null
        Write-Host "  ✓ Uploaded: $newName" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ Upload failed: $_" -ForegroundColor Red
        continue
    }
    
    # Step 4 — Delete old file (only if new name differs)
    if ($newName -ne $f.name) {
        try {
            Remove-PnPFile -ServerRelativeUrl $srcUrl -Force
            Write-Host "  ✓ Deleted old: $($f.name)" -ForegroundColor Gray
        } catch {
            Write-Host "  ⚠ Could not delete old file (may need manual cleanup): $($f.name)" -ForegroundColor Yellow
        }
    }
    
    $results += [PSCustomObject]@{
        OldName = $f.name
        NewName = $newName
        Folder  = $f.folder
        Status  = "OK"
    }
}

# ── Summary ──────────────────────────────────────────────────────────────────
Write-Host "`n════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " DRAFT REVISION FIX COMPLETE" -ForegroundColor Green
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " Files processed: $($results.Count) / $total"
Write-Host ""
$results | Format-Table OldName, NewName, Status -AutoSize

# Cleanup temp
Remove-Item $TMPDIR -Recurse -Force -ErrorAction SilentlyContinue
Write-Host "`nTemp files cleaned up." -ForegroundColor Gray
Write-Host "All draft documents now show Revision: DRAFT in header and filename." -ForegroundColor Green
