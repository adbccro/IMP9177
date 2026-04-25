# fix_draft_revisions_v2.ps1
# Patches all 27 Draft DOCX files: Revision A → DRAFT in XML + renames files
# Pure PowerShell, no Python needed
# Run from SPFx root

Set-StrictMode -Off
$ErrorActionPreference = "Continue"
Add-Type -AssemblyName System.IO.Compression.FileSystem

$SITE   = "https://adbccro.sharepoint.com/sites/IMP9177"
$TMPDIR = "$env:TEMP\imp9177_draft_fix_$(Get-Random)"

Write-Host "`n[1] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SITE -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
Write-Host "Connected." -ForegroundColor Green

New-Item -ItemType Directory -Path $TMPDIR -Force | Out-Null
New-Item -ItemType Directory -Path "$TMPDIR\output" -Force | Out-Null

# ── Patcher: reads all entries to memory, patches, writes fresh ZIP ──────────
function Fix-DraftDocx {
    param([string]$InputPath, [string]$OutputDir)

    $basename = Split-Path $InputPath -Leaf
    $newName  = $basename `
        -replace '_RevA_',       '_DRAFT_' `
        -replace '_RevB_',       '_DRAFT_' `
        -replace '_RevA\.docx$', '_DRAFT.docx' `
        -replace '_RevB\.docx$', '_DRAFT.docx'
    $outputPath = Join-Path $OutputDir $newName

    # Step 1: read all ZIP entries into memory
    $entries = @{}
    $zipIn   = [System.IO.Compression.ZipFile]::OpenRead($InputPath)
    foreach ($entry in $zipIn.Entries) {
        $ms = New-Object System.IO.MemoryStream
        $entry.Open().CopyTo($ms)
        $entries[$entry.FullName] = @{
            bytes = $ms.ToArray()
            isXml = $entry.FullName -match '^word/(document|header\d*|footer\d*)\.xml$'
        }
        $ms.Dispose()
    }
    $zipIn.Dispose()

    # Step 2: patch XML entries in memory
    $patched = 0
    foreach ($key in @($entries.Keys)) {
        if ($entries[$key].isXml) {
            $content  = [System.Text.Encoding]::UTF8.GetString($entries[$key].bytes)
            $original = $content

            # Core replacements
            $content = $content -replace '(?i)Revision:\s*A\b',  'Revision: DRAFT'
            $content = $content -replace '(?i)>Revision A<',     '>Revision DRAFT<'
            $content = $content -replace '(?i)>Rev A<',          '>DRAFT<'
            # Split-run case: <w:t>A</w:t> near "Revision"
            if ($content -match '(?i)Revision') {
                $content = [regex]::Replace($content,
                    '(<w:t(?:\s[^>]*)?>)\s*A\s*(</w:t>)',
                    '${1}DRAFT${2}')
            }

            if ($content -ne $original) {
                $entries[$key].bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
                $patched++
            }
        }
    }

    # Step 3: write fresh ZIP with patched content
    $zipOut = [System.IO.Compression.ZipFile]::Open($outputPath, 'Create')
    foreach ($key in $entries.Keys) {
        $outEntry  = $zipOut.CreateEntry($key, [System.IO.Compression.CompressionLevel]::Optimal)
        $outStream = $outEntry.Open()
        $outStream.Write($entries[$key].bytes, 0, $entries[$key].bytes.Length)
        $outStream.Close()
    }
    $zipOut.Dispose()

    return @{ name = $newName; patched = $patched }
}

# ── File inventory ───────────────────────────────────────────────────────────
$docFiles = @(
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="QM-001_Quality_Manual_RevA.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-QMS-001_RevA_Management_Responsibility.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-QMS-002_RevA_Document_Control.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-QMS-003_RevA_Change_Control.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-SUP-001_RevA_Supplier_Qualification_FINAL.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-SUP-002_RevA_Receiving_Inspection_FINAL.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-FS-001_RevA_Allergen_Control_FINAL.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-FS-002_RevA_Equipment_Cleaning_FINAL.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-FS-003_RevA_Facility_Sanitation_FINAL.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-FS-004_RevA_Environmental_Monitoring_FINAL.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-PC-001_RevA_Pest_Sighting_Response.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-PRD-108_RevA.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-PRD-432_RevA.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-FRS-549_RevA.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="SOP-RCL-321_RevB.docx" },
    @{ folder="Shared Documents/QMS/Documents/Drafts"; name="FPS-001_Lychee_VD3_Gummy_Spec_RevA.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-001_Master_Document_Log_RevA.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-002_Change_Request_Form_RevA.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-003_Document_Change_Order_RevA.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-004_RevA_Approved_Supplier_List_TEMPLATE.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-005_RevA_Receiving_Log_TEMPLATE.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-006_RevA_RawMaterial_Spec_Sheet_TEMPLATE.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-007_RevA_Material_Hold_Label_TEMPLATE.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-008_Supplier_CoA_Requirements_Checklist_RevA.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-027_QU_QS_Designation_Record_RevA.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-030_Finished_Product_Spec_Sheet_RevA.docx" },
    @{ folder="Shared Documents/QMS/Forms/Drafts"; name="FM-ALG_RevA_Allergen_Status_Record_TEMPLATE.docx" }
)

$ok = 0; $fail = 0; $total = $docFiles.Count

foreach ($f in $docFiles) {
    $idx    = [array]::IndexOf($docFiles, $f) + 1
    $srcUrl = "/sites/IMP9177/$($f.folder)/$($f.name)"
    $localIn = "$TMPDIR\$($f.name)"

    Write-Host "`n[$idx/$total] $($f.name)" -ForegroundColor Cyan

    # Download
    try {
        Get-PnPFile -Url $srcUrl -Path $TMPDIR -Filename $f.name -AsFile -Force | Out-Null
        Write-Host "  ✓ Downloaded" -ForegroundColor Gray
    } catch {
        Write-Host "  ✗ Download failed: $_" -ForegroundColor Red; $fail++; continue
    }

    # Fix
    try {
        $res     = Fix-DraftDocx -InputPath $localIn -OutputDir "$TMPDIR\output"
        $newName = $res.name
        Write-Host "  ✓ Fixed → $newName ($($res.patched) XML blocks patched)" -ForegroundColor Gray
    } catch {
        Write-Host "  ✗ Fix failed: $_" -ForegroundColor Red; $fail++; continue
    }

    $localOut = "$TMPDIR\output\$newName"
    if (!(Test-Path $localOut)) {
        Write-Host "  ✗ Output file missing after fix" -ForegroundColor Red; $fail++; continue
    }

    # Upload
    try {
        Add-PnPFile -Path $localOut -Folder $f.folder | Out-Null
        Write-Host "  ✓ Uploaded: $newName" -ForegroundColor Green
    } catch {
        Write-Host "  ✗ Upload failed: $_" -ForegroundColor Red; $fail++; continue
    }

    # Delete old file if renamed
    if ($newName -ne $f.name) {
        try {
            Remove-PnPFile -ServerRelativeUrl $srcUrl -Force
            Write-Host "  ✓ Deleted: $($f.name)" -ForegroundColor Gray
        } catch {
            Write-Host "  ⚠ Manual delete needed: $($f.name)" -ForegroundColor Yellow
        }
    }

    $ok++
}

# Cleanup
Remove-Item $TMPDIR -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " COMPLETE: $ok/$total succeeded, $fail failed" -ForegroundColor $(if ($fail -eq 0) {"Green"} else {"Yellow"})
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
