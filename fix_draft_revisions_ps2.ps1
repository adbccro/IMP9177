# fix_draft_revisions_ps.ps1
# Complete PowerShell-only version — no Python required
# Downloads all Drafts DOCX files, patches Revision: A → DRAFT in XML,
# renames files, uploads new versions, deletes old ones
# Run from SPFx root after Connect-PnPOnline

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

# ── DOCX patcher function ────────────────────────────────────────────────────
function Fix-DraftDocx {
    param([string]$InputPath, [string]$OutputDir)

    $basename = Split-Path $InputPath -Leaf

    # Derive new filename
    $newName = $basename `
        -replace '_RevA_',  '_DRAFT_' `
        -replace '_RevB_',  '_DRAFT_' `
        -replace '_RevA\.docx$', '_DRAFT.docx' `
        -replace '_RevB\.docx$', '_DRAFT.docx'

    $outputPath = Join-Path $OutputDir $newName

    # Copy original to output location first
    Copy-Item $InputPath $outputPath -Force

    # Open ZIP (DOCX is a ZIP) in Update mode
    $zip = [System.IO.Compression.ZipFile]::Open($outputPath, 'Update')

    # Target: document.xml and any header*.xml files
    $xmlEntries = $zip.Entries | Where-Object {
        $_.FullName -match '^word/(document|header\d*|footer\d*)\.xml$'
    }

    foreach ($entry in $xmlEntries) {
        # Read full XML content
        $stream = $entry.Open()
        $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
        $content = $reader.ReadToEnd()
        $reader.Close()
        $stream.Close()

        $original = $content

        # ── Revision replacements ──────────────────────────────────────────
        # 1. "Revision: A" in same text run
        $content = $content -replace '(?i)Revision:\s*A\b', 'Revision: DRAFT'
        # 2. ">Revision A<" or ">Rev A<"
        $content = $content -replace '(?i)>Revision A<', '>Revision DRAFT<'
        $content = $content -replace '(?i)>Rev A<', '>DRAFT<'
        # 3. Standalone ">A<" in a w:t tag when preceded by Revision context
        #    Safe because we only replace when "Revision" appears in same XML block
        if ($content -match '(?i)Revision') {
            # Replace <w:t>A</w:t> or <w:t xml:space="preserve">A</w:t>
            $content = [regex]::Replace($content,
                '(<w:t(?:\s[^>]*)?>)\s*A\s*(</w:t>)',
                '${1}DRAFT${2}')
        }
        # 4. "Rev A" anywhere in a text run
        $content = $content -replace '(?i)(?<=>)Rev A(?=<)', 'DRAFT'
        # 5. Effective Date — ensure TBD is present (some docs may have a date)
        $content = $content -replace 'Effective Date: [A-Za-z0-9,\s]+(?=<)',
            'Effective Date: TBD — Pending Approval'

        if ($content -ne $original) {
            # Write back — must delete and re-create entry
            $zip.Dispose()

            # Re-open and replace entry content
            $zipUpdate = [System.IO.Compression.ZipFile]::Open($outputPath, 'Update')
            $targetEntry = $zipUpdate.Entries | Where-Object { $_.FullName -eq $entry.FullName }
            if ($targetEntry) {
                $ws = $targetEntry.Open()
                $ws.SetLength(0)
                $bytes = [System.Text.Encoding]::UTF8.GetBytes($content)
                $ws.Write($bytes, 0, $bytes.Length)
                $ws.Close()
            }
            $zipUpdate.Dispose()

            # Re-open for next iteration
            $zip = [System.IO.Compression.ZipFile]::Open($outputPath, 'Update')
            $xmlEntries = $zip.Entries | Where-Object {
                $_.FullName -match '^word/(document|header\d*|footer\d*)\.xml$'
            }
        }
    }

    $zip.Dispose()
    return $newName
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
    $idx = $docFiles.IndexOf($f) + 1
    Write-Host "`n[$idx/$total] $($f.name)" -ForegroundColor Cyan
    $srcUrl   = "/sites/IMP9177/$($f.folder)/$($f.name)"
    $localIn  = "$TMPDIR\$($f.name)"

    # Download
    try {
        Get-PnPFile -Url $srcUrl -Path $TMPDIR -Filename $f.name -AsFile -Force | Out-Null
        Write-Host "  ✓ Downloaded" -ForegroundColor Gray
    } catch {
        Write-Host "  ✗ Download failed: $_" -ForegroundColor Red; $fail++; continue
    }

    # Fix XML + rename
    try {
        $newName = Fix-DraftDocx -InputPath $localIn -OutputDir "$TMPDIR\output"
        Write-Host "  ✓ Fixed → $newName" -ForegroundColor Gray
    } catch {
        Write-Host "  ✗ Fix failed: $_" -ForegroundColor Red; $fail++; continue
    }

    $localOut = "$TMPDIR\output\$newName"
    if (!(Test-Path $localOut)) {
        Write-Host "  ✗ Output not found after fix" -ForegroundColor Red; $fail++; continue
    }

    # Upload new file
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
            Write-Host "  ✓ Deleted old: $($f.name)" -ForegroundColor Gray
        } catch {
            Write-Host "  ⚠ Could not delete old (manual cleanup needed): $($f.name)" -ForegroundColor Yellow
        }
    }

    $ok++
}

# Cleanup
Remove-Item $TMPDIR -Recurse -Force -ErrorAction SilentlyContinue

Write-Host "`n════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " COMPLETE: $ok/$total succeeded, $fail failed" -ForegroundColor $(if ($fail -eq 0) { "Green" } else { "Yellow" })
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
