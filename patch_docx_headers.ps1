# patch_docx_headers.ps1
# IMP9177 — Patch all Draft and Published DOCX files:
#   - Replace internal header "Revision: A" / "Revision: B" → "Revision: DRAFT"
#   - Replace effective date fields → "TBD — Pending Approval"
#   - Rename Published files: _RevA_ / _RevB_ → _DRAFT_ (then delete old)
#
# Run from SPFx root (or anywhere) after Connect-PnPOnline.
# Requires: PnP.PowerShell, Connect-PnPOnline already called.
#
# Usage:
#   Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
#   .\patch_docx_headers.ps1

$SiteUrl  = "https://adbccro.sharepoint.com/sites/IMP9177"
$TempDir  = Join-Path $env:TEMP "IMP9177_DocxPatch"
$LogFile  = Join-Path $PSScriptRoot "patch_docx_headers_log.txt"

# Folder paths → zone label (Draft = headers only; Published = headers + rename)
$Folders = @(
    @{ Path = "Shared Documents/QMS/Documents/Drafts";          Zone = "Draft"     }
    @{ Path = "Shared Documents/QMS/Forms/Drafts";              Zone = "Draft"     }
    @{ Path = "Shared Documents/Published/QMS/Documents";       Zone = "Published" }
    @{ Path = "Shared Documents/Published/QMS/Forms";           Zone = "Published" }
    @{ Path = "Shared Documents/Published/QMS/Quality Manual";  Zone = "Published" }
)

# ── Helpers ─────────────────────────────────────────────────────────────────

function Write-Log($msg) {
    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $msg"
    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

function Get-DraftFilename($name) {
    # Replace _RevA_ or _RevB_ (any letter) with _DRAFT_
    # Also handle trailing _RevA.docx / _RevB.docx (no trailing underscore)
    $n = $name -replace '_Rev[A-Z]_', '_DRAFT_'
    $n = $n -replace '_Rev[A-Z]\.docx$', '_DRAFT_.docx'
    # Clean up double underscores that might result
    $n = $n -replace '__+', '_'
    return $n
}

function Patch-DocxHeaders($filePath) {
    # A DOCX is a ZIP. We need to patch word/header1.xml, word/header2.xml,
    # word/header3.xml, and word/document.xml for any revision/date references.
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    $patchedAny = $false
    $zipPath = $filePath

    # Work on a copy
    $workPath = $filePath + ".work"
    Copy-Item $zipPath $workPath -Force

    try {
        $zip = [System.IO.Compression.ZipFile]::Open($workPath, 'Update')

        $targets = $zip.Entries | Where-Object {
            $_.FullName -match '^word/(header\d*|footer\d*|document)\.xml$'
        }

        foreach ($entry in $targets) {
            $reader  = New-Object System.IO.StreamReader($entry.Open())
            $content = $reader.ReadToEnd()
            $reader.Close()

            $original = $content

            # ── Revision field replacements ──────────────────────────────
            # "Revision: A" / "Revision: B" / "Revision: Rev A" etc.
            $content = $content -replace 'Revision:\s*Rev\s+[A-Z]', 'Revision: DRAFT'
            $content = $content -replace 'Revision:\s*[A-Z]\b', 'Revision: DRAFT'
            # Plain "Rev A" / "Rev B" label text (not inside a larger word)
            $content = $content -replace '>Rev [A-Z]<', '>DRAFT<'
            # XML text nodes that are just "Rev A" or "Rev B"
            $content = $content -replace '(?<=>)Rev [A-Z](?=<)', 'DRAFT'

            # ── Effective date replacements ──────────────────────────────
            # Common patterns: "Effective Date: MM/DD/YYYY", date placeholders,
            # "TBD", or any date-like string in the effective date field.
            # We target the text that follows "Effective Date:" or "Effective:"
            $content = $content -replace '(Effective\s*Date\s*:?\s*)(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4})', '$1TBD — Pending Approval'
            $content = $content -replace '(Effective\s*Date\s*:?\s*)(TBD|N\/A|Pending)', '$1TBD — Pending Approval'
            # Also catch standalone date strings near effective date context
            # (conservative — only replace if the surrounding XML text contains "Effective")
            if ($content -match 'Effective') {
                $content = $content -replace '(?<=Effective[^<]{0,50})\b\d{1,2}[\/\-]\d{1,2}[\/\-]\d{2,4}\b', 'TBD — Pending Approval'
            }

            if ($content -ne $original) {
                $patchedAny = $true
                $writer = New-Object System.IO.StreamWriter($entry.Open())
                $writer.BaseStream.SetLength(0)
                $writer.Write($content)
                $writer.Close()
            }
        }

        $zip.Dispose()
    } catch {
        if ($zip) { $zip.Dispose() }
        Remove-Item $workPath -Force -ErrorAction SilentlyContinue
        throw
    }

    if ($patchedAny) {
        Move-Item $workPath $zipPath -Force
    } else {
        Remove-Item $workPath -Force
    }

    return $patchedAny
}

# ── Main ────────────────────────────────────────────────────────────────────

New-Item -ItemType Directory -Force -Path $TempDir | Out-Null
"" | Set-Content $LogFile
Write-Log "=== patch_docx_headers.ps1 started ==="
Write-Log "Temp dir: $TempDir"
Write-Log ""

$totalFiles   = 0
$patchedFiles = 0
$renamedFiles = 0
$errors       = 0

foreach ($folderDef in $Folders) {
    $folderPath = $folderDef.Path
    $zone       = $folderDef.Zone

    Write-Log "── Folder: $folderPath ($zone) ──"

    try {
        $files = Get-PnPFolderItem -FolderSiteRelativeUrl $folderPath -ItemType File |
                 Where-Object { $_.Name -like "*.docx" }
    } catch {
        Write-Log "  [ERROR] Cannot list folder: $_"
        $errors++
        continue
    }

    foreach ($file in $files) {
        $originalName = $file.Name
        $serverRelUrl = $file.ServerRelativeUrl
        $totalFiles++

        Write-Log "  Processing: $originalName"

        # Download to temp
        $localPath = Join-Path $TempDir $originalName
        try {
            Get-PnPFile -Url $serverRelUrl -Path $TempDir -Filename $originalName -AsFile -Force | Out-Null
        } catch {
            Write-Log "    [ERROR] Download failed: $_"
            $errors++
            continue
        }

        # Patch internal headers
        $headerPatched = $false
        try {
            $headerPatched = Patch-DocxHeaders -filePath $localPath
            if ($headerPatched) {
                Write-Log "    [PATCHED] Headers updated"
                $patchedFiles++
            } else {
                Write-Log "    [OK] Headers already clean — no changes needed"
            }
        } catch {
            Write-Log "    [ERROR] Header patch failed: $_"
            $errors++
            continue
        }

        # Determine upload filename
        $uploadName = $originalName
        $renamed    = $false

        if ($zone -eq "Published") {
            $newName = Get-DraftFilename $originalName
            if ($newName -ne $originalName) {
                $uploadName = $newName
                $renamed    = $true
                $renamedFiles++
                Write-Log "    [RENAME] $originalName → $uploadName"
            }
        }

        # Upload (always re-upload if headers were patched OR renamed)
        if ($headerPatched -or $renamed) {
            try {
                # Upload with new filename
                $spFolder = $folderPath  # relative to site root
                Add-PnPFile -Path $localPath -Folder $spFolder -NewFileName $uploadName -Values @{} | Out-Null
                Write-Log "    [UPLOADED] $uploadName → $spFolder"

                # If renamed, delete the old file
                if ($renamed) {
                    try {
                        Remove-PnPFile -ServerRelativeUrl $serverRelUrl -Force -Recycle
                        Write-Log "    [DELETED] Old file: $originalName"
                    } catch {
                        Write-Log "    [WARN] Could not delete old file $originalName : $_"
                    }
                }
            } catch {
                Write-Log "    [ERROR] Upload failed: $_"
                $errors++
            }
        }

        # Clean up temp file
        Remove-Item $localPath -Force -ErrorAction SilentlyContinue
    }

    Write-Log ""
}

Write-Log "=== Summary ==="
Write-Log "  Total files processed : $totalFiles"
Write-Log "  Headers patched       : $patchedFiles"
Write-Log "  Files renamed         : $renamedFiles"
Write-Log "  Errors                : $errors"
Write-Log ""
Write-Log "Log written to: $LogFile"

if ($errors -gt 0) {
    Write-Host "`n⚠️  Completed with $errors error(s). Review log: $LogFile" -ForegroundColor Yellow
} else {
    Write-Host "`n✅  All done. $patchedFiles file(s) patched, $renamedFiles renamed." -ForegroundColor Green
}
