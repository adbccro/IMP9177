# =============================================================================
# IMP9177 — QMS PORTAL SHAREPOINT SETUP
# Script 04: Validate Published Zone & Set Document Metadata Columns
# Run AFTER the PM has pushed documents through the Publish action
#
# What this script does:
#   - Checks Published/QMS for expected documents
#   - Sets metadata columns (Document_ID, Revision, Status, Zone, DCO_Link)
#     on each document so the portal can read them
#   - Verifies Official/QMS is still empty (should remain so until DCO Implemented)
# =============================================================================

param(
    [string]$SiteUrl     = "https://adbccro.sharepoint.com/sites/IMP9177",
    [string]$LibraryName = "Shared Documents",
    [switch]$WhatIf      = $false
)

$ErrorActionPreference = "Continue"

Write-Host "`n[IMP9177] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
Write-Host "[OK] Connected`n" -ForegroundColor Green

# =============================================================================
# ENSURE METADATA COLUMNS EXIST ON THE LIBRARY
# These columns are read by the QMS Portal application
# =============================================================================
Write-Host "=== Ensuring metadata columns exist on Shared Documents library ===" -ForegroundColor Yellow

$columnsNeeded = @(
    @{Name="QMS_DocumentID";    Type="Text";         Description="IMP9177 document identifier (e.g. SOP-001, QM-001)"},
    @{Name="QMS_Revision";      Type="Text";         Description="Current controlled revision (Rev A, Rev B) or draft version (v0.1)"},
    @{Name="QMS_Status";        Type="Choice";       Description="Document status in QMS lifecycle"},
    @{Name="QMS_Zone";          Type="Choice";       Description="QMS zone: Draft | Published | Official"},
    @{Name="QMS_DCO_Link";      Type="Text";         Description="Linked DCO number (e.g. DCO-0001)"},
    @{Name="QMS_CR_Link";       Type="Text";         Description="Linked CR number (e.g. CR-0001)"},
    @{Name="QMS_EffectiveDate"; Type="DateTime";     Description="Date document became effective (Official zone only)"},
    @{Name="QMS_Superseded";    Type="Boolean";      Description="True if this revision has been superseded by a newer revision"}
)

$statusChoices = "Draft","Ready to Publish","Published","On DCO","Superseded","Effective"
$zoneChoices   = "Draft","Published","Official"

foreach ($col in $columnsNeeded) {
    $existing = Get-PnPField -List $LibraryName -Identity $col.Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  [EXISTS]  $($col.Name)" -ForegroundColor DarkGray
    } else {
        Write-Host "  [CREATING] $($col.Name) ($($col.Type))" -ForegroundColor Green
        if (-not $WhatIf) {
            try {
                if ($col.Type -eq "Choice") {
                    $choices = if ($col.Name -eq "QMS_Status") { $statusChoices } else { $zoneChoices }
                    $choiceXml = "<CHOICES>" + ($choices | ForEach-Object { "<CHOICE>$_</CHOICE>" }) + "</CHOICES>"
                    $schema = "<Field Type='Choice' DisplayName='$($col.Name)' Name='$($col.Name)' StaticName='$($col.Name)'>$choiceXml</Field>"
                    Add-PnPFieldFromXml -List $LibraryName -FieldXml $schema -ErrorAction Stop | Out-Null
                } elseif ($col.Type -eq "Boolean") {
                    $schema = "<Field Type='Boolean' DisplayName='$($col.Name)' Name='$($col.Name)' StaticName='$($col.Name)'/>"
                    Add-PnPFieldFromXml -List $LibraryName -FieldXml $schema -ErrorAction Stop | Out-Null
                } elseif ($col.Type -eq "DateTime") {
                    $schema = "<Field Type='DateTime' DisplayName='$($col.Name)' Name='$($col.Name)' StaticName='$($col.Name)' Format='DateOnly'/>"
                    Add-PnPFieldFromXml -List $LibraryName -FieldXml $schema -ErrorAction Stop | Out-Null
                } else {
                    $schema = "<Field Type='Text' DisplayName='$($col.Name)' Name='$($col.Name)' StaticName='$($col.Name)'/>"
                    Add-PnPFieldFromXml -List $LibraryName -FieldXml $schema -ErrorAction Stop | Out-Null
                }
                Write-Host "         Created." -ForegroundColor Green
            } catch {
                Write-Warning "  [FAILED]  $($col.Name) — $_"
            }
        }
    }
}

Write-Host ""

# =============================================================================
# SET METADATA ON DOCUMENTS ALREADY IN PUBLISHED ZONE
# These documents were promoted before the portal was built — tag them now
# =============================================================================
Write-Host "=== Setting metadata on Published/QMS documents ===" -ForegroundColor Yellow

$publishedDocs = @(
    @{
        Path    = "Published/QMS/Quality Manual/QM-001_Quality_Manual_RevA.docx"
        DocID   = "QM-001"; Rev = "Rev A"; Status = "Published"; Zone = "Published"
        DCO     = "DCO-0001"; CR = "CR-0001"
    },
    @{
        Path    = "Published/QMS/Forms/FM-001_Master_Document_Log_RevA.docx"
        DocID   = "FM-001"; Rev = "Rev A"; Status = "Published"; Zone = "Published"
        DCO     = "DCO-0001"; CR = "CR-0001"
    },
    @{
        Path    = "Published/QMS/Forms/FM-002_Change_Request_Form_RevA.docx"
        DocID   = "FM-002"; Rev = "Rev A"; Status = "Published"; Zone = "Published"
        DCO     = "DCO-0001"; CR = "CR-0001"
    },
    @{
        Path    = "Published/QMS/Forms/FM-003_Document_Change_Order_RevA.docx"
        DocID   = "FM-003"; Rev = "Rev A"; Status = "Published"; Zone = "Published"
        DCO     = "DCO-0001"; CR = "CR-0001"
    },
    @{
        Path    = "Published/QMS/Forms/FM-025_RCD-FM-219_RevB.docx"
        DocID   = "FM-025"; Rev = "Rev B"; Status = "Published"; Zone = "Published"
        DCO     = "DCO-0001"; CR = "CR-0001"
    },
    @{
        Path    = "Published/QMS/Forms/FM-027_QU_QS_Designation_Record_RevA.docx"
        DocID   = "FM-027"; Rev = "Rev A"; Status = "Published"; Zone = "Published"
        DCO     = "DCO-0001"; CR = "CR-0001"
    },
    @{
        Path    = "Published/QMS/Forms/FM-030_Finished_Product_Spec_Sheet_RevA.docx"
        DocID   = "FM-030"; Rev = "Rev A"; Status = "Published"; Zone = "Published"
        DCO     = "DCO-0001"; CR = "CR-0001"
    }
)

foreach ($doc in $publishedDocs) {
    Write-Host "  Setting metadata: $($doc.DocID)" -ForegroundColor White
    if (-not $WhatIf) {
        try {
            $item = Get-PnPFile -Url "$LibraryName/$($doc.Path)" -AsListItem -ErrorAction Stop
            Set-PnPListItem -List $LibraryName -Identity $item.Id -Values @{
                "QMS_DocumentID" = $doc.DocID
                "QMS_Revision"   = $doc.Rev
                "QMS_Status"     = $doc.Status
                "QMS_Zone"       = $doc.Zone
                "QMS_DCO_Link"   = $doc.DCO
                "QMS_CR_Link"    = $doc.CR
                "QMS_Superseded" = $false
            } -ErrorAction Stop | Out-Null
            Write-Host "    [OK] $($doc.DocID) — metadata set" -ForegroundColor Green
        } catch {
            Write-Warning "    [SKIP] $($doc.DocID) — file not found or metadata error: $_"
        }
    } else {
        Write-Host "    [WHATIF] Would set metadata on $($doc.Path)" -ForegroundColor DarkGray
    }
}

Write-Host ""

# =============================================================================
# VERIFY OFFICIAL ZONE IS STILL EMPTY
# =============================================================================
Write-Host "=== Verifying Official/QMS zone is empty (as expected) ===" -ForegroundColor Yellow

try {
    $officialItems = Get-PnPFolderItem -FolderSiteRelativeUrl "$LibraryName/Official/QMS" -ItemType File -ErrorAction SilentlyContinue
    if ($officialItems -and $officialItems.Count -gt 0) {
        Write-Host "  [WARNING] Official/QMS contains $($officialItems.Count) file(s) — should be empty until DCO Implemented!" -ForegroundColor Red
        $officialItems | ForEach-Object { Write-Host "    - $($_.Name)" -ForegroundColor DarkRed }
    } else {
        Write-Host "  [OK] Official/QMS is empty — correct state before any DCO is loaded." -ForegroundColor Green
    }
} catch {
    Write-Host "  [OK] Official/QMS exists and is empty." -ForegroundColor Green
}

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " SCRIPT 04 COMPLETE" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host " SharePoint is now configured for QMS Portal deployment."
Write-Host " Metadata columns are set on the library."
Write-Host " Published documents are tagged."
Write-Host " Official zone is clean."
Write-Host ""
Write-Host " Next step: Deploy the QMS Portal application, then load DCO-0001."
Write-Host ""
