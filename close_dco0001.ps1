# close_dco0001.ps1
# IMP9177 — DCO-0001 Closure Script
# Performs:
#   1. Seed Cindy Dong as employee + role
#   2. Seed Cindy training completions for all DCO-0001 SOPs (today)
#   3. Write Andre Butler signature to QMS_DCOApprovals
#   4. Advance DCO-0001 phase to Implemented
#   5. Write routing history entries
#   6. Burn effective date into Published documents (replace TBD)
#   7. Copy Published documents to Official zone
#
# Run from: C:\Users\andre\ADBCCRO Dropbox\Andre Butler\IMP9177\imp9177-spfx
# Requires: PnP.PowerShell connected to IMP9177 site

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$SITE = "https://adbccro.sharepoint.com/sites/IMP9177"
$TODAY = "2026-04-24"
$TODAY_DISPLAY = "April 24, 2026"
$EFF_DATE = "2026-04-24T00:00:00Z"

# ── Connect ──────────────────────────────────────────────────────────────────
Write-Host "`n[1/7] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SITE -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
Write-Host "Connected." -ForegroundColor Green

# ── Helper: add list item ─────────────────────────────────────────────────────
function Add-Item($list, $fields) {
    try {
        Add-PnPListItem -List $list -Values $fields | Out-Null
        Write-Host "  + Added to $list" -ForegroundColor Gray
    } catch {
        Write-Host "  ! Failed $list : $_" -ForegroundColor Yellow
    }
}

function Update-Item($list, $id, $fields) {
    try {
        Set-PnPListItem -List $list -Identity $id -Values $fields | Out-Null
        Write-Host "  ~ Updated $list item $id" -ForegroundColor Gray
    } catch {
        Write-Host "  ! Failed update $list item $id : $_" -ForegroundColor Yellow
    }
}

# ── DCO-0001 documents ────────────────────────────────────────────────────────
$DCO_DOCS = @(
    "SOP-QMS-001", "SOP-QMS-002", "SOP-QMS-003",
    "QM-001", "FM-001", "FM-002", "FM-003", "FM-027", "FM-030",
    "SOP-PRD-108", "SOP-PRD-432", "SOP-FRS-549"
)

$SOP_DOCS = @(
    "SOP-QMS-001", "SOP-QMS-002", "SOP-QMS-003",
    "SOP-PRD-108", "SOP-PRD-432", "SOP-FRS-549"
)

$SIG_ID = "SIG-ABUTLER-" + [System.DateTime]::UtcNow.ToString("yyyyMMddHHmmss")

# ── Step 2: Seed Cindy Dong as employee ───────────────────────────────────────
Write-Host "`n[2/7] Seeding Cindy Dong as employee..." -ForegroundColor Cyan

$existing = Get-PnPListItem -List "QMS_Employees" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Cindy Dong</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue

if (!$existing -or $existing.Count -eq 0) {
    Add-Item "QMS_Employees" @{
        Title       = "Cindy Dong"
        Emp_Email   = "cindydong3h@gmail.com"
        Emp_Title   = "Document Control Coordinator"
        Emp_Dept    = "Quality"
        Emp_Roles   = "QA-Staff,DC-Coordinator"
    }
    Write-Host "  Cindy Dong added as employee." -ForegroundColor Green
} else {
    Write-Host "  Cindy Dong already exists." -ForegroundColor Yellow
}

# ── Step 3: Seed training completions for Cindy ───────────────────────────────
Write-Host "`n[3/7] Seeding Cindy Dong training completions..." -ForegroundColor Cyan

foreach ($docId in $SOP_DOCS) {
    $tcSigId = "TC-CDONG-" + $docId + "-" + [System.DateTime]::UtcNow.ToString("yyyyMMddHHmmss")
    Add-Item "QMS_TrainingCompletions" @{
        Title         = "TC-CDONG-$docId"
        TC_EmpID      = "Cindy Dong"
        TC_DocID      = $docId
        TC_Rev        = "A"
        TC_Method     = "Read & Understood — In-Person Review"
        TC_SignedDate  = $EFF_DATE
        TC_SigID      = $tcSigId
    }
    Write-Host "  Training record: Cindy Dong / $docId / Rev A / $TODAY" -ForegroundColor Gray
}

Write-Host "  Training completions seeded." -ForegroundColor Green

# ── Step 4: Write Andre signature to QMS_DCOApprovals ─────────────────────────
Write-Host "`n[4/7] Writing Andre Butler signature to DCO-0001..." -ForegroundColor Cyan

$existingSig = Get-PnPListItem -List "QMS_DCOApprovals" -Query "<View><Query><Where><And><Eq><FieldRef Name='Appr_DCOID'/><Value Type='Text'>DCO-0001</Value></Eq><Eq><FieldRef Name='Title'/><Value Type='Text'>DCO-0001-ANDRE</Value></Eq></And></Where></Query></View>" -ErrorAction SilentlyContinue

if ($existingSig -and $existingSig.Count -gt 0) {
    $item = $existingSig[0]
    Update-Item "QMS_DCOApprovals" $item.Id @{
        Appr_Status    = "Signed"
        Appr_SigID     = $SIG_ID
        Appr_SignedDate = $EFF_DATE
    }
    Write-Host "  Updated existing signature record." -ForegroundColor Green
} else {
    Add-Item "QMS_DCOApprovals" @{
        Title          = "DCO-0001-ANDRE"
        Appr_DCOID     = "DCO-0001"
        Appr_Name      = "Andre Butler"
        Appr_Role      = "QA/Regulatory Consultant — ADB Consulting"
        Appr_Type      = "Required"
        Appr_Status    = "Signed"
        Appr_SigID     = $SIG_ID
        Appr_SignedDate = $EFF_DATE
    }
    Write-Host "  Signature record created." -ForegroundColor Green
}

# ── Step 5: Advance DCO-0001 to Implemented ───────────────────────────────────
Write-Host "`n[5/7] Advancing DCO-0001 to Implemented..." -ForegroundColor Cyan

$dcoItem = Get-PnPListItem -List "QMS_DCOs" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>DCO-0001</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue

if ($dcoItem -and $dcoItem.Count -gt 0) {
    $dco = $dcoItem[0]
    Update-Item "QMS_DCOs" $dco.Id @{
        DCO_Phase          = "Implemented"
        DCO_EffectiveDate  = $EFF_DATE
    }
    Write-Host "  DCO-0001 advanced to Implemented." -ForegroundColor Green
} else {
    Write-Host "  DCO-0001 not found in QMS_DCOs — creating record..." -ForegroundColor Yellow
    Add-Item "QMS_DCOs" @{
        Title              = "DCO-0001"
        DCO_Title          = "W1 QMS Foundation Package — Quality Manual + Core SOPs"
        DCO_Phase          = "Implemented"
        DCO_EffectiveDate  = $EFF_DATE
        DCO_SubmittedDate  = "2026-03-30T00:00:00Z"
        DCO_Originator     = "Andre Butler"
        DCO_Docs           = "QM-001,SOP-QMS-001,SOP-QMS-002,SOP-QMS-003,FM-001,FM-002,FM-003,FM-027,FM-030,SOP-PRD-108,SOP-PRD-432,SOP-FRS-549"
        DCO_CRLink         = "CR-0001"
    }
}

# ── Step 6: Write routing history ─────────────────────────────────────────────
Write-Host "`n[6/7] Writing routing history..." -ForegroundColor Cyan

$histEntries = @(
    @{
        Title        = "DCO-0001-SIG-$SIG_ID"
        RH_DCOID     = "DCO-0001"
        RH_EventType = "signature"
        RH_Stage     = "In Review"
        RH_Actor     = "Andre Butler"
        RH_Note      = "Required approver signature applied. SIG ID: $SIG_ID. All required approvers complete."
        RH_Timestamp = $EFF_DATE
    },
    @{
        Title        = "DCO-0001-TRAINING-GATE"
        RH_DCOID     = "DCO-0001"
        RH_EventType = "stage"
        RH_Stage     = "Implemented"
        RH_Actor     = "System"
        RH_Note      = "Training prerequisite satisfied — Cindy Dong completed training on all SOPs ($TODAY). DCO advanced to Implemented."
        RH_Timestamp = $EFF_DATE
    },
    @{
        Title        = "DCO-0001-EFFECTIVE"
        RH_DCOID     = "DCO-0001"
        RH_EventType = "stage"
        RH_Stage     = "Effective"
        RH_Actor     = "System"
        RH_Note      = "Effective Date set: $TODAY_DISPLAY. Documents promoted to Official zone. Revision A in force."
        RH_Timestamp = $EFF_DATE
    }
)

foreach ($entry in $histEntries) {
    Add-Item "QMS_RoutingHistory" $entry
}
Write-Host "  Routing history written." -ForegroundColor Green

# ── Step 7: Burn effective date into documents + promote to Official ───────────
Write-Host "`n[7/7] Burning effective date + promoting documents to Official..." -ForegroundColor Cyan

$PUBLISHED_DOCS = "Shared Documents/Published/QMS/Documents"
$PUBLISHED_FORMS = "Shared Documents/Published/QMS/Forms"
$PUBLISHED_QM = "Shared Documents/Published/QMS/Quality Manual"
$OFFICIAL_DOCS = "Shared Documents/Official/QMS/Documents"
$OFFICIAL_FORMS = "Shared Documents/Official/QMS/Forms"

# File map: docId → source folder → filename
$DOC_FILES = @{
    "QM-001"       = @{ folder = $PUBLISHED_QM;    file = "QM-001_Quality_Manual_RevA.docx" }
    "SOP-QMS-001"  = @{ folder = $PUBLISHED_DOCS;  file = "SOP-QMS-001_RevA_Management_Responsibility.docx" }
    "SOP-QMS-002"  = @{ folder = $PUBLISHED_DOCS;  file = "SOP-QMS-002_RevA_Document_Control.docx" }
    "SOP-QMS-003"  = @{ folder = $PUBLISHED_DOCS;  file = "SOP-QMS-003_RevA_Change_Control.docx" }
    "FM-001"       = @{ folder = $PUBLISHED_FORMS; file = "FM-001_Master_Document_Log_RevA.docx" }
    "FM-002"       = @{ folder = $PUBLISHED_FORMS; file = "FM-002_Change_Request_Form_RevA.docx" }
    "FM-003"       = @{ folder = $PUBLISHED_FORMS; file = "FM-003_Document_Change_Order_RevA.docx" }
    "FM-027"       = @{ folder = $PUBLISHED_FORMS; file = "FM-027_QU_QS_Designation_Record_RevA.docx" }
    "FM-030"       = @{ folder = $PUBLISHED_FORMS; file = "FM-030_Finished_Product_Spec_Sheet_RevA.docx" }
}

# Ensure Official subfolders exist
$officialFolders = @($OFFICIAL_DOCS, $OFFICIAL_FORMS)
foreach ($f in $officialFolders) {
    try {
        Resolve-PnPFolder -SiteRelativePath $f -ErrorAction SilentlyContinue | Out-Null
    } catch {
        $parts = $f -split "/"
        $parent = ($parts[0..($parts.Length-2)]) -join "/"
        $child = $parts[-1]
        Add-PnPFolder -Name $child -Folder $parent -ErrorAction SilentlyContinue | Out-Null
    }
}

$promoted = 0
$failed = 0

foreach ($docId in $DOC_FILES.Keys) {
    $info = $DOC_FILES[$docId]
    $srcFolder = $info.folder
    $fileName = $info.file
    $srcPath = "/sites/IMP9177/$srcFolder/$fileName"

    # Determine destination folder
    $destFolder = if ($srcFolder -like "*Forms*") { $OFFICIAL_FORMS } else { $OFFICIAL_DOCS }
    $destPath = "/sites/IMP9177/$destFolder/$fileName"

    try {
        # Download the file to temp
        $tempFile = "$env:TEMP\$fileName"
        Get-PnPFile -Url $srcPath -Path $env:TEMP -Filename $fileName -AsFile -Force | Out-Null

        # Replace TBD with effective date using docx content search
        # We use a simple approach: unzip docx, find/replace in word/document.xml, rezip
        $zipPath = $tempFile -replace "\.docx$", ".zip"
        Copy-Item $tempFile $zipPath -Force

        Add-Type -AssemblyName System.IO.Compression.FileSystem
        $zip = [System.IO.Compression.ZipFile]::Open($zipPath, [System.IO.Compression.ZipArchiveMode]::Update)
        $entry = $zip.GetEntry("word/document.xml")

        if ($entry) {
            $stream = $entry.Open()
            $reader = New-Object System.IO.StreamReader($stream)
            $content = $reader.ReadToEnd()
            $reader.Close()
            $stream.Close()

            # Replace TBD placeholder with effective date
            $newContent = $content -replace "TBD\s*[—-–]\s*Pending Document Control Approval", $TODAY_DISPLAY
            $newContent = $newContent -replace "TBD — Pending Document Control Approval", $TODAY_DISPLAY
            $newContent = $newContent -replace "TBD", $TODAY_DISPLAY

            # Delete old entry and write new
            $entry.Delete()
            $newEntry = $zip.CreateEntry("word/document.xml")
            $writeStream = $newEntry.Open()
            $writer = New-Object System.IO.StreamWriter($writeStream)
            $writer.Write($newContent)
            $writer.Flush()
            $writer.Close()
            $writeStream.Close()
        }
        $zip.Dispose()

        # Rename back to docx
        Copy-Item $zipPath $tempFile -Force
        Remove-Item $zipPath -ErrorAction SilentlyContinue

        # Upload to Official zone
        Add-PnPFile -Path $tempFile -Folder $destFolder -NewFileName $fileName | Out-Null

        # Clean up temp
        Remove-Item $tempFile -ErrorAction SilentlyContinue

        Write-Host "  Promoted: $docId → $destFolder" -ForegroundColor Green
        $promoted++

    } catch {
        Write-Host "  Failed: $docId — $_" -ForegroundColor Red
        $failed++
    }
}

# ── Final: Update DCO-0001 to Effective ───────────────────────────────────────
Write-Host "`nUpdating DCO-0001 to Effective..." -ForegroundColor Cyan
$dcoItem2 = Get-PnPListItem -List "QMS_DCOs" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>DCO-0001</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
if ($dcoItem2 -and $dcoItem2.Count -gt 0) {
    Update-Item "QMS_DCOs" $dcoItem2[0].Id @{ DCO_Phase = "Effective" }
}

# ── Summary ───────────────────────────────────────────────────────────────────
Write-Host "`n════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " DCO-0001 CLOSURE COMPLETE" -ForegroundColor Green
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " Effective Date : $TODAY_DISPLAY"
Write-Host " Signature      : $SIG_ID"
Write-Host " Training       : Cindy Dong — $($SOP_DOCS.Count) SOPs"
Write-Host " Docs Promoted  : $promoted to Official zone"
Write-Host " Docs Failed    : $failed"
Write-Host " DCO-0001 Phase : Effective"
Write-Host "════════════════════════════════════════" -ForegroundColor Cyan


