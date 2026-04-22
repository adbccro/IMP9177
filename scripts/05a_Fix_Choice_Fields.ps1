# =============================================================================
# IMP9177 — Fix 3 Choice fields that failed due to XML special characters
# Emp_Dept (QMS_Employees), Rec_Status (QMS_Records), TC_Method (QMS_TrainingCompletions)
# =============================================================================

param([string]$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177")

$ClientId = "ba48ac81-6f23-43bd-9797-ec2866071102"
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
Write-Host "[OK] Connected`n" -ForegroundColor Green

function Fix-ChoiceField {
    param([string]$List, [string]$FieldName, [string[]]$Choices)
    $existing = Get-PnPField -List $List -Identity $FieldName -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  [EXISTS]  $FieldName on $List" -ForegroundColor DarkGray
        return
    }
    # Build choices using encoded ampersand
    $choiceXml = "<CHOICES>" + ($Choices | ForEach-Object {
        $encoded = $_ -replace "&", "&amp;"
        "<CHOICE>$encoded</CHOICE>"
    }) + "</CHOICES>"
    $schema = "<Field Type='Choice' DisplayName='$FieldName' Name='$FieldName' StaticName='$FieldName'>$choiceXml</Field>"
    try {
        Add-PnPFieldFromXml -List $List -FieldXml $schema -ErrorAction Stop | Out-Null
        Write-Host "  [CREATED] $FieldName on $List" -ForegroundColor Green
    } catch {
        Write-Warning "  [FAILED]  $FieldName — $_"
    }
}

Write-Host "=== Fixing Emp_Dept on QMS_Employees ===" -ForegroundColor Yellow
Fix-ChoiceField "QMS_Employees" "Emp_Dept" @(
    "Quality","Production","Warehouse","R&D","Management","Purchasing","Marketing"
)

Write-Host "`n=== Fixing Rec_Status on QMS_Records ===" -ForegroundColor Yellow
Fix-ChoiceField "QMS_Records" "Rec_Status" @(
    "Draft","In Review","Approved","Pending Signature","Signed & Filed"
)

Write-Host "`n=== Fixing TC_Method on QMS_TrainingCompletions ===" -ForegroundColor Yellow
Fix-ChoiceField "QMS_TrainingCompletions" "TC_Method" @(
    "Read & Understand","Instructor-led","On-the-Job","Computer-based","Observation"
)

Write-Host "`n[OK] Choice field fixes complete. Run 05_Provision_Lists.ps1 -SeedData next.`n" -ForegroundColor Cyan
