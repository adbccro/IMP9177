# setup_new_fields.ps1
# Adds 8 new fields to QMS_DCOs for Tasks 1-8
# Run once against the IMP9177 site.

param(
    [string]$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177"
)

Connect-PnPOnline -Url $SiteUrl -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

Write-Host "`nAdding fields to QMS_DCOs..." -ForegroundColor Cyan

# Task 1: DCO_EffDelayRequired — Choice (Yes/No) default No
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Eff. Delay Required" `
        -InternalName "DCO_EffDelayRequired" -Type Choice -AddToDefaultView
    Set-PnPField -List "QMS_DCOs" -Identity "DCO_EffDelayRequired" `
        -Values @{ Choices = [string[]]@("No", "Yes"); DefaultValue = "No" }
    Write-Host "  [OK] DCO_EffDelayRequired" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_EffDelayRequired — $($_.Exception.Message)" -ForegroundColor Yellow }

# Task 1: DCO_EffDelayReason — multiline text
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Eff. Delay Reason" `
        -InternalName "DCO_EffDelayReason" -Type Note -AddToDefaultView
    Write-Host "  [OK] DCO_EffDelayReason" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_EffDelayReason — $($_.Exception.Message)" -ForegroundColor Yellow }

# Task 3: DCO_ImplOwner — single line text
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Impl. Owner" `
        -InternalName "DCO_ImplOwner" -Type Text -AddToDefaultView
    Write-Host "  [OK] DCO_ImplOwner" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_ImplOwner — $($_.Exception.Message)" -ForegroundColor Yellow }

# Task 3: DCO_ImplDescription — multiline text
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Impl. Description" `
        -InternalName "DCO_ImplDescription" -Type Note -AddToDefaultView
    Write-Host "  [OK] DCO_ImplDescription" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_ImplDescription — $($_.Exception.Message)" -ForegroundColor Yellow }

# Task 3: DCO_ImplRisks — multiline text
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Impl. Risks" `
        -InternalName "DCO_ImplRisks" -Type Note -AddToDefaultView
    Write-Host "  [OK] DCO_ImplRisks" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_ImplRisks — $($_.Exception.Message)" -ForegroundColor Yellow }

# Task 3: DCO_ImplVerification — multiline text
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Impl. Verification" `
        -InternalName "DCO_ImplVerification" -Type Note -AddToDefaultView
    Write-Host "  [OK] DCO_ImplVerification" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_ImplVerification — $($_.Exception.Message)" -ForegroundColor Yellow }

# Task 7: DCO_CancelReason — multiline text
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Cancellation Reason" `
        -InternalName "DCO_CancelReason" -Type Note -AddToDefaultView
    Write-Host "  [OK] DCO_CancelReason" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_CancelReason — $($_.Exception.Message)" -ForegroundColor Yellow }

# Task 8: DCO_DocsLastUpdated — DateTime
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Docs Last Updated" `
        -InternalName "DCO_DocsLastUpdated" -Type DateTime -AddToDefaultView
    Write-Host "  [OK] DCO_DocsLastUpdated" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_DocsLastUpdated — $($_.Exception.Message)" -ForegroundColor Yellow }

# Task 4: DCO_DocPurposes — multiline text (JSON blob for per-doc P&J)
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Doc Purposes JSON" `
        -InternalName "DCO_DocPurposes" -Type Note -AddToDefaultView
    Write-Host "  [OK] DCO_DocPurposes" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_DocPurposes — $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host "`nDone. All new QMS_DCOs fields processed." -ForegroundColor Cyan
