# setup_dco_fields.ps1
# Adds 4 implementation-activity fields to QMS_DCOs
# Adds Emp_PortalRole choice field to QMS_Employees

param(
    [string]$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177"
)

Connect-PnPOnline -Url $SiteUrl -Interactive

Write-Host "`nAdding fields to QMS_DCOs..." -ForegroundColor Cyan

# DCO_ImplActivityRequired — Yes/No default No
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Implementation Activity Required" `
        -InternalName "DCO_ImplActivityRequired" -Type Boolean -AddToDefaultView
    Set-PnPField -List "QMS_DCOs" -Identity "DCO_ImplActivityRequired" `
        -Values @{DefaultValue="0"}
    Write-Host "  [OK] DCO_ImplActivityRequired" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_ImplActivityRequired — $($_.Exception.Message)" -ForegroundColor Yellow }

# DCO_ImplPlan — multiline text
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Implementation Plan" `
        -InternalName "DCO_ImplPlan" -Type Note -AddToDefaultView
    Write-Host "  [OK] DCO_ImplPlan" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_ImplPlan — $($_.Exception.Message)" -ForegroundColor Yellow }

# DCO_ImplNoSoonerThan — DateTime
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Implementation No-Sooner-Than" `
        -InternalName "DCO_ImplNoSoonerThan" -Type DateTime -AddToDefaultView
    Write-Host "  [OK] DCO_ImplNoSoonerThan" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_ImplNoSoonerThan — $($_.Exception.Message)" -ForegroundColor Yellow }

# DCO_EffNoSoonerThan — DateTime
try {
    Add-PnPField -List "QMS_DCOs" -DisplayName "Effective No-Sooner-Than" `
        -InternalName "DCO_EffNoSoonerThan" -Type DateTime -AddToDefaultView
    Write-Host "  [OK] DCO_EffNoSoonerThan" -ForegroundColor Green
} catch { Write-Host "  [SKIP] DCO_EffNoSoonerThan — $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host "`nAdding Emp_PortalRole to QMS_Employees..." -ForegroundColor Cyan

try {
    Add-PnPField -List "QMS_Employees" -DisplayName "Portal Role" `
        -InternalName "Emp_PortalRole" -Type Choice -AddToDefaultView
    Set-PnPField -List "QMS_Employees" -Identity "Emp_PortalRole" `
        -Values @{ Choices = [string[]]@("PM", "Change Analyst", "External") }
    Write-Host "  [OK] Emp_PortalRole" -ForegroundColor Green
} catch { Write-Host "  [SKIP] Emp_PortalRole — $($_.Exception.Message)" -ForegroundColor Yellow }

Write-Host "`nDone. All fields processed." -ForegroundColor Cyan
