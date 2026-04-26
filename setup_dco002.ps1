# setup_dco002.ps1
# Creates DCO-0002 in QMS_DCOs with full 27-document set, Draft status, linked to CR-0002.
# Run once after deploying the updated portal.

param(
    [string]$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177"
)

Connect-PnPOnline -Url $SiteUrl -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

$docs = "QM-001,SOP-QMS-001,SOP-QMS-002,SOP-QMS-003,SOP-PRD-108,SOP-PRD-432,SOP-FRS-549,SOP-RCL-321,SOP-SUP-001,SOP-SUP-002,SOP-FS-001,SOP-FS-002,SOP-FS-003,SOP-FS-004,SOP-PC-001,FM-001,FM-002,FM-003,FM-004,FM-005,FM-006,FM-007,FM-008,FM-025,FM-027,FM-030,FM-ALG"

Write-Host "Creating DCO-0002..." -ForegroundColor Cyan

try {
    Add-PnPListItem -List "QMS_DCOs" -Values @{
        Title           = "DCO-0002"
        DCO_Title       = "QMS Initial Release — Full Document Set (Cycle 2)"
        DCO_Phase       = "Draft"
        DCO_CRLink      = "CR-0002"
        DCO_Originator  = "Andre Butler"
        DCO_Docs        = $docs
        DCO_TrainingGate = $false
        DCO_ImplActivityRequired = $false
    }
    Write-Host "  [OK] DCO-0002 created with $($docs.Split(',').Count) documents" -ForegroundColor Green
} catch {
    Write-Host "  [FAIL] $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host "`nDone." -ForegroundColor Cyan
