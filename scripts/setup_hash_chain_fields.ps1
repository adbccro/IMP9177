# setup_hash_chain_fields.ps1
# Adds Part 11 hash-chain columns to QMS_DCOApprovals, QMS_RoutingHistory,
# and QMS_TrainingCompletions on the IMP9177 SharePoint site.
#
# Run after Connect-PnPOnline:
#   Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" `
#     -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
#   .\scripts\setup_hash_chain_fields.ps1

$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177"

function Add-FieldIfMissing {
    param(
        [string]$ListName,
        [string]$FieldName,
        [string]$FieldType,    # "Text" or "Note"
        [string]$DisplayName
    )
    try {
        $existing = Get-PnPField -List $ListName -Identity $FieldName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  [SKIP] $ListName.$FieldName already exists" -ForegroundColor Yellow
        } else {
            Add-PnPField -List $ListName -InternalName $FieldName -DisplayName $DisplayName `
                -Type $FieldType -AddToDefaultView:$false | Out-Null
            Write-Host "  [OK]   $ListName.$FieldName added ($FieldType)" -ForegroundColor Green
        }
    } catch {
        Write-Host "  [ERR]  $ListName.$FieldName — $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Host "`n=== QMS_DCOApprovals — Signature hash chain fields ===" -ForegroundColor Cyan
Add-FieldIfMissing -ListName "QMS_DCOApprovals" -FieldName "Sig_Hash"            -FieldType "Text" -DisplayName "Sig_Hash"
Add-FieldIfMissing -ListName "QMS_DCOApprovals" -FieldName "Sig_PrevHash"        -FieldType "Text" -DisplayName "Sig_PrevHash"
Add-FieldIfMissing -ListName "QMS_DCOApprovals" -FieldName "Sig_Timestamp_UTC"   -FieldType "Text" -DisplayName "Sig_Timestamp_UTC"
Add-FieldIfMissing -ListName "QMS_DCOApprovals" -FieldName "Sig_IPAddress"       -FieldType "Text" -DisplayName "Sig_IPAddress"
Add-FieldIfMissing -ListName "QMS_DCOApprovals" -FieldName "Sig_RevocationNote"  -FieldType "Note" -DisplayName "Sig_RevocationNote"

Write-Host "`n=== QMS_RoutingHistory — Activity log hash chain fields ===" -ForegroundColor Cyan
Add-FieldIfMissing -ListName "QMS_RoutingHistory" -FieldName "AL_Hash"     -FieldType "Text" -DisplayName "AL_Hash"
Add-FieldIfMissing -ListName "QMS_RoutingHistory" -FieldName "AL_PrevHash" -FieldType "Text" -DisplayName "AL_PrevHash"

Write-Host "`n=== QMS_TrainingCompletions — Training cert hash field ===" -ForegroundColor Cyan
Add-FieldIfMissing -ListName "QMS_TrainingCompletions" -FieldName "TC_Hash" -FieldType "Text" -DisplayName "TC_Hash"

Write-Host "`nDone. All hash-chain fields provisioned.`n" -ForegroundColor Cyan
