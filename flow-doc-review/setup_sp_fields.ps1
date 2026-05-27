# IMP9177 — Setup SP fields and QMS_Config entries for document review gate
# Run after: Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

$siteUrl = "https://adbccro.sharepoint.com/sites/IMP9177"

# ── Seed QMS_Config with tenant_id ───────────────────────────────────────────
$tenantId = "729fb621-9df9-4895-9c2a-e0526a1a5912"

$existing = Get-PnPListItem -List "QMS_Config" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>tenant_id</Value></Eq></Where></Query></View>"
if ($existing) {
    Set-PnPListItem -List "QMS_Config" -Identity $existing.Id -Values @{ Cfg_Value = $tenantId }
    Write-Host "[UPDATE] tenant_id = $tenantId"
} else {
    Add-PnPListItem -List "QMS_Config" -Values @{ Title = "tenant_id"; Cfg_Value = $tenantId }
    Write-Host "[CREATE] tenant_id = $tenantId"
}

# ── Add fields to QMS_RoutingHistory ─────────────────────────────────────────
$list = "QMS_RoutingHistory"

function EnsureField($internalName, $displayName, $type) {
    $existing = Get-PnPField -List $list -Identity $internalName -ErrorAction SilentlyContinue
    if ($existing) { Write-Host "  [SKIP] $internalName already exists"; return }
    Add-PnPField -List $list -InternalName $internalName -DisplayName $displayName `
        -Type $type -AddToDefaultView $false
    Write-Host "  [OK]   $internalName added"
}

EnsureField "AL_DocID"     "Document ID"       "Text"
EnsureField "AL_Source"    "Event Source"      "Text"
EnsureField "AL_Timestamp" "Event Timestamp"   "DateTime"

Write-Host "`nAll SP setup complete."
