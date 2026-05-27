# ============================================================================
# IMP9177 — Pre-build manual steps + smoke test
# ============================================================================

# ── STEP 0: AAD App Permission (do this BEFORE building flows) ───────────────
#
# 1. Go to: https://portal.azure.com
# 2. Azure Active Directory → App registrations
# 3. Find: IMP9177-PnP-Script (client ID: ba48ac81-6f23-43bd-9797-ec2866071102)
# 4. API permissions → Add a permission
# 5. APIs my organization uses → search "Office 365 Management APIs"
# 6. Application permissions → ActivityFeed.Read → Add
# 7. Grant admin consent (button at top of API permissions page)
# 8. Create a client secret:
#    Certificates & secrets → New client secret
#    Description: "IMP9177-Flow1-UAL"
#    Expires: 24 months
#    COPY THE VALUE IMMEDIATELY — only shown once
#    Store in Flow 1 parameter 'ual_client_secret'
#
# Also add to QMS_Config in SharePoint:
#   Title: "tenant_id"   | Cfg_Value: "<your-tenant-id>"
# Find tenant ID: Azure AD → Overview → Tenant ID (GUID)

# ── STEP 1: Get your tenant ID ───────────────────────────────────────────────
Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" `
  -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

$tenant = Get-PnPTenantId
Write-Host "Tenant ID: $tenant"
# Store this — needed for Flow 1 variables

# ── STEP 2: Verify UAL is enabled ────────────────────────────────────────────
# In Microsoft 365 admin center (admin.microsoft.com):
# Security & Compliance → Audit → Start recording
# OR: compliance.microsoft.com → Audit → turn on if not already on
# E5 license = audit is on by default with 1-year retention

# ── STEP 3: Seed QMS_Config with tenant_id ────────────────────────────────────
$tenantId = $tenant  # from step 1
Add-PnPListItem -List "QMS_Config" -Values @{
    Title = "tenant_id"
    Cfg_Value = $tenantId
}
Write-Host "tenant_id saved to QMS_Config"

# ── STEP 4: Add QMS_RoutingHistory fields needed by Flow 1 ───────────────────
# AL_DocID and AL_Source may not exist yet — add them:
$list = "QMS_RoutingHistory"

function EnsureField($internalName, $displayName, $type) {
    $existing = Get-PnPField -List $list -Identity $internalName -ErrorAction SilentlyContinue
    if ($existing) { Write-Host "  [SKIP] $internalName"; return }
    Add-PnPField -List $list -InternalName $internalName -DisplayName $displayName `
        -Type $type -AddToDefaultView $false
    Write-Host "  [OK]   $internalName"
}

EnsureField "AL_DocID"    "Document ID"   "Text"
EnsureField "AL_Source"   "Event Source"  "Text"
EnsureField "AL_Timestamp" "Event Timestamp" "DateTime"
Write-Host "QMS_RoutingHistory fields ready"

# ── SMOKE TEST: Simulate a DocumentOpened event manually ─────────────────────
# Use this to test Flow 3 without waiting 15-20 min for UAL latency.
# Run after building Flow 3 but before waiting for UAL events.

$dcoId = "DCO-0001"
$docId = "QM-001"
$approverEmail = "tinaqwork@gmail.com"  # or your test email

# Get last hash
$lastHash = (Get-PnPListItem -List "QMS_RoutingHistory" `
    -Query "<View><Query><Where><Eq><FieldRef Name='AL_DCOID'/><Value Type='Text'>$dcoId</Value></Eq></Where><OrderBy><FieldRef Name='Id' Ascending='False'/></OrderBy></Query><RowLimit>1</RowLimit></View>") |
    Select-Object -ExpandProperty FieldValues |
    Select-Object -ExpandProperty AL_Hash

$hash = [System.Convert]::ToBase64String(
    [System.Text.Encoding]::UTF8.GetBytes("DocumentOpened$dcoId$docId$approverEmail$(Get-Date -Format 'o')")
)

Add-PnPListItem -List "QMS_RoutingHistory" -Values @{
    Title        = "DocumentOpened-$dcoId-$docId-SMOKETEST"
    AL_EventType = "DocumentOpened"
    AL_DCOID     = $dcoId
    AL_DocID     = $docId
    AL_Actor     = $approverEmail
    AL_Timestamp = (Get-Date).ToUniversalTime().ToString("o")
    AL_Source    = "SMOKE_TEST_MANUAL"
    AL_Note      = "Manually injected for smoke test"
    AL_PrevHash  = if ($lastHash) { $lastHash } else { "GENESIS" }
    AL_Hash      = $hash
}
Write-Host "Smoke test DocumentOpened event written for $docId"

# ── TEST Flow 3 endpoint directly ────────────────────────────────────────────
# After building Flow 3, copy the HTTP trigger URL and paste below:
$flow3Url = "PASTE_FLOW3_HTTP_TRIGGER_URL_HERE"

# Test gate status:
$statusResponse = Invoke-RestMethod -Uri "$flow3Url&op=status&dcoId=$dcoId&approverEmail=$approverEmail" -Method GET
Write-Host "Gate status: $($statusResponse | ConvertTo-Json)"

# Test approve (after all docs opened):
$approveBody = @{
    dcoId = $dcoId
    approverEmail = $approverEmail
    action = "approve"
    reason = ""
} | ConvertTo-Json

$approveResponse = Invoke-RestMethod -Uri $flow3Url -Method POST `
    -Body $approveBody -ContentType "application/json"
Write-Host "Approve response: $($approveResponse | ConvertTo-Json)"

# ── VERIFY audit trail ────────────────────────────────────────────────────────
$auditEvents = Get-PnPListItem -List "QMS_RoutingHistory" `
    -Query "<View><Query><Where><Eq><FieldRef Name='AL_DCOID'/><Value Type='Text'>$dcoId</Value></Eq></Where></Query></View>"

Write-Host "`nAudit trail for $dcoId:"
$auditEvents | ForEach-Object {
    Write-Host "  [$($_.FieldValues.AL_EventType)] $($_.FieldValues.AL_Actor) — $($_.FieldValues.AL_Timestamp)"
}
# Expected output:
#   [ReviewEmailSent]  System — <timestamp>
#   [DocumentOpened]   tinaqwork@gmail.com — <timestamp>  (one per doc)
#   [DCOApproved]      tinaqwork@gmail.com — <timestamp>
