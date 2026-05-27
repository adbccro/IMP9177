# IMP9177 — Create Flow 3 via Power Automate REST API
# Run with: ! .\flow-doc-review\create_flow3.ps1
# Requires: PnP PowerShell (already installed), PowerShell 7+

$tenantId     = "729fb621-9df9-4895-9c2a-e0526a1a5912"
$clientId     = "ba48ac81-6f23-43bd-9797-ec2866071102"
$clientSecret = "5e6c2b84-e0b7-45d5-9daf-2449cb04388d"
$paBase       = "https://api.flow.microsoft.com"
$apiVer       = "2016-11-01"

# ── Step 1: Acquire token ─────────────────────────────────────────────────────
Write-Host "`n[1] Acquiring OAuth token for Power Automate API..."

$tok = $null

# Try 1: client credentials (works if app has Flows.ReadWrite.All app permission)
try {
    $r = Invoke-RestMethod `
        -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
        -Method POST `
        -Body @{
            grant_type    = "client_credentials"
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = "https://service.flow.microsoft.com/.default"
        }
    $tok = $r.access_token
    Write-Host "  [OK] Client credentials token acquired"
} catch {
    Write-Host "  [INFO] CC auth failed (expected if app lacks Flows.ReadWrite.All): $($_.Exception.Message)"
}

# Try 2: PnP access token for PA resource (works if PnP session is active)
if (-not $tok) {
    try {
        Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" `
            -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
        # PnP doesn't natively issue PA tokens, but the connection proves auth works
        Write-Host "  [INFO] PnP connected — will attempt resource token via AAD"
    } catch {
        Write-Host "  [INFO] PnP connect skipped: $($_.Exception.Message)"
    }
}

# Try 3: Device code (interactive, works for delegated auth)
if (-not $tok) {
    Write-Host "`n  Falling back to device code (delegated auth)..."
    try {
        $dc = Invoke-RestMethod `
            -Uri "https://login.microsoftonline.com/729fb621-9df9-4895-9c2a-e0526a1a5912/oauth2/v2.0/devicecode" `
            -Method POST `
            -Body @{
                client_id = $clientId
                scope     = "https://service.flow.microsoft.com/.default"
            }

        Write-Host "`n  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
        Write-Host "  >>> Open:  https://aka.ms/devicelogin"
        Write-Host "  >>> Code:  $($dc.user_code)"
        Write-Host "  >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
        Write-Host "  Polling every $($dc.interval)s (expires in $($dc.expires_in)s)..."

        $deadline = (Get-Date).AddSeconds($dc.expires_in)
        while (-not $tok -and (Get-Date) -lt $deadline) {
            Start-Sleep -Seconds $dc.interval
            try {
                $pollR = Invoke-RestMethod `
                    -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" `
                    -Method POST `
                    -Body @{
                        grant_type  = "urn:ietf:params:oauth:grant-type:device_code"
                        client_id   = $clientId
                        device_code = $dc.device_code
                    } -ErrorAction Stop
                $tok = $pollR.access_token
                Write-Host "  [OK] Device code token acquired"
            } catch {
                $e = $_.ErrorDetails.Message | ConvertFrom-Json -ErrorAction SilentlyContinue
                if ($e?.error -eq "authorization_pending") { continue }
                if ($e?.error -eq "slow_down") { Start-Sleep -Seconds 5; continue }
                throw
            }
        }
    } catch {
        Write-Host "  [ERROR] Device code failed: $($_.Exception.Message)"
        Write-Host "  If error is AADSTS90100 (public client not enabled):"
        Write-Host "  Azure Portal → App registrations → IMP9177-PnP-Script"
        Write-Host "  → Authentication → Allow public client flows → Yes → Save"
        Write-Host "  Then re-run this script."
        exit 1
    }
}

if (-not $tok) { Write-Host "[FATAL] No token acquired. Exiting."; exit 1 }

$hdr = @{ Authorization = "Bearer $tok"; "Content-Type" = "application/json" }

# ── Step 2: Find default environment ─────────────────────────────────────────
Write-Host "`n[2] Discovering Power Automate environment..."
$envs   = Invoke-RestMethod "$paBase/providers/Microsoft.ProcessSimple/environments?api-version=$apiVer" -Headers $hdr
$envObj = $envs.value | Where-Object { $_.properties.isDefault -eq $true } | Select-Object -First 1
if (-not $envObj) { $envObj = $envs.value | Select-Object -First 1 }
$envId  = $envObj.name
Write-Host "  [OK] $envId ($($envObj.properties.displayName))"

# ── Step 3: Find existing SharePoint + Office 365 connections ────────────────
Write-Host "`n[3] Looking up existing connector connections..."
$allConns = (Invoke-RestMethod "$paBase/providers/Microsoft.ProcessSimple/environments/$envId/connections?api-version=$apiVer" -Headers $hdr).value

$spConn  = ($allConns | Where-Object { $_.properties.apiId -like "*sharepointonline*" } | Select-Object -First 1)
$o365Conn = ($allConns | Where-Object { $_.properties.apiId -like "*office365*" -and $_.properties.apiId -notlike "*groups*" -and $_.properties.apiId -notlike "*users*" } | Select-Object -First 1)

$spId    = if ($spConn)   { $spConn.name   } else { $null }
$o365Id  = if ($o365Conn) { $o365Conn.name } else { $null }

Write-Host "  SharePoint:   $spId"
Write-Host "  Office365:    $o365Id"

if (-not $spId)   { Write-Host "  [WARN] No SharePoint connection found — flow will need connection wired in UI after creation" }
if (-not $o365Id) { Write-Host "  [WARN] No Office 365 connection found — rejection email action will need connection in UI" }

# ── Step 4: Build flow definition ────────────────────────────────────────────
Write-Host "`n[4] Building flow definition..."

# Helper: build SP action inputs
function SpInputs($method, $pathSuffix, $extraBody = $null, $queries = $null) {
    $h = @{
        host   = @{ connection = @{ name = "@parameters('`$connections')['sp']['connectionId']" } }
        method = $method
        path   = "/datasets/@{encodeURIComponent(encodeURIComponent('https://adbccro.sharepoint.com/sites/IMP9177'))}/tables/@{encodeURIComponent(encodeURIComponent('$pathSuffix'))}/items"
    }
    if ($queries)    { $h.queries  = $queries }
    if ($extraBody)  { $h.body     = $extraBody }
    return $h
}

$flowDef = @{
    '$schema'      = "https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#"
    contentVersion = "1.0.0.0"
    parameters     = @{
        '$connections' = @{ defaultValue = @{}; type = "Object" }
    }
    triggers = @{
        manual = @{
            type   = "Request"
            kind   = "Http"
            inputs = @{
                schema = @{
                    type       = "object"
                    properties = @{
                        dcoId         = @{ type = "string" }
                        approverEmail = @{ type = "string" }
                        action        = @{ type = "string" }
                        reason        = @{ type = "string" }
                    }
                }
            }
        }
    }
    actions = @{

        Route_by_op = @{
            type       = "If"
            runAfter   = @{}
            expression = @{ and = @( @{ equals = @( "@{triggerOutputs()?['queries']?['op']}", "status" ) } ) }

            # ── YES branch: GET gate status ──────────────────────────
            actions = @{

                GS_Get_DCO = @{
                    type     = "ApiConnection"
                    runAfter = @{}
                    inputs   = SpInputs "get" "QMS_DCOs" -queries @{
                        '$filter' = "Title eq '@{triggerOutputs()?[''queries'']?[''dcoId'']}'"
                        '$top'    = "1"
                        '$select' = "Title,DCO_Title,DCO_Docs,ID"
                    }
                }

                GS_Parse_Docs = @{
                    type     = "Compose"
                    runAfter = @{ GS_Get_DCO = @("Succeeded") }
                    inputs   = "@split(replace(first(body('GS_Get_DCO')?['value'])?['DCO_Docs'],' ',''),',')"
                }

                GS_Get_Opened = @{
                    type     = "ApiConnection"
                    runAfter = @{ GS_Parse_Docs = @("Succeeded") }
                    inputs   = SpInputs "get" "QMS_RoutingHistory" -queries @{
                        '$filter' = "AL_EventType eq 'DocumentOpened' and AL_DCOID eq '@{triggerOutputs()?[''queries'']?[''dcoId'']}' and AL_Actor eq '@{triggerOutputs()?[''queries'']?[''approverEmail'']}'"
                        '$top'    = "100"
                    }
                }

                GS_Respond = @{
                    type     = "Response"
                    kind     = "Http"
                    runAfter = @{ GS_Get_Opened = @("Succeeded") }
                    inputs   = @{
                        statusCode = 200
                        headers    = @{ "Content-Type" = "application/json" }
                        body       = @{
                            dcoId           = "@{triggerOutputs()?['queries']?['dcoId']}"
                            dcoTitle        = "@{first(body('GS_Get_DCO')?['value'])?['DCO_Title']}"
                            totalDocs       = "@{length(outputs('GS_Parse_Docs'))}"
                            openedDocs      = "@{length(body('GS_Get_Opened')?['value'])}"
                            gateOpen        = "@{greaterOrEquals(length(body('GS_Get_Opened')?['value']),length(outputs('GS_Parse_Docs')))}"
                            remainingCount  = "@{sub(length(outputs('GS_Parse_Docs')),length(body('GS_Get_Opened')?['value']))}"
                            progressColor   = "@{if(greaterOrEquals(length(body('GS_Get_Opened')?['value']),length(outputs('GS_Parse_Docs'))),'Good','Warning')}"
                            progressMessage = "@{if(greaterOrEquals(length(body('GS_Get_Opened')?['value']),length(outputs('GS_Parse_Docs'))),'All documents reviewed','Open remaining documents before approving')}"
                            approverEmail   = "@{triggerOutputs()?['queries']?['approverEmail']}"
                        }
                    }
                }
            }

            # ── NO branch: POST approve/reject ───────────────────────
            else = @{
                actions = @{

                    AP_Parse = @{
                        type     = "ParseJson"
                        runAfter = @{}
                        inputs   = @{
                            content = "@triggerBody()"
                            schema  = @{
                                type       = "object"
                                properties = @{
                                    dcoId         = @{ type = "string" }
                                    approverEmail = @{ type = "string" }
                                    action        = @{ type = "string" }
                                    reason        = @{ type = "string" }
                                }
                            }
                        }
                    }

                    AP_Get_Opened = @{
                        type     = "ApiConnection"
                        runAfter = @{ AP_Parse = @("Succeeded") }
                        inputs   = SpInputs "get" "QMS_RoutingHistory" -queries @{
                            '$filter' = "AL_EventType eq 'DocumentOpened' and AL_DCOID eq '@{body(''AP_Parse'')?[''dcoId'']}' and AL_Actor eq '@{body(''AP_Parse'')?[''approverEmail'']}'"
                            '$top'    = "100"
                        }
                    }

                    AP_Get_DCO = @{
                        type     = "ApiConnection"
                        runAfter = @{ AP_Parse = @("Succeeded") }
                        inputs   = SpInputs "get" "QMS_DCOs" -queries @{
                            '$filter' = "Title eq '@{body(''AP_Parse'')?[''dcoId'']}'"
                            '$top'    = "1"
                            '$select' = "Title,DCO_Title,DCO_Docs,DCO_Originator,ID"
                        }
                    }

                    AP_Parse_Docs = @{
                        type     = "Compose"
                        runAfter = @{ AP_Get_DCO = @("Succeeded") }
                        inputs   = "@split(replace(first(body('AP_Get_DCO')?['value'])?['DCO_Docs'],' ',''),',')"
                    }

                    AP_Gate_Check = @{
                        type       = "If"
                        runAfter   = @{ AP_Get_Opened = @("Succeeded"); AP_Parse_Docs = @("Succeeded") }
                        expression = @{ and = @( @{ greaterOrEquals = @( "@length(body('AP_Get_Opened')?['value'])", "@length(outputs('AP_Parse_Docs'))" ) } ) }

                        actions = @{
                            AP_Approve_Or_Reject = @{
                                type       = "If"
                                runAfter   = @{}
                                expression = @{ and = @( @{ equals = @( "@body('AP_Parse')?['action']", "approve" ) } ) }

                                actions = @{
                                    APPR_Patch_DCO = @{
                                        type     = "ApiConnection"
                                        runAfter = @{}
                                        inputs   = @{
                                            host    = @{ connection = @{ name = "@parameters('`$connections')['sp']['connectionId']" } }
                                            method  = "patch"
                                            path    = "/datasets/@{encodeURIComponent(encodeURIComponent('https://adbccro.sharepoint.com/sites/IMP9177'))}/tables/@{encodeURIComponent(encodeURIComponent('QMS_DCOs'))}/items/@{first(body('AP_Get_DCO')?['value'])?['ID']}"
                                            body    = @{ DCO_Phase = "Approved" }
                                        }
                                    }
                                    APPR_Get_ApprRow = @{
                                        type     = "ApiConnection"
                                        runAfter = @{ APPR_Patch_DCO = @("Succeeded") }
                                        inputs   = SpInputs "get" "QMS_DCOApprovals" -queries @{
                                            '$filter' = "Appr_DCOID eq '@{body(''AP_Parse'')?[''dcoId'']}' and Sig_ApproverEmail eq '@{body(''AP_Parse'')?[''approverEmail'']}'"
                                            '$top'    = "1"
                                        }
                                    }
                                    APPR_Patch_Approval = @{
                                        type     = "ApiConnection"
                                        runAfter = @{ APPR_Get_ApprRow = @("Succeeded") }
                                        inputs   = @{
                                            host   = @{ connection = @{ name = "@parameters('`$connections')['sp']['connectionId']" } }
                                            method = "patch"
                                            path   = "/datasets/@{encodeURIComponent(encodeURIComponent('https://adbccro.sharepoint.com/sites/IMP9177'))}/tables/@{encodeURIComponent(encodeURIComponent('QMS_DCOApprovals'))}/items/@{first(body('APPR_Get_ApprRow')?['value'])?['ID']}"
                                            body   = @{ Appr_Status = "Approved"; Appr_SignedDate = "@{utcNow()}" }
                                        }
                                    }
                                    APPR_Write_Audit = @{
                                        type     = "ApiConnection"
                                        runAfter = @{ APPR_Patch_Approval = @("Succeeded") }
                                        inputs   = SpInputs "post" "QMS_RoutingHistory" -extraBody @{
                                            Title        = "DCOApproved-@{body('AP_Parse')?['dcoId']}"
                                            AL_EventType = "DCOApproved"
                                            AL_DCOID     = "@{body('AP_Parse')?['dcoId']}"
                                            AL_Actor     = "@{body('AP_Parse')?['approverEmail']}"
                                            AL_Note      = "Approved via Adaptive Card — all documents reviewed"
                                            AL_Timestamp = "@{utcNow()}"
                                            AL_Source    = "AdaptiveCard_ActionExecute"
                                        }
                                    }
                                    APPR_Respond = @{
                                        type     = "Response"
                                        kind     = "Http"
                                        runAfter = @{ APPR_Write_Audit = @("Succeeded") }
                                        inputs   = @{
                                            statusCode = 200
                                            headers    = @{ "Content-Type" = "application/json" }
                                            body       = @{ success = $true; message = "DCO approved"; dcoId = "@{body('AP_Parse')?['dcoId']}" }
                                        }
                                    }
                                }

                                else = @{
                                    actions = @{
                                        REJ_Check_Reason = @{
                                            type       = "If"
                                            runAfter   = @{}
                                            expression = @{ and = @( @{ not = @{ equals = @( "@body('AP_Parse')?['reason']", "" ) } } ) }
                                            actions    = @{
                                                REJ_Patch_DCO = @{
                                                    type     = "ApiConnection"
                                                    runAfter = @{}
                                                    inputs   = @{
                                                        host   = @{ connection = @{ name = "@parameters('`$connections')['sp']['connectionId']" } }
                                                        method = "patch"
                                                        path   = "/datasets/@{encodeURIComponent(encodeURIComponent('https://adbccro.sharepoint.com/sites/IMP9177'))}/tables/@{encodeURIComponent(encodeURIComponent('QMS_DCOs'))}/items/@{first(body('AP_Get_DCO')?['value'])?['ID']}"
                                                        body   = @{ DCO_Phase = "Rejected" }
                                                    }
                                                }
                                                REJ_Write_Audit = @{
                                                    type     = "ApiConnection"
                                                    runAfter = @{ REJ_Patch_DCO = @("Succeeded") }
                                                    inputs   = SpInputs "post" "QMS_RoutingHistory" -extraBody @{
                                                        Title        = "DCORejected-@{body('AP_Parse')?['dcoId']}"
                                                        AL_EventType = "DCORejected"
                                                        AL_DCOID     = "@{body('AP_Parse')?['dcoId']}"
                                                        AL_Actor     = "@{body('AP_Parse')?['approverEmail']}"
                                                        AL_Note      = "@{body('AP_Parse')?['reason']}"
                                                        AL_Timestamp = "@{utcNow()}"
                                                        AL_Source    = "AdaptiveCard_ActionExecute"
                                                    }
                                                }
                                                REJ_Email = @{
                                                    type     = "ApiConnection"
                                                    runAfter = @{ REJ_Write_Audit = @("Succeeded") }
                                                    inputs   = @{
                                                        host   = @{ connection = @{ name = "@parameters('`$connections')['o365']['connectionId']" } }
                                                        method = "post"
                                                        path   = "/v2/Mail"
                                                        body   = @{
                                                            To         = "@{first(body('AP_Get_DCO')?['value'])?['DCO_Originator']}"
                                                            Subject    = "[IMP9177] @{body('AP_Parse')?['dcoId']} — Rejected"
                                                            Body       = "<p>DCO <strong>@{body('AP_Parse')?['dcoId']}</strong> was rejected.</p><p><strong>Reason:</strong> @{body('AP_Parse')?['reason']}</p>"
                                                            Importance = "Normal"
                                                        }
                                                    }
                                                }
                                                REJ_Respond = @{
                                                    type     = "Response"
                                                    kind     = "Http"
                                                    runAfter = @{ REJ_Email = @("Succeeded") }
                                                    inputs   = @{
                                                        statusCode = 200
                                                        headers    = @{ "Content-Type" = "application/json" }
                                                        body       = @{ success = $true; message = "DCO rejected"; dcoId = "@{body('AP_Parse')?['dcoId']}" }
                                                    }
                                                }
                                            }
                                            else = @{
                                                actions = @{
                                                    REJ_No_Reason = @{
                                                        type     = "Response"
                                                        kind     = "Http"
                                                        runAfter = @{}
                                                        inputs   = @{
                                                            statusCode = 400
                                                            headers    = @{ "Content-Type" = "application/json" }
                                                            body       = @{ error = "Rejection reason is required" }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        else = @{
                            actions = @{
                                AP_Gate_Blocked = @{
                                    type     = "Response"
                                    kind     = "Http"
                                    runAfter = @{}
                                    inputs   = @{
                                        statusCode = 400
                                        headers    = @{ "Content-Type" = "application/json" }
                                        body       = @{
                                            error    = "Not all documents reviewed"
                                            opened   = "@{length(body('AP_Get_Opened')?['value'])}"
                                            required = "@{length(outputs('AP_Parse_Docs'))}"
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
            }
        }
    }
    outputs = @{}
}

# ── Step 5: Build connection references ───────────────────────────────────────
$connRefs = @{
    sp = @{
        connectionName = if ($spId)   { $spId   } else { "" }
        source         = "Invoker"
        id             = "/providers/Microsoft.PowerApps/apis/shared_sharepointonline"
        tier           = "NotSpecified"
    }
    o365 = @{
        connectionName = if ($o365Id) { $o365Id } else { "" }
        source         = "Invoker"
        id             = "/providers/Microsoft.PowerApps/apis/shared_office365"
        tier           = "NotSpecified"
    }
}

# ── Step 6: POST to create flow ───────────────────────────────────────────────
Write-Host "`n[5] Creating flow via PA REST API..."

$bodyObj = @{
    properties = @{
        displayName          = "IMP9177 — Approval Gate Endpoint"
        definition           = $flowDef
        connectionReferences = $connRefs
    }
}

$bodyJson = $bodyObj | ConvertTo-Json -Depth 100 -Compress
$createUri = "$paBase/providers/Microsoft.ProcessSimple/environments/$envId/flows?api-version=$apiVer"

try {
    $created = Invoke-RestMethod -Uri $createUri -Method POST -Headers $hdr -Body $bodyJson
    $flowId  = $created.name
    Write-Host "  [OK] Flow created: $flowId"
    Write-Host "  State: $($created.properties.state)"
} catch {
    $errText = $_.ErrorDetails.Message
    Write-Host "  [ERROR] Create failed: $($_.Exception.Message)"
    Write-Host "  Body: $errText"
    exit 1
}

# ── Step 7: Retrieve trigger URL ──────────────────────────────────────────────
Write-Host "`n[6] Retrieving HTTP trigger URL..."
$cbUri = "$paBase/providers/Microsoft.ProcessSimple/environments/$envId/flows/$flowId/triggers/manual/listCallbackUrl?api-version=$apiVer"

try {
    $cb = Invoke-RestMethod -Uri $cbUri -Method POST -Headers $hdr
    $triggerUrl = $cb.value

    Write-Host "`n============================================================"
    Write-Host "  FLOW 3 TRIGGER URL:"
    Write-Host "  $triggerUrl"
    Write-Host "============================================================"
    Write-Host "  Flow ID:  $flowId"
    Write-Host "  Env ID:   $envId"
    Write-Host "`n  Next: paste this URL back to Claude to wire up Flow 2 + Adaptive Card."

    # Save trigger URL to file for reference
    $triggerUrl | Out-File (Join-Path $PSScriptRoot "flow3_trigger_url.txt") -Encoding UTF8 -NoNewline
    Write-Host "  (Also saved to flow-doc-review\flow3_trigger_url.txt)"
} catch {
    Write-Host "  [WARN] Could not retrieve trigger URL: $($_.Exception.Message)"
    Write-Host "  Flow was created (ID: $flowId). Open it in make.powerautomate.com"
    Write-Host "  and copy the HTTP trigger URL from the first action."
}
