# save_pj_to_sp.ps1
# Saves pj_results.json P&J values to QMS_DCOs via PnP PowerShell.
# Assumes Connect-PnPOnline already called before running this script.

$ErrorActionPreference = "Stop"

# ── Load results ──────────────────────────────────────────────────────────────

$raw     = Get-Content -Path "pj_results.json" -Raw
$results = $raw | ConvertFrom-Json

# ── Build DCO_DocPurposes JSON for each DCO ───────────────────────────────────

$dco1Keys = @(
  'QM-001','SOP-QMS-001','SOP-QMS-002','SOP-QMS-003','SOP-PRD-108',
  'SOP-RCL-321','SOP-PRD-432','SOP-FRS-549','FM-001','FM-002',
  'FM-003','FM-025','FM-027','FM-030','FPS-001'
)

$dco2Keys = @(
  'SOP-SUP-001','SOP-SUP-002','SOP-FS-001','SOP-FS-002','SOP-FS-003',
  'SOP-FS-004','SOP-PC-001','FM-004','FM-005','FM-006',
  'FM-007','FM-ALG','FM-008'
)

function Build-PJObject ($source, [string[]]$keys) {
  $obj = [ordered]@{}
  foreach ($k in $keys) {
    $entry = $source.$k
    if ($null -eq $entry) {
      Write-Warning "  Key not found in pj_results.json: $k"
      $obj[$k] = [ordered]@{ purpose = ''; justification = '' }
    } else {
      $obj[$k] = [ordered]@{
        purpose       = $entry.purpose
        justification = $entry.justification
      }
    }
  }
  return $obj
}

$dco1Obj  = Build-PJObject $results.'DCO-0001' $dco1Keys
$dco2Obj  = Build-PJObject $results.'DCO-0002' $dco2Keys

$dco1Json = $dco1Obj | ConvertTo-Json -Depth 5 -Compress
$dco2Json = $dco2Obj | ConvertTo-Json -Depth 5 -Compress

Write-Host "DCO-0001 JSON length: $($dco1Json.Length) chars ($($dco1Keys.Count) docs)"
Write-Host "DCO-0002 JSON length: $($dco2Json.Length) chars ($($dco2Keys.Count) docs)"

# ── PATCH both records ────────────────────────────────────────────────────────

Write-Host "`nPatching QMS_DCOs item 1 (DCO-0001)..." -NoNewline
Set-PnPListItem -List "QMS_DCOs" -Identity 1 -Values @{ DCO_DocPurposes = $dco1Json } | Out-Null
Write-Host " OK"

Write-Host "Patching QMS_DCOs item 4 (DCO-0002)..." -NoNewline
Set-PnPListItem -List "QMS_DCOs" -Identity 4 -Values @{ DCO_DocPurposes = $dco2Json } | Out-Null
Write-Host " OK"

Write-Host "`nBoth DCO records updated successfully."
