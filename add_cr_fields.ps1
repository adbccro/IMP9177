# add_cr_fields.ps1
# Adds CR workflow fields to QMS_ChangeRequests. Idempotent — safe to re-run.
#
# Usage:
#   Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
#   .\add_cr_fields.ps1

$ErrorActionPreference = "Stop"
$ListName = "QMS_ChangeRequests"

function EnsureField($list, $internalName, $type, $displayName, $extra = @{}) {
  $existing = Get-PnPField -List $list -Identity $internalName -ErrorAction SilentlyContinue
  if ($existing) {
    Write-Host "  [EXISTS] $internalName" -ForegroundColor Gray
    return
  }
  $params = @{
    List             = $list
    InternalName     = $internalName
    DisplayName      = $displayName
    Type             = $type
    AddToDefaultView = $true
  }
  foreach ($k in $extra.Keys) { $params[$k] = $extra[$k] }
  Add-PnPField @params | Out-Null
  Write-Host "  [ADDED]  $internalName ($type)" -ForegroundColor Green
}

Write-Host "Adding fields to $ListName..."

EnsureField $ListName "CR_Requestor"       "Text"     "CR Requestor"
EnsureField $ListName "CR_Category"        "Choice"   "CR Category" @{ Choices = @("Document","Process","Equipment","Supplier","Other") }
EnsureField $ListName "CR_AffectedDocs"    "Note"     "Affected Documents"
EnsureField $ListName "CR_SubmittedDate"   "DateTime" "Submitted Date"
EnsureField $ListName "CR_ApprovedDate"    "DateTime" "Approved Date"
EnsureField $ListName "CR_RejectedDate"    "DateTime" "Rejected Date"
EnsureField $ListName "CR_RejectionReason" "Note"     "Rejection Reason"

Write-Host "`nDone. Fields added to $ListName."
