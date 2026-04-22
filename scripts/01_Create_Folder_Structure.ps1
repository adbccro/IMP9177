# =============================================================================
# IMP9177 — QMS PORTAL SHAREPOINT SETUP
# Script 01: Create Required Folder Structure
# Run FIRST before any file moves or application deployment
#
# Site: https://adbccro.sharepoint.com/sites/IMP9177
# Drive: b!TjkbwMYC5EiRJVzpwdOMb7D2R8-AVVZEpk9Rx3u59Y8oTtzCk9OzQZtdxpr5MeqU
#
# Prerequisites:
#   Install-Module PnP.PowerShell -Force
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
# =============================================================================

param(
    [string]$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177",
    [switch]$WhatIf = $false
)

$ErrorActionPreference = "Stop"
$WarningPreference     = "Continue"

# ── CONNECT ──────────────────────────────────────────────────────────────────
Write-Host "`n[IMP9177] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive

$web = Get-PnPWeb
Write-Host "[OK] Connected to: $($web.Title)" -ForegroundColor Green

# ── HELPER ───────────────────────────────────────────────────────────────────
# Resolve-PnPFolder creates the full path recursively — no -Library param needed
function Ensure-Folder {
    param([string]$LibraryName, [string]$FolderPath)
    $serverRelUrl = "/sites/IMP9177/$LibraryName/$FolderPath"
    try {
        $existing = Get-PnPFolder -Url $serverRelUrl -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Host "  [EXISTS]  $LibraryName/$FolderPath" -ForegroundColor DarkGray
        } else {
            if (-not $WhatIf) {
                Resolve-PnPFolder -SiteRelativePath "$LibraryName/$FolderPath" -ErrorAction Stop | Out-Null
            }
            Write-Host "  [CREATED] $LibraryName/$FolderPath" -ForegroundColor Green
        }
    } catch {
        Write-Warning "  [WARN]    Could not create $LibraryName/$FolderPath — $_"
    }
}

$lib = "Shared Documents"

# =============================================================================
# ZONE 1: DRAFT — Shared Documents/QMS/
# Already partially exists. Ensure all required subfolders are present.
# =============================================================================
Write-Host "`n[ZONE 1] Draft Zone — QMS/" -ForegroundColor Yellow

Ensure-Folder $lib "QMS/Documents"
Ensure-Folder $lib "QMS/Documents/Drafts"
Ensure-Folder $lib "QMS/Documents/Archive"
Ensure-Folder $lib "QMS/Forms"
Ensure-Folder $lib "QMS/Forms/Drafts"
Ensure-Folder $lib "QMS/Forms/Archive"
Ensure-Folder $lib "QMS/Records"
Ensure-Folder $lib "QMS/Records/Batch Records"
Ensure-Folder $lib "QMS/Records/Training"
Ensure-Folder $lib "QMS/Records/Supplier"
Ensure-Folder $lib "QMS/Records/CAPA"
Ensure-Folder $lib "QMS/Records/Environmental"
Ensure-Folder $lib "QMS/Records/Pest Control"
Ensure-Folder $lib "QMS/Records/Deviation"
Ensure-Folder $lib "QMS/Change Orders"
Ensure-Folder $lib "QMS/Change Orders/Archive"
Ensure-Folder $lib "QMS/Change Requests"
Ensure-Folder $lib "QMS/Archive"

# =============================================================================
# ZONE 2: PUBLISHED — Shared Documents/Published/QMS/
# Currently only has Forms, Quality Manual, Change Order subfolders.
# Missing: Documents (SOPs), Records.
# =============================================================================
Write-Host "`n[ZONE 2] Published Zone — Published/QMS/" -ForegroundColor Yellow

Ensure-Folder $lib "Published/QMS"
Ensure-Folder $lib "Published/QMS/Quality Manual"
Ensure-Folder $lib "Published/QMS/Documents"
Ensure-Folder $lib "Published/QMS/Documents/Archive"
Ensure-Folder $lib "Published/QMS/Forms"
Ensure-Folder $lib "Published/QMS/Forms/Archive"
Ensure-Folder $lib "Published/QMS/Records"
Ensure-Folder $lib "Published/QMS/Records/Batch Records"
Ensure-Folder $lib "Published/QMS/Records/Training"
Ensure-Folder $lib "Published/QMS/Records/CAPA"
Ensure-Folder $lib "Published/QMS/Change Order"
Ensure-Folder $lib "Published/QMS/Change Request"
Ensure-Folder $lib "Published/QMS/Archive"

# =============================================================================
# ZONE 3: OFFICIAL — Shared Documents/Official/QMS/
# Currently EMPTY. Create the full subfolder tree.
# This zone receives documents ONLY via the DCO Implemented transition.
# =============================================================================
Write-Host "`n[ZONE 3] Official Zone — Official/QMS/" -ForegroundColor Yellow

Ensure-Folder $lib "Official/QMS"
Ensure-Folder $lib "Official/QMS/Quality Manual"
Ensure-Folder $lib "Official/QMS/Documents"
Ensure-Folder $lib "Official/QMS/Documents/Archive"
Ensure-Folder $lib "Official/QMS/Forms"
Ensure-Folder $lib "Official/QMS/Forms/Archive"
Ensure-Folder $lib "Official/QMS/Records"
Ensure-Folder $lib "Official/QMS/Records/Batch Records"
Ensure-Folder $lib "Official/QMS/Records/Training"
Ensure-Folder $lib "Official/QMS/Records/Supplier"
Ensure-Folder $lib "Official/QMS/Records/CAPA"
Ensure-Folder $lib "Official/QMS/Records/Environmental"
Ensure-Folder $lib "Official/QMS/Records/Pest Control"
Ensure-Folder $lib "Official/QMS/Records/Deviation"
Ensure-Folder $lib "Official/QMS/Change Orders"
Ensure-Folder $lib "Official/QMS/Change Requests"
Ensure-Folder $lib "Official/QMS/Archive"

Write-Host "`n[IMP9177] Folder structure creation complete." -ForegroundColor Cyan
Write-Host "Run 02_Cleanup_Misplaced_Files.ps1 next.`n" -ForegroundColor White
