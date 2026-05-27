# ============================================================================
# IMP9177 — CR Feature Build: Deploy Runner
# Run from PowerShell in the SPFx project root.
# (C:\Users\andre\ADBCCRO Dropbox\Andre Butler\IMP9177\imp9177-spfx)
#
# PREREQUISITES:
#   - nvm use 22.14.0
#   - Connected to PnP: Connect-PnPOnline ... -Interactive
#   - Node script files from Claude in project root or scripts/ folder
# ============================================================================

Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host " IMP9177 — CR Feature Build Deploy" -ForegroundColor Cyan  
Write-Host "═══════════════════════════════════════════════════════" -ForegroundColor Cyan

# ── Step 1: Add missing SP fields ──────────────────────────────────────────
Write-Host "`n[1/5] Adding missing CR fields to QMS_ChangeRequests..." -ForegroundColor Yellow
& ".\add_cr_fields.ps1"

# ── Step 2: Apply main CR methods patch ────────────────────────────────────
Write-Host "`n[2/5] Applying CR feature patch to QmsPortalWebPart.ts..." -ForegroundColor Yellow
node .\patch_add_cr_feature.js
if ($LASTEXITCODE -ne 0) { Write-Host "[STOP] Patch failed. Fix issues before continuing." -ForegroundColor Red; exit 1 }

# ── Step 3: Apply buildShell CR panel patch ────────────────────────────────
Write-Host "`n[3/5] Patching buildShell() sc-cr panel..." -ForegroundColor Yellow
node .\patch_buildshell_cr.js
if ($LASTEXITCODE -ne 0) { Write-Host "[STOP] buildShell patch failed." -ForegroundColor Red; exit 1 }

# ── Step 4: Build ──────────────────────────────────────────────────────────
Write-Host "`n[4/5] Building SPFx solution..." -ForegroundColor Yellow
.\node_modules\.bin\heft clean
.\node_modules\.bin\heft build --production
if ($LASTEXITCODE -ne 0) { Write-Host "[STOP] Build failed. Check TypeScript errors." -ForegroundColor Red; exit 1 }
.\node_modules\.bin\heft package-solution --production

# ── Step 5: Deploy to SharePoint ───────────────────────────────────────────
Write-Host "`n[5/5] Deploying to SharePoint..." -ForegroundColor Yellow
Add-PnPApp -Path ".\sharepoint\solution\imp-9177-spfx.sppkg" `
  -Scope Site -Overwrite -Publish

# ── Step 6: Git push ───────────────────────────────────────────────────────
Write-Host "`nPushing to GitHub..." -ForegroundColor Yellow
$env:PATH += ";C:\Program Files\Git\bin;C:\Program Files\Git\cmd"
git add .
git commit -m "[CR] Full CR feature build — list panel, detail modal, approval flow, DCO link, audit trail"
git push origin main

Write-Host "`n[DONE] CR feature deployed." -ForegroundColor Green
Write-Host "Jira tickets IMP9177-23 through IMP9177-28 created." -ForegroundColor Gray
Write-Host "Mark tickets In Progress, then Done after smoke test." -ForegroundColor Gray
