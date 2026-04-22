# ============================================================
# IMP9177 SharePoint List Provisioning Script
# Site: https://adbccro.sharepoint.com/sites/IMP9177
# Run AFTER: Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
# Run AFTER: Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -UseWebLogin
# ============================================================

$site = "https://adbccro.sharepoint.com/sites/IMP9177"
$ErrorActionPreference = "Continue"

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " IMP9177 SharePoint List Provisioning" -ForegroundColor Cyan
Write-Host " Site: $site" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

function Ensure-List {
    param([string]$Title, [string]$Description = "")
    $existing = Get-PnPList -Identity $Title -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        New-PnPList -Title $Title -Template GenericList -EnableVersioning
        Write-Host "  [CREATED] $Title" -ForegroundColor Green
    } else {
        Write-Host "  [EXISTS]  $Title" -ForegroundColor Yellow
    }
}

function Ensure-Field {
    param([string]$List, [string]$Name, [string]$Type, [string[]]$Choices = @(), [bool]$Required = $false)
    $existing = Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        switch ($Type) {
            "Text"     { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Text -AddToDefaultView | Out-Null }
            "Note"     { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Note -AddToDefaultView | Out-Null }
            "DateTime" { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type DateTime -AddToDefaultView | Out-Null }
            "Number"   { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Number -AddToDefaultView | Out-Null }
            "Boolean"  { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Boolean -AddToDefaultView | Out-Null }
            "Choice"   { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Choice -Choices $Choices -AddToDefaultView | Out-Null }
            "URL"      { Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type URL -AddToDefaultView | Out-Null }
        }
        Write-Host "    + $Name ($Type)" -ForegroundColor DarkGreen
    }
}

# ============================================================
# 1. MRS_Documents
# ============================================================
Write-Host "`n[1/11] MRS_Documents" -ForegroundColor Cyan
Ensure-List "MRS_Documents" "IMP9177 Document Registry"
Ensure-Field "MRS_Documents" "DocNumber"   "Text"
Ensure-Field "MRS_Documents" "DocTitle"    "Text"
Ensure-Field "MRS_Documents" "DocType"     "Choice" @("SOP","Form","Policy","Plan","Quality Manual","DCO","Change Request","Record")
Ensure-Field "MRS_Documents" "DocWeek"     "Choice" @("W1","W2","W3","W4","Existing")
Ensure-Field "MRS_Documents" "DocRev"      "Text"
Ensure-Field "MRS_Documents" "DocStatus"   "Choice" @("Not Started","ADB Drafting","ADB Review","Published","Pending Feedback","PAST DUE -- Sign-Off Overdue","Signed Off","Superseded")
Ensure-Field "MRS_Documents" "ADBComplete" "Choice" @("Yes","No","In Progress")
Ensure-Field "MRS_Documents" "DocURL"      "URL"
Ensure-Field "MRS_Documents" "DocNotes"    "Note"

# ============================================================
# 2. MRS_GapRegistry
# ============================================================
Write-Host "`n[2/11] MRS_GapRegistry" -ForegroundColor Cyan
Ensure-List "MRS_GapRegistry" "IMP9177 Gap Registry"
Ensure-Field "MRS_GapRegistry" "GapID"       "Text"
Ensure-Field "MRS_GapRegistry" "GapSeries"   "Choice" @("GAP-DC","GAP-TR","GAP-PC","GAP-SUP","GAP-LAB","GAP-MFG","GAP-REC")
Ensure-Field "MRS_GapRegistry" "GapDesc"     "Note"
Ensure-Field "MRS_GapRegistry" "GapCFR"      "Text"
Ensure-Field "MRS_GapRegistry" "GapSeverity" "Choice" @("Critical","Major","Minor")
Ensure-Field "MRS_GapRegistry" "GapOwner"    "Choice" @("ADB","3H","Joint")
Ensure-Field "MRS_GapRegistry" "GapStatus"   "Choice" @("Open","In Progress","Closed","Deferred")
Ensure-Field "MRS_GapRegistry" "GapNotes"    "Note"

# ============================================================
# 3. MRS_SupplierRegistry
# ============================================================
Write-Host "`n[3/11] MRS_SupplierRegistry" -ForegroundColor Cyan
Ensure-List "MRS_SupplierRegistry" "IMP9177 Supplier Registry"
Ensure-Field "MRS_SupplierRegistry" "SUPNumber"      "Text"
Ensure-Field "MRS_SupplierRegistry" "SupplierName"   "Text"
Ensure-Field "MRS_SupplierRegistry" "SupplierCat"    "Choice" @("Raw Materials","Packaging Materials","Laboratory Consumables","Services","Other")
Ensure-Field "MRS_SupplierRegistry" "SUPQualStatus"  "Choice" @("PENDING","APPROVED","CONDITIONAL","DISQUALIFIED")
Ensure-Field "MRS_SupplierRegistry" "SUPContact"     "Text"
Ensure-Field "MRS_SupplierRegistry" "SUPCoAOnFile"   "Choice" @("Yes","No","Partial")
Ensure-Field "MRS_SupplierRegistry" "SUPNotes"       "Note"

# ============================================================
# 4. MRS_IngredientRegistry
# ============================================================
Write-Host "`n[4/11] MRS_IngredientRegistry" -ForegroundColor Cyan
Ensure-List "MRS_IngredientRegistry" "IMP9177 Ingredient Registry"
Ensure-Field "MRS_IngredientRegistry" "INGNumber"    "Text"
Ensure-Field "MRS_IngredientRegistry" "INGName"      "Text"
Ensure-Field "MRS_IngredientRegistry" "INGSupplier"  "Text"
Ensure-Field "MRS_IngredientRegistry" "INGCoAStatus" "Choice" @("CoA on File","CoA Missing","CRITICAL -- No CoA Active Production","Spec Required","Exempted")
Ensure-Field "MRS_IngredientRegistry" "INGLotNum"    "Text"
Ensure-Field "MRS_IngredientRegistry" "INGNotes"     "Note"

# ============================================================
# 5. MRS_CAPALog
# ============================================================
Write-Host "`n[5/11] MRS_CAPALog" -ForegroundColor Cyan
Ensure-List "MRS_CAPALog" "IMP9177 CAPA Log"
Ensure-Field "MRS_CAPALog" "CAPANumber"  "Text"
Ensure-Field "MRS_CAPALog" "CAPAType"    "Choice" @("CAPA","Correction","Preventive Action","OOS Investigation")
Ensure-Field "MRS_CAPALog" "CAPADesc"    "Note"
Ensure-Field "MRS_CAPALog" "CAPAStatus"  "Choice" @("OPEN -- In Progress","OPEN -- Pending Verification","Closed -- Effective","Closed -- Ineffective")
Ensure-Field "MRS_CAPALog" "CAPAOwner"   "Text"
Ensure-Field "MRS_CAPALog" "CAPADueDate" "DateTime"
Ensure-Field "MRS_CAPALog" "CAPANotes"   "Note"

# ============================================================
# 6. RAID_Actions
# ============================================================
Write-Host "`n[6/11] RAID_Actions" -ForegroundColor Cyan
Ensure-List "RAID_Actions" "IMP9177 Action Register"
Ensure-Field "RAID_Actions" "ActionCategory"  "Choice" @("Document Control","Compliance","Supplier","Training","Facility","IT/Systems","Administration","QA")
Ensure-Field "RAID_Actions" "ActionDesc"      "Note"
Ensure-Field "RAID_Actions" "ActionOwner"     "Text"
Ensure-Field "RAID_Actions" "ActionPriority"  "Choice" @("Critical","High","Medium","Low")
Ensure-Field "RAID_Actions" "ActionStatus"    "Choice" @("Open","Open -- Urgent","Open -- Critical","BLOCKING","PAST DUE","In Progress","Complete","Closed","Deferred")
Ensure-Field "RAID_Actions" "DueDate"         "DateTime"
Ensure-Field "RAID_Actions" "RaisedInMeeting" "Text"
Ensure-Field "RAID_Actions" "CompletionDate"  "DateTime"
Ensure-Field "RAID_Actions" "BlockedBy"       "Text"
Ensure-Field "RAID_Actions" "UpdateNotes"     "Note"

# ============================================================
# 7. RAID_Issues
# ============================================================
Write-Host "`n[7/11] RAID_Issues" -ForegroundColor Cyan
Ensure-List "RAID_Issues" "IMP9177 Issue Register"
Ensure-Field "RAID_Issues" "IssueDescription"      "Note"
Ensure-Field "RAID_Issues" "RegulationRef"         "Text"
Ensure-Field "RAID_Issues" "IssueSeverity"         "Choice" @("Critical","High","Major","Minor")
Ensure-Field "RAID_Issues" "IssueStatus"           "Choice" @("Open -- Critical","Open -- Urgent","Open","Partially Mitigated","Closed","Monitoring")
Ensure-Field "RAID_Issues" "IssueOwner"            "Text"
Ensure-Field "RAID_Issues" "MitigationDescription" "Note"
Ensure-Field "RAID_Issues" "ResolutionDescription" "Note"
Ensure-Field "RAID_Issues" "ClosureDate"           "DateTime"
Ensure-Field "RAID_Issues" "LinkedCAPARef"         "Text"

# ============================================================
# 8. RAID_Decisions
# ============================================================
Write-Host "`n[8/11] RAID_Decisions" -ForegroundColor Cyan
Ensure-List "RAID_Decisions" "IMP9177 Decision Log"
Ensure-Field "RAID_Decisions" "DecisionDate"      "DateTime"
Ensure-Field "RAID_Decisions" "MeetingRef"        "Text"
Ensure-Field "RAID_Decisions" "DecisionDesc"      "Note"
Ensure-Field "RAID_Decisions" "MadeBy"            "Text"
Ensure-Field "RAID_Decisions" "LinkedActionRefs"  "Text"

# ============================================================
# 9. RAID_Meetings
# ============================================================
Write-Host "`n[9/11] RAID_Meetings" -ForegroundColor Cyan
Ensure-List "RAID_Meetings" "IMP9177 Meeting Log"
Ensure-Field "RAID_Meetings" "MeetingDate"       "DateTime"
Ensure-Field "RAID_Meetings" "MeetingType"       "Choice" @("Project Review","Kickoff","Status Update","Ad Hoc","Training")
Ensure-Field "RAID_Meetings" "Platform"          "Choice" @("Teams","Zoom","In Person","Phone","Email")
Ensure-Field "RAID_Meetings" "Duration"          "Number"
Ensure-Field "RAID_Meetings" "AgendaSummary"     "Note"
Ensure-Field "RAID_Meetings" "MeetingNotes"      "Note"
Ensure-Field "RAID_Meetings" "ActionsRaised"     "Number"
Ensure-Field "RAID_Meetings" "DecisionsMade"     "Number"
Ensure-Field "RAID_Meetings" "NextMeetingTarget" "DateTime"
Ensure-Field "RAID_Meetings" "MinutesURL"        "URL"

# ============================================================
# 10. PM_Milestones
# ============================================================
Write-Host "`n[10/11] PM_Milestones" -ForegroundColor Cyan
Ensure-List "PM_Milestones" "IMP9177 PM Milestones"
Ensure-Field "PM_Milestones" "MilestonePhase"  "Choice" @("Phase 1","Phase 2A","Phase 2B","Phase 3","Phase 4")
Ensure-Field "PM_Milestones" "MilestoneDesc"   "Note"
Ensure-Field "PM_Milestones" "MilestoneStatus" "Choice" @("Not Started","In Progress","Complete","At Risk -- PAST DUE","Deferred")
Ensure-Field "PM_Milestones" "MilestonePct"    "Number"
Ensure-Field "PM_Milestones" "MilestoneTarget" "Text"
Ensure-Field "PM_Milestones" "MilestoneNotes"  "Note"

# ============================================================
# 11. PM_OpenItems  /  PM_DocumentDeliverables  /  PM_Budget
# ============================================================
Write-Host "`n[11/11] PM_OpenItems + PM_DocumentDeliverables + PM_Budget" -ForegroundColor Cyan

Ensure-List "PM_OpenItems" "IMP9177 PM Open Items"
Ensure-Field "PM_OpenItems" "OIRef"     "Text"
Ensure-Field "PM_OpenItems" "OITitle"   "Note"
Ensure-Field "PM_OpenItems" "OIOwner"   "Text"
Ensure-Field "PM_OpenItems" "OIPriority" "Choice" @("Critical","High","Medium","Low")
Ensure-Field "PM_OpenItems" "OIStatus"  "Choice" @("Open -- Critical","Open -- Urgent","Open","In Progress","Closed")

Ensure-List "PM_DocumentDeliverables" "IMP9177 PM Document Deliverables"
Ensure-Field "PM_DocumentDeliverables" "PMDocID"     "Text"
Ensure-Field "PM_DocumentDeliverables" "PMDocNumber" "Text"
Ensure-Field "PM_DocumentDeliverables" "PMDocWeek"   "Choice" @("W1","W2","W3","W4","Existing")
Ensure-Field "PM_DocumentDeliverables" "PMDocTitle"  "Text"
Ensure-Field "PM_DocumentDeliverables" "PMDocStatus" "Choice" @("Planned","ADB Drafting","Delivered -- Pending Feedback","PAST DUE -- Sign-Off Overdue","Signed Off","Complete")

Ensure-List "PM_Budget" "IMP9177 PM Budget"
Ensure-Field "PM_Budget" "BudgetSpent"     "Number"
Ensure-Field "PM_Budget" "BudgetCeiling"   "Number"
Ensure-Field "PM_Budget" "BudgetGapsAdded" "Number"
Ensure-Field "PM_Budget" "BudgetInvoice"   "Text"
Ensure-Field "PM_Budget" "BudgetInv1"      "Number"
Ensure-Field "PM_Budget" "BudgetInv2"      "Number"
Ensure-Field "PM_Budget" "BudgetInv2Ref"   "Text"

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host " LIST PROVISIONING COMPLETE" -ForegroundColor Green
Write-Host " Next: run seed_data.ps1 to load baseline records" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
