# =============================================================================
# IMP9177 — QMS PORTAL SHAREPOINT SETUP
# Script 05: Provision Application Lists
# Run AFTER 04_Validate_Published_Zone.ps1
#
# Creates 11 SharePoint lists that back the QMS Portal application:
#   QMS_DCOs, QMS_ChangeRequests, QMS_Approvers, QMS_DCOApprovals,
#   QMS_Records, QMS_Employees, QMS_Roles, QMS_TrainingMatrix,
#   QMS_TrainingCompletions, QMS_Config, QMS_RoutingHistory
#
# Run with -SeedData to populate lists with IMP9177 baseline data
# Run with -WhatIf to preview without making changes
# =============================================================================

param(
    [string]$SiteUrl  = "https://adbccro.sharepoint.com/sites/IMP9177",
    [switch]$SeedData = $false,
    [switch]$WhatIf   = $false
)

$ErrorActionPreference = "Continue"
$ClientId = "ba48ac81-6f23-43bd-9797-ec2866071102"

Write-Host "`n[IMP9177] Connecting to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -ClientId $ClientId -Interactive
Write-Host "[OK] Connected`n" -ForegroundColor Green

# =============================================================================
# HELPERS
# =============================================================================

function Ensure-List {
    param([string]$Name, [string]$Description)
    $existing = Get-PnPList -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  [EXISTS]  $Name" -ForegroundColor DarkGray
        return $existing
    }
    Write-Host "  [CREATING] $Name" -ForegroundColor Green
    if (-not $WhatIf) {
        $list = New-PnPList -Title $Name -Template GenericList -ErrorAction Stop
        Set-PnPList -Identity $Name -Description $Description | Out-Null
        Write-Host "             Created." -ForegroundColor Green
        return $list
    }
    return $null
}

function Ensure-Field {
    param([string]$List, [string]$Name, [string]$Type, [string[]]$Choices)
    $existing = Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue
    if ($existing) { return }
    if ($WhatIf) { Write-Host "    [WHATIF] Would add field $Name ($Type) to $List"; return }
    try {
        switch ($Type) {
            "Text"     { $schema = "<Field Type='Text' DisplayName='$Name' Name='$Name' StaticName='$Name'/>" }
            "Note"     { $schema = "<Field Type='Note' DisplayName='$Name' Name='$Name' StaticName='$Name' NumLines='6' RichText='FALSE'/>" }
            "DateTime" { $schema = "<Field Type='DateTime' DisplayName='$Name' Name='$Name' StaticName='$Name' Format='DateOnly'/>" }
            "Boolean"  { $schema = "<Field Type='Boolean' DisplayName='$Name' Name='$Name' StaticName='$Name'/>" }
            "Number"   { $schema = "<Field Type='Number' DisplayName='$Name' Name='$Name' StaticName='$Name'/>" }
            "Choice"   {
                $choiceXml = "<CHOICES>" + ($Choices | ForEach-Object { "<CHOICE>$_</CHOICE>" }) + "</CHOICES>"
                $schema = "<Field Type='Choice' DisplayName='$Name' Name='$Name' StaticName='$Name'>$choiceXml</Field>"
            }
        }
        Add-PnPFieldFromXml -List $List -FieldXml $schema -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "    [WARN] Could not add $Name to $List — $_"
    }
}

function Add-Item {
    param([string]$List, [hashtable]$Values)
    if ($WhatIf) { return }
    try {
        Add-PnPListItem -List $List -Values $Values -ErrorAction Stop | Out-Null
    } catch {
        Write-Warning "  [WARN] Could not add item to $List — $_"
    }
}

function Clear-List {
    param([string]$List)
    if ($WhatIf) { return }
    try {
        $items = Get-PnPListItem -List $List -ErrorAction SilentlyContinue
        foreach ($item in $items) {
            Remove-PnPListItem -List $List -Identity $item.Id -Force -ErrorAction SilentlyContinue | Out-Null
        }
    } catch {}
}

# =============================================================================
# 1. QMS_Config — System configuration (single row)
# =============================================================================
Write-Host "=== 1. QMS_Config ===" -ForegroundColor Yellow
Ensure-List "QMS_Config" "QMS Portal system configuration settings — single row"

Ensure-Field "QMS_Config" "CFG_ApprovalOverdueDays"  "Number"
Ensure-Field "QMS_Config" "CFG_ApprovalWarningDays"  "Number"
Ensure-Field "QMS_Config" "CFG_DraftStaleDays"        "Number"
Ensure-Field "QMS_Config" "CFG_TrainingDueDays"       "Number"
Ensure-Field "QMS_Config" "CFG_TrainingWarningDays"   "Number"
Ensure-Field "QMS_Config" "CFG_TrainingGateDays"      "Number"
Ensure-Field "QMS_Config" "CFG_MinTrainingRecs"       "Number"
Ensure-Field "QMS_Config" "CFG_RevScheme"             "Choice" @("Alpha","Numeric")
Ensure-Field "QMS_Config" "CFG_NotifyOnSubmit"        "Boolean"
Ensure-Field "QMS_Config" "CFG_NotifyOnOverdue"       "Boolean"
Ensure-Field "QMS_Config" "CFG_NotifyOnRejection"     "Boolean"
Ensure-Field "QMS_Config" "CFG_EscalationEmail"       "Text"
Ensure-Field "QMS_Config" "CFG_NotificationFrom"      "Text"
Ensure-Field "QMS_Config" "CFG_SPSiteUrl"             "Text"
Ensure-Field "QMS_Config" "CFG_DraftZonePath"         "Text"
Ensure-Field "QMS_Config" "CFG_PublishedZonePath"     "Text"
Ensure-Field "QMS_Config" "CFG_OfficialZonePath"      "Text"

# =============================================================================
# 2. QMS_Roles — Role definitions
# =============================================================================
Write-Host "`n=== 2. QMS_Roles ===" -ForegroundColor Yellow
Ensure-List "QMS_Roles" "QMS role definitions — each role maps to a set of required documents"

Ensure-Field "QMS_Roles" "Role_ID"          "Text"
Ensure-Field "QMS_Roles" "Role_Name"        "Text"
Ensure-Field "QMS_Roles" "Role_Description" "Note"
Ensure-Field "QMS_Roles" "Role_RequiredDocs" "Note"   # Comma-separated doc IDs

# =============================================================================
# 3. QMS_Employees — Employee table with role assignments
# =============================================================================
Write-Host "`n=== 3. QMS_Employees ===" -ForegroundColor Yellow
Ensure-List "QMS_Employees" "3H Pharmaceuticals employees with QMS role assignments"

Ensure-Field "QMS_Employees" "Emp_ID"        "Text"
Ensure-Field "QMS_Employees" "Emp_FullName"  "Text"
Ensure-Field "QMS_Employees" "Emp_Title"     "Text"
Ensure-Field "QMS_Employees" "Emp_Dept"      "Choice" @("Quality","Production","Warehouse","R&D","Management","Purchasing","Marketing")
Ensure-Field "QMS_Employees" "Emp_Email"     "Text"
Ensure-Field "QMS_Employees" "Emp_RoleIDs"   "Text"   # Comma-separated Role_IDs
Ensure-Field "QMS_Employees" "Emp_HireDate"  "DateTime"
Ensure-Field "QMS_Employees" "Emp_Status"    "Choice" @("Active","Inactive","On Leave")

# =============================================================================
# 4. QMS_Approvers — DCO approver configuration
# =============================================================================
Write-Host "`n=== 4. QMS_Approvers ===" -ForegroundColor Yellow
Ensure-List "QMS_Approvers" "Configured approvers for DCO routing — defines who signs what"

Ensure-Field "QMS_Approvers" "Approver_Name"     "Text"
Ensure-Field "QMS_Approvers" "Approver_Title"    "Text"
Ensure-Field "QMS_Approvers" "Approver_Email"    "Text"
Ensure-Field "QMS_Approvers" "Approver_Initials" "Text"
Ensure-Field "QMS_Approvers" "Approver_Type"     "Choice" @("Required","Conditional","Optional")
Ensure-Field "QMS_Approvers" "Approver_Scope"    "Choice" @("All DCOs","Quality Manuals only","SOPs only","Forms only","Records only","Custom")
Ensure-Field "QMS_Approvers" "Approver_ScopeDetail" "Text"   # e.g. "QM-001,FM-027"
Ensure-Field "QMS_Approvers" "Approver_Mode"     "Choice" @("Parallel","Sequential")
Ensure-Field "QMS_Approvers" "Approver_Active"   "Boolean"

# =============================================================================
# 5. QMS_ChangeRequests — CR lifecycle management
# =============================================================================
Write-Host "`n=== 5. QMS_ChangeRequests ===" -ForegroundColor Yellow
Ensure-List "QMS_ChangeRequests" "Change Request lifecycle — Create, Review, Approve, Link to DCO, Close"

Ensure-Field "QMS_ChangeRequests" "CR_ID"           "Text"
Ensure-Field "QMS_ChangeRequests" "CR_Title"         "Text"
Ensure-Field "QMS_ChangeRequests" "CR_Category"      "Choice" @("Document Revision","New Document","Process Change","Supplier Change","Regulatory Update","CAPA-Driven","Customer Complaint","Internal Audit Finding")
Ensure-Field "QMS_ChangeRequests" "CR_Priority"      "Choice" @("Routine","Urgent","Critical")
Ensure-Field "QMS_ChangeRequests" "CR_Status"        "Choice" @("Draft","In Review","Approved","Linked to DCO","Closed","Rejected")
Ensure-Field "QMS_ChangeRequests" "CR_Originator"    "Text"
Ensure-Field "QMS_ChangeRequests" "CR_Reviewer"      "Text"
Ensure-Field "QMS_ChangeRequests" "CR_OpenedDate"    "DateTime"
Ensure-Field "QMS_ChangeRequests" "CR_ClosedDate"    "DateTime"
Ensure-Field "QMS_ChangeRequests" "CR_Description"   "Note"
Ensure-Field "QMS_ChangeRequests" "CR_Justification" "Note"
Ensure-Field "QMS_ChangeRequests" "CR_AffectedDocs"  "Text"
Ensure-Field "QMS_ChangeRequests" "CR_LinkedDCOs"    "Text"   # Comma-separated DCO IDs

# =============================================================================
# 6. QMS_DCOs — Document Change Order routing
# =============================================================================
Write-Host "`n=== 6. QMS_DCOs ===" -ForegroundColor Yellow
Ensure-List "QMS_DCOs" "Document Change Orders — full 6-phase routing lifecycle"

Ensure-Field "QMS_DCOs" "DCO_ID"             "Text"
Ensure-Field "QMS_DCOs" "DCO_Title"          "Text"
Ensure-Field "QMS_DCOs" "DCO_Phase"          "Choice" @("Draft","Submitted","In Review","Implemented","Awaiting Training","Effective")
Ensure-Field "QMS_DCOs" "DCO_LinkedCR"       "Text"
Ensure-Field "QMS_DCOs" "DCO_Originator"     "Text"
Ensure-Field "QMS_DCOs" "DCO_Documents"      "Note"   # Newline-separated doc IDs
Ensure-Field "QMS_DCOs" "DCO_SubmittedDate"  "DateTime"
Ensure-Field "QMS_DCOs" "DCO_ImplementedDate" "DateTime"
Ensure-Field "QMS_DCOs" "DCO_EffectiveDate"  "DateTime"
Ensure-Field "QMS_DCOs" "DCO_TargetEffDate"  "DateTime"
Ensure-Field "QMS_DCOs" "DCO_HasSOPs"        "Boolean"   # Drives training gate
Ensure-Field "QMS_DCOs" "DCO_TrainingGateCleared" "Boolean"
Ensure-Field "QMS_DCOs" "DCO_IsLate"         "Boolean"
Ensure-Field "QMS_DCOs" "DCO_Notes"          "Note"

# =============================================================================
# 7. QMS_DCOApprovals — Per-DCO per-approver signing record
# =============================================================================
Write-Host "`n=== 7. QMS_DCOApprovals ===" -ForegroundColor Yellow
Ensure-List "QMS_DCOApprovals" "Individual approver signing records for each DCO"

Ensure-Field "QMS_DCOApprovals" "Sig_DCOID"        "Text"
Ensure-Field "QMS_DCOApprovals" "Sig_ApproverName" "Text"
Ensure-Field "QMS_DCOApprovals" "Sig_ApproverEmail" "Text"
Ensure-Field "QMS_DCOApprovals" "Sig_Role"         "Text"
Ensure-Field "QMS_DCOApprovals" "Sig_Type"         "Choice" @("Required","Conditional","Optional")
Ensure-Field "QMS_DCOApprovals" "Sig_Status"       "Choice" @("Waiting","Pending","Signed","Blocked","Rejected")
Ensure-Field "QMS_DCOApprovals" "Sig_SignedDate"   "DateTime"
Ensure-Field "QMS_DCOApprovals" "Sig_SignatureID"  "Text"
Ensure-Field "QMS_DCOApprovals" "Sig_BlockReason"  "Text"
Ensure-Field "QMS_DCOApprovals" "Sig_Method"       "Text"   # e.g. SharePoint E-Signature + M365 MFA

# =============================================================================
# 8. QMS_RoutingHistory — Audit log for all DCO and CR transitions
# =============================================================================
Write-Host "`n=== 8. QMS_RoutingHistory ===" -ForegroundColor Yellow
Ensure-List "QMS_RoutingHistory" "Immutable audit trail for all DCO and CR stage transitions and rejections"

Ensure-Field "QMS_RoutingHistory" "RH_EntityType"   "Choice" @("DCO","CR","Record")
Ensure-Field "QMS_RoutingHistory" "RH_EntityID"     "Text"
Ensure-Field "QMS_RoutingHistory" "RH_EventType"    "Choice" @("Stage Change","Signature","Rejection","Comment","Creation","System")
Ensure-Field "QMS_RoutingHistory" "RH_FromPhase"    "Text"
Ensure-Field "QMS_RoutingHistory" "RH_ToPhase"      "Text"
Ensure-Field "QMS_RoutingHistory" "RH_User"         "Text"
Ensure-Field "QMS_RoutingHistory" "RH_Timestamp"    "DateTime"
Ensure-Field "QMS_RoutingHistory" "RH_Note"         "Note"
Ensure-Field "QMS_RoutingHistory" "RH_RejCategory"  "Text"   # Rejection category if applicable
Ensure-Field "QMS_RoutingHistory" "RH_RejReason"    "Note"   # Full rejection reason

# =============================================================================
# 9. QMS_Records — Record instances (batch records, CoA logs, training, etc.)
# =============================================================================
Write-Host "`n=== 9. QMS_Records ===" -ForegroundColor Yellow
Ensure-List "QMS_Records" "QMS record instances — routed through 5-stage pipeline to Official zone"

Ensure-Field "QMS_Records" "Rec_ID"           "Text"
Ensure-Field "QMS_Records" "Rec_Type"         "Choice" @("Batch Record","CoA Log","CAPA Record","Training Record","Pest Control Log","Deviation Report","Environmental Record","Supplier Record")
Ensure-Field "QMS_Records" "Rec_Title"        "Text"
Ensure-Field "QMS_Records" "Rec_Status"       "Choice" @("Draft","In Review","Approved","Pending Signature","Signed & Filed")
Ensure-Field "QMS_Records" "Rec_Originator"   "Text"
Ensure-Field "QMS_Records" "Rec_OrigDate"     "DateTime"
Ensure-Field "QMS_Records" "Rec_Reviewer"     "Text"
Ensure-Field "QMS_Records" "Rec_Signer"       "Text"
Ensure-Field "QMS_Records" "Rec_RelatedSOP"   "Text"
Ensure-Field "QMS_Records" "Rec_Content"      "Note"
Ensure-Field "QMS_Records" "Rec_SignatureID"  "Text"
Ensure-Field "QMS_Records" "Rec_SignedDate"   "DateTime"
Ensure-Field "QMS_Records" "Rec_FiledPath"    "Text"   # Official zone path after filing
Ensure-Field "QMS_Records" "Rec_ReviewNote"   "Note"
Ensure-Field "QMS_Records" "Rec_RejReason"    "Note"

# =============================================================================
# 10. QMS_TrainingMatrix — Role x Document requirements
# =============================================================================
Write-Host "`n=== 10. QMS_TrainingMatrix ===" -ForegroundColor Yellow
Ensure-List "QMS_TrainingMatrix" "Training matrix — defines which roles must be trained to which documents"

Ensure-Field "QMS_TrainingMatrix" "TM_RoleID"      "Text"
Ensure-Field "QMS_TrainingMatrix" "TM_RoleName"    "Text"
Ensure-Field "QMS_TrainingMatrix" "TM_DocID"       "Text"
Ensure-Field "QMS_TrainingMatrix" "TM_DocTitle"    "Text"
Ensure-Field "QMS_TrainingMatrix" "TM_DocType"     "Choice" @("QM","SOP","FM","REC")
Ensure-Field "QMS_TrainingMatrix" "TM_Required"    "Boolean"
Ensure-Field "QMS_TrainingMatrix" "TM_CurrentRev"  "Text"
Ensure-Field "QMS_TrainingMatrix" "TM_EffectiveDate" "DateTime"

# =============================================================================
# 11. QMS_TrainingCompletions — Signed training records
# =============================================================================
Write-Host "`n=== 11. QMS_TrainingCompletions ===" -ForegroundColor Yellow
Ensure-List "QMS_TrainingCompletions" "Completed signed training records per employee per document per revision"

Ensure-Field "QMS_TrainingCompletions" "TC_EmpID"      "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_EmpName"    "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_DocID"      "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_DocTitle"   "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_Revision"   "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_Method"     "Choice" @("Read & Understand","Instructor-led","On-the-Job","Computer-based","Observation")
Ensure-Field "QMS_TrainingCompletions" "TC_TrainDate"  "DateTime"
Ensure-Field "QMS_TrainingCompletions" "TC_RecordID"   "Text"   # Links to QMS_Records
Ensure-Field "QMS_TrainingCompletions" "TC_SignatureID" "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_Trainer"    "Text"

Write-Host "`n[OK] All 11 lists created/verified with columns." -ForegroundColor Cyan

# =============================================================================
# SEED DATA — IMP9177 baseline data
# Run with -SeedData switch
# =============================================================================
if ($SeedData) {
    Write-Host "`n=== SEEDING BASELINE DATA ===" -ForegroundColor Magenta

    # ── Config ──────────────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_Config..." -ForegroundColor White
    Clear-List "QMS_Config"
    Add-Item "QMS_Config" @{
        Title                  = "IMP9177 Configuration"
        CFG_ApprovalOverdueDays = 14
        CFG_ApprovalWarningDays = 7
        CFG_DraftStaleDays      = 30
        CFG_TrainingDueDays     = 30
        CFG_TrainingWarningDays = 7
        CFG_TrainingGateDays    = 30
        CFG_MinTrainingRecs     = 1
        CFG_RevScheme           = "Alpha"
        CFG_NotifyOnSubmit      = $true
        CFG_NotifyOnOverdue     = $true
        CFG_NotifyOnRejection   = $true
        CFG_EscalationEmail     = "tina@3hpharma.com"
        CFG_NotificationFrom    = "qms-noreply@3hpharma.com"
        CFG_SPSiteUrl           = "https://adbccro.sharepoint.com/sites/IMP9177"
        CFG_DraftZonePath       = "Shared Documents/QMS/Documents/Drafts"
        CFG_PublishedZonePath   = "Shared Documents/Published/QMS"
        CFG_OfficialZonePath    = "Shared Documents/Official/QMS"
    }

    # ── Roles ───────────────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_Roles..." -ForegroundColor White
    Clear-List "QMS_Roles"
    $roles = @(
        @{Title="Management";           Role_ID="ROLE-MGT";  Role_Description="Senior leadership — quality system awareness and sign-off authority";                    Role_RequiredDocs="QM-001,SOP-001,SOP-002,SOP-004"},
        @{Title="QA / Quality Approver";Role_ID="ROLE-QA";   Role_Description="QA oversight — full document control, CAPA, supplier, and training authority";          Role_RequiredDocs="QM-001,SOP-001,SOP-002,SOP-003,SOP-004,FM-001,FM-027"},
        @{Title="R&D / QA Director";    Role_ID="ROLE-RD";   Role_Description="R&D and quality oversight — regulatory, CAPA, supplier qualification";                  Role_RequiredDocs="QM-001,SOP-001,SOP-002,SOP-003,SOP-004"},
        @{Title="Production Operator";  Role_ID="ROLE-PROD"; Role_Description="Production floor operations — batch records, GMP, production SOPs";                    Role_RequiredDocs="QM-001,SOP-001,SOP-004,FM-001"},
        @{Title="Warehouse / Receiving";Role_ID="ROLE-WH";   Role_Description="Receiving, storage, and inventory — supplier, receiving, and pest control";            Role_RequiredDocs="QM-001,SOP-001,SOP-003,SOP-004,SOP-PC-001"},
        @{Title="QC Technician";        Role_ID="ROLE-QC";   Role_Description="Quality control testing and CoA verification";                                          Role_RequiredDocs="QM-001,SOP-001,SOP-002,SOP-003,SOP-004,FM-001"},
        @{Title="Project Manager";      Role_ID="ROLE-PM";   Role_Description="QMS project management — full system awareness";                                        Role_RequiredDocs="QM-001,SOP-001,SOP-002,SOP-003,SOP-004"}
    )
    foreach ($r in $roles) { Add-Item "QMS_Roles" $r }

    # ── Employees ───────────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_Employees..." -ForegroundColor White
    Clear-List "QMS_Employees"
    $employees = @(
        @{Title="Liu Hong (Ace)"; Emp_ID="EMP-001"; Emp_Title="CEO / Owner";               Emp_Dept="Management"; Emp_Email="ace@3hpharma.com";            Emp_RoleIDs="ROLE-MGT";           Emp_Status="Active"},
        @{Title="Tina Qin";       Emp_ID="EMP-002"; Emp_Title="Director of Purchasing";    Emp_Dept="Quality";    Emp_Email="tina@3hpharma.com";           Emp_RoleIDs="ROLE-QA,ROLE-MGT";   Emp_Status="Active"},
        @{Title="QD Yang";        Emp_ID="EMP-003"; Emp_Title="Director of R&D/QA";        Emp_Dept="R&D";        Emp_Email="qd@3hpharma.com";             Emp_RoleIDs="ROLE-RD,ROLE-QA";    Emp_Status="Active"},
        @{Title="Baolong Sheng";  Emp_ID="EMP-004"; Emp_Title="Director of Production";    Emp_Dept="Production"; Emp_Email="bs@3hpharma.com";             Emp_RoleIDs="ROLE-PROD,ROLE-MGT"; Emp_Status="Active"},
        @{Title="Kevin Xu";       Emp_ID="EMP-005"; Emp_Title="QC/IT Technician";          Emp_Dept="Quality";    Emp_Email="kx@3hpharma.com";             Emp_RoleIDs="ROLE-QC";            Emp_Status="Active"},
        @{Title="Cindy Dong";     Emp_ID="EMP-006"; Emp_Title="Marketing Assistant";       Emp_Dept="Marketing";  Emp_Email="cindydong3h@gmail.com";       Emp_RoleIDs="ROLE-PM";            Emp_Status="Active"},
        @{Title="KJ";             Emp_ID="EMP-007"; Emp_Title="Warehouse / Receiving";     Emp_Dept="Warehouse";  Emp_Email="kj@3hpharma.com";             Emp_RoleIDs="ROLE-WH";            Emp_Status="Active"},
        @{Title="Andre Butler";   Emp_ID="EMP-008"; Emp_Title="Project Manager (ADB)";    Emp_Dept="Quality";    Emp_Email="andre.butler@adbccro.com";    Emp_RoleIDs="ROLE-PM";            Emp_Status="Active"}
    )
    foreach ($e in $employees) { Add-Item "QMS_Employees" $e }

    # ── Approvers ───────────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_Approvers..." -ForegroundColor White
    Clear-List "QMS_Approvers"
    $approvers = @(
        @{Title="Tina Qin";      Approver_Name="Tina Qin";      Approver_Title="Director of Purchasing"; Approver_Email="tina@3hpharma.com";          Approver_Initials="TQ"; Approver_Type="Required";    Approver_Scope="All DCOs";            Approver_Mode="Parallel"; Approver_Active=$true},
        @{Title="Liu Hong";      Approver_Name="Liu Hong (Ace)"; Approver_Title="CEO / Owner";            Approver_Email="ace@3hpharma.com";            Approver_Initials="LH"; Approver_Type="Conditional"; Approver_Scope="Custom";              Approver_ScopeDetail="QM-001,FM-027"; Approver_Mode="Parallel"; Approver_Active=$true},
        @{Title="QD Yang";       Approver_Name="QD Yang";        Approver_Title="Director R&D/QA";        Approver_Email="qd@3hpharma.com";             Approver_Initials="QY"; Approver_Type="Optional";    Approver_Scope="Custom";              Approver_ScopeDetail="R&D-impacted"; Approver_Mode="Parallel"; Approver_Active=$true},
        @{Title="Baolong Sheng"; Approver_Name="Baolong Sheng";  Approver_Title="Director of Production"; Approver_Email="bs@3hpharma.com";             Approver_Initials="BS"; Approver_Type="Optional";    Approver_Scope="SOPs only";           Approver_Mode="Parallel"; Approver_Active=$true}
    )
    foreach ($a in $approvers) { Add-Item "QMS_Approvers" $a }

    # ── Change Requests ──────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_ChangeRequests..." -ForegroundColor White
    Clear-List "QMS_ChangeRequests"
    $crs = @(
        @{Title="CR-0001"; CR_ID="CR-0001"; CR_Title="Melatonin Label Revision — Week 1 Document Package";
          CR_Category="Document Revision"; CR_Priority="Routine"; CR_Status="Linked to DCO";
          CR_Originator="Andre Butler"; CR_Reviewer="Tina Qin";
          CR_OpenedDate="2026-03-20"; CR_LinkedDCOs="DCO-0001";
          CR_AffectedDocs="QM-001,SOP-001,SOP-002,SOP-003,SOP-004,FM-001,FM-002,FM-027";
          CR_Description="Initiate formal issue of the first 15 QMS documents identified during Phase 1 gap assessment.";
          CR_Justification="21 CFR 111.68 and 111.103 require written procedures. Gap assessment identified informal SOPs only."},
        @{Title="CR-0002"; CR_ID="CR-0002"; CR_Title="Week 2 QMS Document Package";
          CR_Category="Document Revision"; CR_Priority="Routine"; CR_Status="Linked to DCO";
          CR_Originator="Andre Butler"; CR_Reviewer="Tina Qin";
          CR_OpenedDate="2026-04-05"; CR_LinkedDCOs="DCO-0002";
          CR_AffectedDocs="FM-008,SOP-SUP-001,SOP-PC-001,FM-004,FM-005,FM-006,FM-007";
          CR_Description="Issue the second batch of 12 QMS documents as identified in the project document register.";
          CR_Justification="Continuation of Phase 2A QMS implementation for full 21 CFR 111 and FSMA compliance."},
        @{Title="CR-0003"; CR_ID="CR-0003"; CR_Title="Orkin Stored Product Pest Control Addendum";
          CR_Category="Regulatory Update"; CR_Priority="Urgent"; CR_Status="Approved";
          CR_Originator="Andre Butler"; CR_Reviewer="Tina Qin";
          CR_OpenedDate="2026-04-18"; CR_LinkedDCOs="DCO-0003";
          CR_AffectedDocs="SOP-PC-001";
          CR_Description="Add formal Orkin stored product pest control addendum to the pest control program.";
          CR_Justification="Gap GAP-PC-03 from Phase 1 assessment. 21 CFR 117.35(c) requires pest control coverage."}
    )
    foreach ($cr in $crs) { Add-Item "QMS_ChangeRequests" $cr }

    # ── DCOs ────────────────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_DCOs..." -ForegroundColor White
    Clear-List "QMS_DCOs"
    $dcos = @(
        @{Title="DCO-0001"; DCO_ID="DCO-0001"; DCO_Title="Week 1 QMS Document Package";
          DCO_Phase="In Review"; DCO_LinkedCR="CR-0001"; DCO_Originator="Andre Butler";
          DCO_SubmittedDate="2026-03-28"; DCO_HasSOPs=$true; DCO_TrainingGateCleared=$false; DCO_IsLate=$true;
          DCO_Documents="QM-001`nSOP-001`nSOP-002`nSOP-003`nSOP-004`nSOP-005`nSOP-006`nSOP-007`nFM-001`nFM-002`nFM-003`nFM-008`nFM-025`nFM-027`nFM-030";
          DCO_Notes="DCO-0001 sign-off PAST DUE — AC-034 blocking QD Yang authorization"},
        @{Title="DCO-0002"; DCO_ID="DCO-0002"; DCO_Title="Week 2 Document Package";
          DCO_Phase="Submitted"; DCO_LinkedCR="CR-0002"; DCO_Originator="Andre Butler";
          DCO_SubmittedDate="2026-04-07"; DCO_HasSOPs=$true; DCO_TrainingGateCleared=$false; DCO_IsLate=$true;
          DCO_Documents="SOP-FS-001`nSOP-FS-002`nSOP-FS-003`nSOP-FS-004`nSOP-PC-001`nSOP-SUP-001`nSOP-SUP-002`nFM-004`nFM-005`nFM-006`nFM-007`nFM-ALG";
          DCO_Notes="Delivered to 3H April 10 — pending approval routing"},
        @{Title="DCO-0003"; DCO_ID="DCO-0003"; DCO_Title="Orkin Pest Control Addendum";
          DCO_Phase="Draft"; DCO_LinkedCR="CR-0003"; DCO_Originator="Andre Butler";
          DCO_HasSOPs=$true; DCO_TrainingGateCleared=$false; DCO_IsLate=$false;
          DCO_Documents="SOP-PC-001";
          DCO_Notes="AC-023 open — Orkin addendum required before submission"}
    )
    foreach ($d in $dcos) { Add-Item "QMS_DCOs" $d }

    # ── DCO Approvals ────────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_DCOApprovals..." -ForegroundColor White
    Clear-List "QMS_DCOApprovals"
    $approvals = @(
        @{Title="DCO-0001-TQ"; Sig_DCOID="DCO-0001"; Sig_ApproverName="Tina Qin";      Sig_ApproverEmail="tina@3hpharma.com"; Sig_Role="QA Approver";       Sig_Type="Required";    Sig_Status="Signed";  Sig_SignedDate="2026-04-19"; Sig_SignatureID="SIG-TQ-8B2C"; Sig_Method="SharePoint E-Signature + M365 MFA"},
        @{Title="DCO-0001-LH"; Sig_DCOID="DCO-0001"; Sig_ApproverName="Liu Hong";      Sig_ApproverEmail="ace@3hpharma.com";  Sig_Role="CEO / Owner";        Sig_Type="Conditional"; Sig_Status="Signed";  Sig_SignedDate="2026-04-19"; Sig_SignatureID="SIG-LH-4F9A"; Sig_Method="SharePoint E-Signature + M365 MFA"},
        @{Title="DCO-0001-QY"; Sig_DCOID="DCO-0001"; Sig_ApproverName="QD Yang";       Sig_ApproverEmail="qd@3hpharma.com";   Sig_Role="R&D/QA Director";    Sig_Type="Optional";    Sig_Status="Blocked"; Sig_BlockReason="AC-034: Work authorization pending"},
        @{Title="DCO-0002-TQ"; Sig_DCOID="DCO-0002"; Sig_ApproverName="Tina Qin";      Sig_ApproverEmail="tina@3hpharma.com"; Sig_Role="QA Approver";       Sig_Type="Required";    Sig_Status="Pending"},
        @{Title="DCO-0002-LH"; Sig_DCOID="DCO-0002"; Sig_ApproverName="Liu Hong";      Sig_ApproverEmail="ace@3hpharma.com";  Sig_Role="CEO / Owner";        Sig_Type="Conditional"; Sig_Status="Waiting"},
        @{Title="DCO-0002-QY"; Sig_DCOID="DCO-0002"; Sig_ApproverName="QD Yang";       Sig_ApproverEmail="qd@3hpharma.com";   Sig_Role="R&D/QA Director";    Sig_Type="Optional";    Sig_Status="Waiting"},
        @{Title="DCO-0003-TQ"; Sig_DCOID="DCO-0003"; Sig_ApproverName="Tina Qin";      Sig_ApproverEmail="tina@3hpharma.com"; Sig_Role="QA Approver";       Sig_Type="Required";    Sig_Status="Waiting"},
        @{Title="DCO-0003-BS"; Sig_DCOID="DCO-0003"; Sig_ApproverName="Baolong Sheng"; Sig_ApproverEmail="bs@3hpharma.com";   Sig_Role="Production Director"; Sig_Type="Optional";    Sig_Status="Waiting"}
    )
    foreach ($a in $approvals) { Add-Item "QMS_DCOApprovals" $a }

    # ── Routing History ──────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_RoutingHistory..." -ForegroundColor White
    Clear-List "QMS_RoutingHistory"
    $history = @(
        @{Title="DCO-0001-01"; RH_EntityType="DCO"; RH_EntityID="DCO-0001"; RH_EventType="Creation";     RH_ToPhase="Draft";       RH_User="Andre Butler"; RH_Timestamp="2026-03-20"; RH_Note="DCO-0001 created with 15 documents"},
        @{Title="DCO-0001-02"; RH_EntityType="DCO"; RH_EntityID="DCO-0001"; RH_EventType="Stage Change"; RH_FromPhase="Draft";     RH_ToPhase="Submitted";   RH_User="Andre Butler"; RH_Timestamp="2026-03-28"; RH_Note="DCO submitted for approval routing"},
        @{Title="DCO-0001-03"; RH_EntityType="DCO"; RH_EntityID="DCO-0001"; RH_EventType="Stage Change"; RH_FromPhase="Submitted"; RH_ToPhase="In Review";   RH_User="System";        RH_Timestamp="2026-03-28"; RH_Note="All approvers notified via SharePoint e-signature"},
        @{Title="DCO-0001-04"; RH_EntityType="DCO"; RH_EntityID="DCO-0001"; RH_EventType="Signature";    RH_User="Tina Qin";       RH_Timestamp="2026-04-19"; RH_Note="Signed (QA Approver) via SharePoint E-Signature SIG-TQ-8B2C"},
        @{Title="DCO-0001-05"; RH_EntityType="DCO"; RH_EntityID="DCO-0001"; RH_EventType="Signature";    RH_User="Liu Hong";       RH_Timestamp="2026-04-19"; RH_Note="Signed (CEO approval) via SharePoint E-Signature SIG-LH-4F9A"},
        @{Title="CR-0001-01";  RH_EntityType="CR";  RH_EntityID="CR-0001";  RH_EventType="Creation";     RH_ToPhase="Draft";       RH_User="Andre Butler"; RH_Timestamp="2026-03-20"; RH_Note="CR-0001 created from Phase 1 gap assessment findings"},
        @{Title="CR-0001-02";  RH_EntityType="CR";  RH_EntityID="CR-0001";  RH_EventType="Stage Change"; RH_FromPhase="Draft";     RH_ToPhase="In Review";   RH_User="Andre Butler"; RH_Timestamp="2026-03-22"; RH_Note="Submitted for QA review"},
        @{Title="CR-0001-03";  RH_EntityType="CR";  RH_EntityID="CR-0001";  RH_EventType="Stage Change"; RH_FromPhase="In Review"; RH_ToPhase="Approved";    RH_User="Tina Qin";     RH_Timestamp="2026-03-27"; RH_Note="Review complete — change approved for execution"},
        @{Title="CR-0002-01";  RH_EntityType="CR";  RH_EntityID="CR-0002";  RH_EventType="Rejection";    RH_FromPhase="In Review"; RH_ToPhase="Draft";       RH_User="Tina Qin";     RH_Timestamp="2026-04-08"; RH_RejCategory="Document content error"; RH_RejReason="FM-008 CAPA form missing section 4 root cause fields. Revise before resubmitting."},
        @{Title="CR-0002-02";  RH_EntityType="CR";  RH_EntityID="CR-0002";  RH_EventType="Stage Change"; RH_FromPhase="Draft";     RH_ToPhase="Approved";    RH_User="Tina Qin";     RH_Timestamp="2026-04-12"; RH_Note="Week 2 package approved after FM-008 correction"}
    )
    foreach ($h in $history) { Add-Item "QMS_RoutingHistory" $h }

    # ── Training Matrix ──────────────────────────────────────────────────────
    Write-Host "  Seeding QMS_TrainingMatrix..." -ForegroundColor White
    Clear-List "QMS_TrainingMatrix"
    $matrix = @(
        # Management role
        @{Title="ROLE-MGT-QM-001";  TM_RoleID="ROLE-MGT";  TM_RoleName="Management";           TM_DocID="QM-001";  TM_DocTitle="Quality Manual";           TM_DocType="QM";  TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-15"},
        @{Title="ROLE-MGT-SOP-001"; TM_RoleID="ROLE-MGT";  TM_RoleName="Management";           TM_DocID="SOP-001"; TM_DocTitle="Document Control";         TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-17"},
        @{Title="ROLE-MGT-SOP-002"; TM_RoleID="ROLE-MGT";  TM_RoleName="Management";           TM_DocID="SOP-002"; TM_DocTitle="CAPA Procedure";           TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-18"},
        @{Title="ROLE-MGT-SOP-004"; TM_RoleID="ROLE-MGT";  TM_RoleName="Management";           TM_DocID="SOP-004"; TM_DocTitle="Training Control";         TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-16"},
        # QA role
        @{Title="ROLE-QA-QM-001";   TM_RoleID="ROLE-QA";   TM_RoleName="QA / Quality Approver"; TM_DocID="QM-001";  TM_DocTitle="Quality Manual";           TM_DocType="QM";  TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-15"},
        @{Title="ROLE-QA-SOP-001";  TM_RoleID="ROLE-QA";   TM_RoleName="QA / Quality Approver"; TM_DocID="SOP-001"; TM_DocTitle="Document Control";         TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-17"},
        @{Title="ROLE-QA-SOP-002";  TM_RoleID="ROLE-QA";   TM_RoleName="QA / Quality Approver"; TM_DocID="SOP-002"; TM_DocTitle="CAPA Procedure";           TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-18"},
        @{Title="ROLE-QA-SOP-003";  TM_RoleID="ROLE-QA";   TM_RoleName="QA / Quality Approver"; TM_DocID="SOP-003"; TM_DocTitle="Supplier Qualification";   TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-16"},
        @{Title="ROLE-QA-SOP-004";  TM_RoleID="ROLE-QA";   TM_RoleName="QA / Quality Approver"; TM_DocID="SOP-004"; TM_DocTitle="Training Control";         TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-16"},
        @{Title="ROLE-QA-FM-001";   TM_RoleID="ROLE-QA";   TM_RoleName="QA / Quality Approver"; TM_DocID="FM-001";  TM_DocTitle="Batch Record";             TM_DocType="FM";  TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-14"},
        # Production role
        @{Title="ROLE-PROD-QM-001"; TM_RoleID="ROLE-PROD"; TM_RoleName="Production Operator";   TM_DocID="QM-001";  TM_DocTitle="Quality Manual";           TM_DocType="QM";  TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-15"},
        @{Title="ROLE-PROD-SOP-001";TM_RoleID="ROLE-PROD"; TM_RoleName="Production Operator";   TM_DocID="SOP-001"; TM_DocTitle="Document Control";         TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-17"},
        @{Title="ROLE-PROD-SOP-004";TM_RoleID="ROLE-PROD"; TM_RoleName="Production Operator";   TM_DocID="SOP-004"; TM_DocTitle="Training Control";         TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-16"},
        @{Title="ROLE-PROD-FM-001"; TM_RoleID="ROLE-PROD"; TM_RoleName="Production Operator";   TM_DocID="FM-001";  TM_DocTitle="Batch Record";             TM_DocType="FM";  TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-14"},
        # Warehouse role
        @{Title="ROLE-WH-QM-001";   TM_RoleID="ROLE-WH";   TM_RoleName="Warehouse / Receiving"; TM_DocID="QM-001";  TM_DocTitle="Quality Manual";           TM_DocType="QM";  TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-15"},
        @{Title="ROLE-WH-SOP-001";  TM_RoleID="ROLE-WH";   TM_RoleName="Warehouse / Receiving"; TM_DocID="SOP-001"; TM_DocTitle="Document Control";         TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-17"},
        @{Title="ROLE-WH-SOP-003";  TM_RoleID="ROLE-WH";   TM_RoleName="Warehouse / Receiving"; TM_DocID="SOP-003"; TM_DocTitle="Supplier Qualification";   TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-16"},
        @{Title="ROLE-WH-SOP-004";  TM_RoleID="ROLE-WH";   TM_RoleName="Warehouse / Receiving"; TM_DocID="SOP-004"; TM_DocTitle="Training Control";         TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-16"},
        @{Title="ROLE-WH-SOP-PC";   TM_RoleID="ROLE-WH";   TM_RoleName="Warehouse / Receiving"; TM_DocID="SOP-PC-001"; TM_DocTitle="Pest Sighting Response"; TM_DocType="SOP"; TM_Required=$true; TM_CurrentRev="Rev A"; TM_EffectiveDate="2026-04-12"}
    )
    foreach ($m in $matrix) { Add-Item "QMS_TrainingMatrix" $m }

    # ── Training Completions (known signed records) ──────────────────────────
    Write-Host "  Seeding QMS_TrainingCompletions..." -ForegroundColor White
    Clear-List "QMS_TrainingCompletions"
    $completions = @(
        @{Title="TC-EMP001-QM001";  TC_EmpID="EMP-001"; TC_EmpName="Liu Hong (Ace)"; TC_DocID="QM-001";  TC_DocTitle="Quality Manual";         TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-16"; TC_RecordID="TR-001"},
        @{Title="TC-EMP001-SOP001"; TC_EmpID="EMP-001"; TC_EmpName="Liu Hong (Ace)"; TC_DocID="SOP-001"; TC_DocTitle="Document Control";       TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-17"; TC_RecordID="TR-002"},
        @{Title="TC-EMP002-QM001";  TC_EmpID="EMP-002"; TC_EmpName="Tina Qin";       TC_DocID="QM-001";  TC_DocTitle="Quality Manual";         TC_Revision="Rev A"; TC_Method="Instructor-led";   TC_TrainDate="2026-04-16"; TC_RecordID="TR-003"},
        @{Title="TC-EMP002-SOP001"; TC_EmpID="EMP-002"; TC_EmpName="Tina Qin";       TC_DocID="SOP-001"; TC_DocTitle="Document Control";       TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-17"; TC_RecordID="TR-004"},
        @{Title="TC-EMP002-SOP002"; TC_EmpID="EMP-002"; TC_EmpName="Tina Qin";       TC_DocID="SOP-002"; TC_DocTitle="CAPA Procedure";         TC_Revision="Rev A"; TC_Method="Instructor-led";   TC_TrainDate="2026-04-19"; TC_RecordID="TR-005"},
        @{Title="TC-EMP002-SOP003"; TC_EmpID="EMP-002"; TC_EmpName="Tina Qin";       TC_DocID="SOP-003"; TC_DocTitle="Supplier Qualification"; TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-19"; TC_RecordID="TR-006"},
        @{Title="TC-EMP003-QM001";  TC_EmpID="EMP-003"; TC_EmpName="QD Yang";         TC_DocID="QM-001";  TC_DocTitle="Quality Manual";         TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-16"; TC_RecordID="TR-007"},
        @{Title="TC-EMP004-QM001";  TC_EmpID="EMP-004"; TC_EmpName="Baolong Sheng";   TC_DocID="QM-001";  TC_DocTitle="Quality Manual";         TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-16"; TC_RecordID="TR-008"},
        @{Title="TC-EMP004-FM001";  TC_EmpID="EMP-004"; TC_EmpName="Baolong Sheng";   TC_DocID="FM-001";  TC_DocTitle="Batch Record";           TC_Revision="Rev A"; TC_Method="Instructor-led";   TC_TrainDate="2026-04-15"; TC_RecordID="TR-009"},
        @{Title="TC-EMP005-QM001";  TC_EmpID="EMP-005"; TC_EmpName="Kevin Xu";         TC_DocID="QM-001";  TC_DocTitle="Quality Manual";         TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-16"; TC_RecordID="TR-010"},
        @{Title="TC-EMP005-FM001";  TC_EmpID="EMP-005"; TC_EmpName="Kevin Xu";         TC_DocID="FM-001";  TC_DocTitle="Batch Record";           TC_Revision="Rev A"; TC_Method="Instructor-led";   TC_TrainDate="2026-04-15"; TC_RecordID="TR-011"},
        @{Title="TC-EMP006-QM001";  TC_EmpID="EMP-006"; TC_EmpName="Cindy Dong";       TC_DocID="QM-001";  TC_DocTitle="Quality Manual";         TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-17"; TC_RecordID="TR-012"},
        @{Title="TC-EMP007-QM001";  TC_EmpID="EMP-007"; TC_EmpName="KJ";               TC_DocID="QM-001";  TC_DocTitle="Quality Manual";         TC_Revision="Rev A"; TC_Method="Read & Understand"; TC_TrainDate="2026-04-17"; TC_RecordID="TR-013"}
    )
    foreach ($c in $completions) { Add-Item "QMS_TrainingCompletions" $c }

    Write-Host "`n[OK] All baseline data seeded successfully." -ForegroundColor Green
}

# =============================================================================
# SUMMARY
# =============================================================================
Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host " SCRIPT 05 COMPLETE" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " 11 SharePoint lists created/verified"
Write-Host " All columns provisioned"
if ($SeedData) {
    Write-Host " Baseline IMP9177 data seeded:"
    Write-Host "   QMS_Config         — 1 configuration row"
    Write-Host "   QMS_Roles          — 7 roles"
    Write-Host "   QMS_Employees      — 8 employees"
    Write-Host "   QMS_Approvers      — 4 approvers"
    Write-Host "   QMS_ChangeRequests — 3 CRs (CR-0001, CR-0002, CR-0003)"
    Write-Host "   QMS_DCOs           — 3 DCOs (DCO-0001, DCO-0002, DCO-0003)"
    Write-Host "   QMS_DCOApprovals   — 8 approval records"
    Write-Host "   QMS_RoutingHistory — 10 history entries"
    Write-Host "   QMS_TrainingMatrix — 19 role x document requirements"
    Write-Host "   QMS_TrainingCompletions — 13 signed training records"
} else {
    Write-Host ""
    Write-Host " To seed IMP9177 baseline data, run:"
    Write-Host "   .\05_Provision_Lists.ps1 -SeedData" -ForegroundColor White
}
Write-Host "`n Next step: SPFx web part scaffolding`n"
