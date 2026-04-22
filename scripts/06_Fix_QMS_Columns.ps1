# ============================================================
# IMP9177 QMS List Column Repair Script
# Run while connected to IMP9177 site
# ============================================================
$ErrorActionPreference = "Continue"

function Ensure-Field {
    param([string]$List, [string]$Name, [string]$Type, [string[]]$Choices = @())
    $existing = Get-PnPField -List $List -Identity $Name -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        if ($Type -eq "Choice") {
            Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Choice -Choices $Choices -AddToDefaultView | Out-Null
        } elseif ($Type -eq "Boolean") {
            Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Boolean -AddToDefaultView | Out-Null
        } elseif ($Type -eq "DateTime") {
            Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type DateTime -AddToDefaultView | Out-Null
        } elseif ($Type -eq "Number") {
            Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Number -AddToDefaultView | Out-Null
        } elseif ($Type -eq "Note") {
            Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Note -AddToDefaultView | Out-Null
        } else {
            Add-PnPField -List $List -DisplayName $Name -InternalName $Name -Type Text -AddToDefaultView | Out-Null
        }
        Write-Host "  + $Name ($Type)" -ForegroundColor Green
    } else {
        Write-Host "  ~ $Name (exists)" -ForegroundColor DarkYellow
    }
}

Write-Host "`n[QMS_Approvers]" -ForegroundColor Cyan
Ensure-Field "QMS_Approvers" "Appr_Role"        "Text"
Ensure-Field "QMS_Approvers" "Appr_Type"        "Choice" @("Required","Conditional","Optional")
Ensure-Field "QMS_Approvers" "Appr_Scope"       "Text"
Ensure-Field "QMS_Approvers" "Appr_SigningMode"  "Choice" @("Parallel","Sequential")
Ensure-Field "QMS_Approvers" "Appr_Active"       "Boolean"
Ensure-Field "QMS_Approvers" "Appr_DCOID"        "Text"
Ensure-Field "QMS_Approvers" "Appr_Name"         "Text"
Ensure-Field "QMS_Approvers" "Appr_Status"       "Choice" @("Waiting","Signed","Blocked","Not Required")
Ensure-Field "QMS_Approvers" "Appr_SigID"        "Text"

Write-Host "`n[QMS_DCOs]" -ForegroundColor Cyan
Ensure-Field "QMS_DCOs" "DCO_Phase"        "Choice" @("Draft","Submitted","In Review","Implemented","Awaiting Training","Effective")
Ensure-Field "QMS_DCOs" "DCO_Title"        "Text"
Ensure-Field "QMS_DCOs" "DCO_CRLink"       "Text"
Ensure-Field "QMS_DCOs" "DCO_SubmittedDate" "DateTime"
Ensure-Field "QMS_DCOs" "DCO_Originator"   "Text"
Ensure-Field "QMS_DCOs" "DCO_Docs"         "Note"
Ensure-Field "QMS_DCOs" "DCO_TrainingGate" "Text"
Ensure-Field "QMS_DCOs" "DCO_LateDays"     "Number"

Write-Host "`n[QMS_ChangeRequests]" -ForegroundColor Cyan
Ensure-Field "QMS_ChangeRequests" "CR_Title"       "Text"
Ensure-Field "QMS_ChangeRequests" "CR_Status"      "Choice" @("Draft","In Review","Approved","Linked to DCO","Closed")
Ensure-Field "QMS_ChangeRequests" "CR_Priority"    "Choice" @("Critical","High","Medium","Low")
Ensure-Field "QMS_ChangeRequests" "CR_Originator"  "Text"
Ensure-Field "QMS_ChangeRequests" "CR_LinkedDCOs"  "Text"
Ensure-Field "QMS_ChangeRequests" "CR_Description" "Note"
Ensure-Field "QMS_ChangeRequests" "CR_CreatedDate" "DateTime"

Write-Host "`n[QMS_DCOApprovals]" -ForegroundColor Cyan
Ensure-Field "QMS_DCOApprovals" "Appr_DCOID"      "Text"
Ensure-Field "QMS_DCOApprovals" "Appr_Name"       "Text"
Ensure-Field "QMS_DCOApprovals" "Appr_Role"       "Text"
Ensure-Field "QMS_DCOApprovals" "Appr_Type"       "Choice" @("Required","Conditional","Optional")
Ensure-Field "QMS_DCOApprovals" "Appr_Status"     "Choice" @("Waiting","Signed","Blocked","Not Required")
Ensure-Field "QMS_DCOApprovals" "Appr_SignedDate" "DateTime"
Ensure-Field "QMS_DCOApprovals" "Appr_SigID"      "Text"

Write-Host "`n[QMS_Records]" -ForegroundColor Cyan
Ensure-Field "QMS_Records" "Rec_Type"        "Choice" @("Batch Record","CoA Log","Training Record","CAPA Record","Pest Control Log","Deviation Report","Supplier Record")
Ensure-Field "QMS_Records" "Rec_Title"       "Text"
Ensure-Field "QMS_Records" "Rec_Originator"  "Text"
Ensure-Field "QMS_Records" "Rec_Reviewer"    "Text"
Ensure-Field "QMS_Records" "Rec_CreatedDate" "DateTime"
Ensure-Field "QMS_Records" "Rec_SigID"       "Text"

Write-Host "`n[QMS_Employees]" -ForegroundColor Cyan
Ensure-Field "QMS_Employees" "Emp_Email"  "Text"
Ensure-Field "QMS_Employees" "Emp_Title"  "Text"
Ensure-Field "QMS_Employees" "Emp_Roles"  "Note"

Write-Host "`n[QMS_Roles]" -ForegroundColor Cyan
Ensure-Field "QMS_Roles" "Role_Desc"        "Note"
Ensure-Field "QMS_Roles" "Role_RequiredDocs" "Note"

Write-Host "`n[QMS_TrainingMatrix]" -ForegroundColor Cyan
Ensure-Field "QMS_TrainingMatrix" "TM_RoleID"   "Text"
Ensure-Field "QMS_TrainingMatrix" "TM_DocID"    "Text"
Ensure-Field "QMS_TrainingMatrix" "TM_Required" "Boolean"

Write-Host "`n[QMS_TrainingCompletions]" -ForegroundColor Cyan
Ensure-Field "QMS_TrainingCompletions" "TC_EmpID"     "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_DocID"     "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_Rev"       "Text"
Ensure-Field "QMS_TrainingCompletions" "TC_SignedDate" "DateTime"
Ensure-Field "QMS_TrainingCompletions" "TC_SigID"     "Text"

Write-Host "`n[QMS_RoutingHistory]" -ForegroundColor Cyan
Ensure-Field "QMS_RoutingHistory" "RH_DCOID"     "Text"
Ensure-Field "QMS_RoutingHistory" "RH_EventType" "Choice" @("stage","signature","rejection","system","comment")
Ensure-Field "QMS_RoutingHistory" "RH_Stage"     "Text"
Ensure-Field "QMS_RoutingHistory" "RH_Actor"     "Text"
Ensure-Field "QMS_RoutingHistory" "RH_Note"      "Note"
Ensure-Field "QMS_RoutingHistory" "RH_Reason"    "Note"
Ensure-Field "QMS_RoutingHistory" "RH_Timestamp" "DateTime"

Write-Host "`n[QMS_Config]" -ForegroundColor Cyan
Ensure-Field "QMS_Config" "Cfg_Value" "Text"

# ── Seed Approvers ──────────────────────────────────────────
Write-Host "`n[Seeding Approvers]" -ForegroundColor Cyan
$approvers = @(
    @{Title="Andre Butler";   Appr_Name="Andre Butler";   Appr_Role="QA/Regulatory Consultant"; Appr_Type="Required";    Appr_Scope="All Documents"; Appr_SigningMode="Parallel"; Appr_Active=$true},
    @{Title="Tina Qin";       Appr_Name="Tina Qin";       Appr_Role="QA Approver (Interim)";     Appr_Type="Required";    Appr_Scope="All Documents"; Appr_SigningMode="Parallel"; Appr_Active=$true},
    @{Title="Liu Hong";       Appr_Name="Liu Hong";       Appr_Role="Management";                Appr_Type="Required";    Appr_Scope="All Documents"; Appr_SigningMode="Parallel"; Appr_Active=$true},
    @{Title="QD Yang";        Appr_Name="QD Yang";        Appr_Role="QA Director";               Appr_Type="Conditional"; Appr_Scope="SOPs only";     Appr_SigningMode="Parallel"; Appr_Active=$true},
    @{Title="Baolong";        Appr_Name="Baolong";        Appr_Role="Production";                Appr_Type="Optional";    Appr_Scope="Production SOPs";Appr_SigningMode="Parallel"; Appr_Active=$true}
)

foreach ($a in $approvers) {
    $existing = Get-PnPListItem -List "QMS_Approvers" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($a.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "QMS_Approvers" -Values $a | Out-Null
        Write-Host "  + $($a.Title)" -ForegroundColor Green
    } else {
        Write-Host "  ~ $($a.Title) (exists)" -ForegroundColor DarkYellow
    }
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host " QMS LIST REPAIR COMPLETE" -ForegroundColor Green
Write-Host " All columns added. Approvers seeded." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
