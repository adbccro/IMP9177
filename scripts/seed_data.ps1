# ============================================================
# IMP9177 SharePoint Seed Data Script
# Run AFTER provision_lists.ps1 completes successfully
# ============================================================

$ErrorActionPreference = "Continue"
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " IMP9177 Baseline Data Seeding" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan

# ============================================================
# MRS_Documents — W1 + W2 deliverables
# ============================================================
Write-Host "`n[1] Seeding MRS_Documents..." -ForegroundColor Cyan
$docs = @(
    @{Title="QM-001";DocNumber="QM-001";DocTitle="Quality Manual";DocType="Quality Manual";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-QMS-001";DocNumber="SOP-QMS-001";DocTitle="Management Responsibility SOP";DocType="SOP";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-DC-001";DocNumber="SOP-DC-001";DocTitle="Document Control SOP";DocType="SOP";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-TR-001";DocNumber="SOP-TR-001";DocTitle="Training and Competency SOP";DocType="SOP";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-CAPA-001";DocNumber="SOP-CAPA-001";DocTitle="CAPA SOP";DocType="SOP";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-INT-001";DocNumber="SOP-INT-001";DocTitle="Internal Audit SOP";DocType="SOP";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-CHG-001";DocNumber="SOP-CHG-001";DocTitle="Change Control SOP";DocType="SOP";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-COMP-001";DocNumber="SOP-COMP-001";DocTitle="Complaint Handling SOP";DocType="SOP";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-REC-001";DocNumber="SOP-REC-001";DocTitle="Records Management SOP";DocType="SOP";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="FM-001";DocNumber="FM-001";DocTitle="Document Change Order Form";DocType="Form";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="FM-002";DocNumber="FM-002";DocTitle="Training Record Form";DocType="Form";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="FM-003";DocNumber="FM-003";DocTitle="CAPA Form";DocType="Form";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="FM-004";DocNumber="FM-004";DocTitle="Internal Audit Checklist";DocType="Form";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="FM-027";DocNumber="FM-027";DocTitle="Management Review Minutes";DocType="Form";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="DCO-0001";DocNumber="DCO-0001";DocTitle="Document Change Order -- W1 Package";DocType="DCO";DocWeek="W1";DocRev="A";DocStatus="PAST DUE -- Sign-Off Overdue";ADBComplete="Yes"},
    @{Title="SOP-SUP-001";DocNumber="SOP-SUP-001";DocTitle="Supplier Qualification SOP";DocType="SOP";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="SOP-PC-001";DocNumber="SOP-PC-001";DocTitle="Pest Control SOP";DocType="SOP";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="SOP-ALG-001";DocNumber="SOP-ALG-001";DocTitle="Allergen Control SOP";DocType="SOP";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="SOP-SAN-001";DocNumber="SOP-SAN-001";DocTitle="Sanitation and Hygiene SOP";DocType="SOP";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="SOP-LAB-001";DocNumber="SOP-LAB-001";DocTitle="Laboratory Controls SOP";DocType="SOP";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="SOP-REJ-001";DocNumber="SOP-REJ-001";DocTitle="Nonconforming Material SOP";DocType="SOP";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="SOP-HARPC-001";DocNumber="SOP-HARPC-001";DocTitle="HARPC / Food Safety Plan SOP";DocType="SOP";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="FM-005";DocNumber="FM-005";DocTitle="Supplier Qualification Checklist";DocType="Form";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="FM-007";DocNumber="FM-007";DocTitle="Allergen Risk Assessment Form";DocType="Form";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="FM-008";DocNumber="FM-008";DocTitle="Supplier CoA Requirements Checklist";DocType="Form";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="FM-009";DocNumber="FM-009";DocTitle="Nonconforming Material Report";DocType="Form";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="FM-010";DocNumber="FM-010";DocTitle="Pest Control Log";DocType="Form";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"},
    @{Title="DCO-0002";DocNumber="DCO-0002";DocTitle="Document Change Order -- W2 Package";DocType="DCO";DocWeek="W2";DocRev="A";DocStatus="Pending Feedback";ADBComplete="Yes"}
)
foreach ($d in $docs) {
    $existing = Get-PnPListItem -List "MRS_Documents" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($d.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "MRS_Documents" -Values $d | Out-Null
        Write-Host "  + $($d.Title)" -ForegroundColor DarkGreen
    } else {
        Write-Host "  ~ $($d.Title) (exists)" -ForegroundColor DarkYellow
    }
}

# ============================================================
# MRS_CAPALog
# ============================================================
Write-Host "`n[2] Seeding MRS_CAPALog..." -ForegroundColor Cyan
$capas = @(
    @{Title="CAPA-0001";CAPANumber="CAPA-0001";CAPAType="CAPA";CAPADesc="Systemic CoA Deficiency -- Supplier Qualification Program Remediation. All active production ingredients require CoA on file per 21 CFR 111.75(a)(1).";CAPAStatus="OPEN -- In Progress";CAPAOwner="3H -- Tina Qin"}
)
foreach ($c in $capas) {
    $existing = Get-PnPListItem -List "MRS_CAPALog" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($c.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "MRS_CAPALog" -Values $c | Out-Null
        Write-Host "  + $($c.Title)" -ForegroundColor DarkGreen
    }
}

# ============================================================
# RAID_Actions — AC-001 through AC-039
# ============================================================
Write-Host "`n[3] Seeding RAID_Actions..." -ForegroundColor Cyan
$actions = @(
    @{Title="AC-023";ActionCategory="Compliance";ActionDesc="Request Orkin stored product pest addendum for facility. GAP-PC-03. Critical gap.";ActionOwner="3H -- Tina Qin";ActionPriority="Critical";ActionStatus="Open -- Urgent"},
    @{Title="AC-029";ActionCategory="Document Control";ActionDesc="Issue DCO-0001 W1 sign-off package to 3H for execution. Tina Qin + Liu Hong signatures required.";ActionOwner="3H -- Tina Qin";ActionPriority="Critical";ActionStatus="PAST DUE"},
    @{Title="AC-034";ActionCategory="Compliance";ActionDesc="QD Yang work authorization -- determine signatory role. BLOCKING DCO-0001 execution.";ActionOwner="ADB (Andre)";ActionPriority="Critical";ActionStatus="BLOCKING"},
    @{Title="AC-035";ActionCategory="Compliance";ActionDesc="ISS-010 Melatonin Powder CoA -- obtain from supplier immediately. Active production ingredient.";ActionOwner="3H -- Tina Qin";ActionPriority="Critical";ActionStatus="Open -- Critical"},
    @{Title="AC-006";ActionCategory="Administration";ActionDesc="Warehouse inventory count -- complete physical count and submit to ADB.";ActionOwner="3H -- KJ";ActionPriority="High";ActionStatus="PAST DUE"},
    @{Title="AC-028";ActionCategory="Training";ActionDesc="Training matrix -- complete for all 3H personnel per SOP-TR-001.";ActionOwner="3H -- Tina Qin";ActionPriority="High";ActionStatus="In Progress"},
    @{Title="AC-030";ActionCategory="Document Control";ActionDesc="W2 feedback -- review 12-document package and return comments by Apr 14.";ActionOwner="3H -- Tina Qin";ActionPriority="High";ActionStatus="In Progress"},
    @{Title="AC-031";ActionCategory="Supplier";ActionDesc="Globe Gums Inc -- initiate supplier qualification per SOP-SUP-001.";ActionOwner="3H -- Tina Qin";ActionPriority="High";ActionStatus="Open"},
    @{Title="AC-032";ActionCategory="Supplier";ActionDesc="Globe Gums -- obtain CoA for Agar Lot A8020251110001 currently on-site.";ActionOwner="3H -- Tina Qin";ActionPriority="High";ActionStatus="Open"},
    @{Title="AC-033";ActionCategory="Supplier";ActionDesc="Globe Gums -- obtain allergen statement / Letter of Guarantee (LOG) for Big 9.";ActionOwner="3H -- Tina Qin";ActionPriority="High";ActionStatus="Open"},
    @{Title="AC-036";ActionCategory="Administration";ActionDesc="PSO-CO-001 Phase 2 Change Order -- execute to authorize Phase 2B.";ActionOwner="3H -- Liu Hong";ActionPriority="High";ActionStatus="Open"},
    @{Title="AC-037";ActionCategory="Compliance";ActionDesc="QD Yang -- add to FM-027 Management Review signatories and QM-001 Section 5.";ActionOwner="ADB (Andre)";ActionPriority="Medium";ActionStatus="In Progress"},
    @{Title="AC-038";ActionCategory="Administration";ActionDesc="MM-006 -- schedule week of Apr 21 after Tina travel Apr 13-17.";ActionOwner="ADB (Andre)";ActionPriority="Medium";ActionStatus="Open"},
    @{Title="AC-039";ActionCategory="Document Control";ActionDesc="W3 deliverables planning -- finalize scope and timeline for next package.";ActionOwner="ADB (Andre)";ActionPriority="Medium";ActionStatus="Open"}
)
foreach ($a in $actions) {
    $existing = Get-PnPListItem -List "RAID_Actions" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($a.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "RAID_Actions" -Values $a | Out-Null
        Write-Host "  + $($a.Title)" -ForegroundColor DarkGreen
    } else {
        Write-Host "  ~ $($a.Title) (exists)" -ForegroundColor DarkYellow
    }
}

# ============================================================
# RAID_Issues
# ============================================================
Write-Host "`n[4] Seeding RAID_Issues..." -ForegroundColor Cyan
$issues = @(
    @{Title="ISS-010";IssueDescription="Melatonin Powder -- No Certificate of Analysis on file. Active production ingredient. Violates 21 CFR 111.75(a)(1). Supplier unknown.";RegulationRef="21 CFR 111.75(a)(1)";IssueSeverity="Critical";IssueStatus="Open -- Critical";IssueOwner="3H -- Tina Qin";LinkedCAPARef="CAPA-0001"},
    @{Title="ISS-002";IssueDescription="Active stored product pest infestation identified during Phase 1 gap assessment. Orkin contract in place but stored product addendum missing.";RegulationRef="21 CFR 111.15(a); FSMA PC Rule";IssueSeverity="Critical";IssueStatus="Partially Mitigated";IssueOwner="3H Management";LinkedCAPARef=""},
    @{Title="ISS-005";IssueDescription="Supplier qualification program non-existent at time of Phase 1 assessment. 18 active production suppliers with no approved supplier list.";RegulationRef="21 CFR 111.70(b)";IssueSeverity="Critical";IssueStatus="Open";IssueOwner="3H -- Tina Qin";LinkedCAPARef="CAPA-0001"},
    @{Title="ISS-007";IssueDescription="Training records -- no formal documented training program or competency records for any 3H personnel at time of assessment.";RegulationRef="21 CFR 111.13; 111.14";IssueSeverity="Major";IssueStatus="Open";IssueOwner="3H -- Tina Qin";LinkedCAPARef=""}
)
foreach ($i in $issues) {
    $existing = Get-PnPListItem -List "RAID_Issues" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($i.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "RAID_Issues" -Values $i | Out-Null
        Write-Host "  + $($i.Title)" -ForegroundColor DarkGreen
    } else {
        Write-Host "  ~ $($i.Title) (exists)" -ForegroundColor DarkYellow
    }
}

# ============================================================
# RAID_Decisions
# ============================================================
Write-Host "`n[5] Seeding RAID_Decisions..." -ForegroundColor Cyan
$decisions = @(
    @{Title="DEC-035";MeetingRef="MM-005";DecisionDesc="Amazon-sourced ingredients excluded from quality scope. R&D/development use only. 18 production vendors confirmed as in-scope.";MadeBy="Andre Butler + Tina Qin";LinkedActionRefs=""},
    @{Title="DEC-036";MeetingRef="MM-005";DecisionDesc="Cindy Dong confirmed as Document Control Coordinator and primary operational contact.";MadeBy="Andre Butler + Tina Qin";LinkedActionRefs=""},
    @{Title="DEC-037";MeetingRef="MM-005";DecisionDesc="Globe Gums Inc agar must be placed on supplier qualification hold pending SQ completion and CoA receipt.";MadeBy="Andre Butler + Tina Qin";LinkedActionRefs="AC-031,AC-032,AC-033"},
    @{Title="DEC-038";MeetingRef="MM-005";DecisionDesc="No meetings week of Apr 13-17 due to Tina travel. MM-006 scheduled week of Apr 21.";MadeBy="Andre Butler + Tina Qin";LinkedActionRefs="AC-038"},
    @{Title="DEC-039";MeetingRef="MM-005";DecisionDesc="ISS-010 Melatonin Powder -- treat as Priority 1 open issue. CoA must be obtained before next production run.";MadeBy="Andre Butler + Tina Qin";LinkedActionRefs="AC-035"},
    @{Title="DEC-040";MeetingRef="MM-005";DecisionDesc="Tina Qin confirmed as interim QA Approver and sole document signatory. QD Yang assigned operational QA duties only -- not a document signer.";MadeBy="Andre Butler + Tina Qin + Liu Hong";LinkedActionRefs="AC-034,AC-037"},
    @{Title="DEC-018";MeetingRef="MM-004";DecisionDesc="One CR and one DCO per weekly deliverable set. Batch documents within weekly cycle.";MadeBy="Andre Butler";LinkedActionRefs=""},
    @{Title="DEC-028";MeetingRef="MM-004";DecisionDesc="All preexisting 3H SOPs treated as informal drafts. Formally issued at Rev A under DCO-0001.";MadeBy="Andre Butler + Tina Qin";LinkedActionRefs=""}
)
foreach ($d in $decisions) {
    $existing = Get-PnPListItem -List "RAID_Decisions" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($d.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "RAID_Decisions" -Values $d | Out-Null
        Write-Host "  + $($d.Title)" -ForegroundColor DarkGreen
    } else {
        Write-Host "  ~ $($d.Title) (exists)" -ForegroundColor DarkYellow
    }
}

# ============================================================
# RAID_Meetings
# ============================================================
Write-Host "`n[6] Seeding RAID_Meetings..." -ForegroundColor Cyan
$meetings = @(
    @{Title="MM-001";MeetingType="Kickoff";Platform="Teams";Duration=45;AgendaSummary="Project kickoff. Scope review. Phase 1 gap assessment initiation. Key personnel introductions.";ActionsRaised=5;DecisionsMade=3},
    @{Title="MM-002";MeetingType="Status Update";Platform="Teams";Duration=30;AgendaSummary="Phase 1 gap assessment preliminary findings. Facility walkthrough results. Supplier qualification gaps identified.";ActionsRaised=4;DecisionsMade=2},
    @{Title="MM-003";MeetingType="Project Review";Platform="Teams";Duration=40;AgendaSummary="Phase 1 gap assessment final report review. 77 gaps confirmed. Phase 2A scope and timeline established.";ActionsRaised=6;DecisionsMade=4},
    @{Title="MM-004";MeetingType="Project Review";Platform="Teams";Duration=29;AgendaSummary="W1 deliverables orientation. SharePoint access. CR/DCO workflow explained. Document review process confirmed.";ActionsRaised=3;DecisionsMade=5;MeetingNotes="SharePoint access resent to Cindy and Tina. Training matrix action assigned (AC-028)."},
    @{Title="MM-005";MeetingType="Project Review";Platform="Teams";Duration=35;AgendaSummary="W1/W2 package review. Supplier controls. Amazon exclusion confirmed. Tina Qin as interim QA Approver. ISS-010 Melatonin escalated.";ActionsRaised=9;DecisionsMade=6;MeetingNotes="9 new actions AC-029 through AC-037. DEC-035 through DEC-040 logged. Globe Gums qualification hold."}
)
foreach ($m in $meetings) {
    $existing = Get-PnPListItem -List "RAID_Meetings" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($m.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "RAID_Meetings" -Values $m | Out-Null
        Write-Host "  + $($m.Title)" -ForegroundColor DarkGreen
    } else {
        Write-Host "  ~ $($m.Title) (exists)" -ForegroundColor DarkYellow
    }
}

# ============================================================
# PM_Milestones
# ============================================================
Write-Host "`n[7] Seeding PM_Milestones..." -ForegroundColor Cyan
$milestones = @(
    @{Title="M-1";MilestonePhase="Phase 2A";MilestoneDesc="MILESTONE 1: Phase 2A Complete -- All W1-W4 documents signed off";MilestoneStatus="In Progress";MilestonePct=44;MilestoneTarget="Apr 22, 2026"},
    @{Title="W1-DEL";MilestonePhase="Phase 2A";MilestoneDesc="W1 -- 15-Document Package Delivered to 3H";MilestoneStatus="Complete";MilestonePct=100;MilestoneTarget="Mar 31, 2026"},
    @{Title="W1-SO";MilestonePhase="Phase 2A";MilestoneDesc="CR-0001 / DCO-0001 -- 3H Sign-Off Obtained";MilestoneStatus="At Risk -- PAST DUE";MilestonePct=0;MilestoneTarget="Apr 5, 2026"},
    @{Title="W2-DEL";MilestonePhase="Phase 2A";MilestoneDesc="W2 -- 12-Document Package Delivered to 3H";MilestoneStatus="Complete";MilestonePct=100;MilestoneTarget="Apr 9, 2026"},
    @{Title="MM-005";MilestonePhase="Phase 2A";MilestoneDesc="MM-005 -- W1/W2 Review Meeting";MilestoneStatus="Complete";MilestonePct=100;MilestoneTarget="Apr 8, 2026"},
    @{Title="W2-SO";MilestonePhase="Phase 2A";MilestoneDesc="DCO-0002 -- W2 3H Sign-Off";MilestoneStatus="In Progress";MilestonePct=0;MilestoneTarget="Apr 21, 2026"},
    @{Title="MM-006";MilestonePhase="Phase 2A";MilestoneDesc="MM-006 -- W2 Feedback + W3 Kickoff";MilestoneStatus="Not Started";MilestonePct=0;MilestoneTarget="Apr 21, 2026"},
    @{Title="PSO-CO";MilestonePhase="Phase 2A";MilestoneDesc="PSO-CO-001 -- Phase 2 Change Order Execution";MilestoneStatus="In Progress";MilestonePct=0;MilestoneTarget="Apr 22, 2026"}
)
foreach ($m in $milestones) {
    $existing = Get-PnPListItem -List "PM_Milestones" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($m.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "PM_Milestones" -Values $m | Out-Null
        Write-Host "  + $($m.Title)" -ForegroundColor DarkGreen
    } else {
        Write-Host "  ~ $($m.Title) (exists)" -ForegroundColor DarkYellow
    }
}

# ============================================================
# PM_OpenItems
# ============================================================
Write-Host "`n[8] Seeding PM_OpenItems..." -ForegroundColor Cyan
$openItems = @(
    @{Title="ISS-010";OIRef="ISS-010";OITitle="Melatonin Powder -- No CoA. Active production ingredient. 21 CFR 111.75(a)(1).";OIOwner="3H -- Tina Qin";OIPriority="Critical";OIStatus="Open -- Critical"},
    @{Title="DCO-0001";OIRef="DCO-0001";OITitle="W1 Sign-off PAST DUE. Tina Qin + Liu Hong must sign. QD Yang blocking issue (AC-034).";OIOwner="3H -- Tina Qin";OIPriority="Critical";OIStatus="Open -- Critical"},
    @{Title="AC-023";OIRef="AC-023";OITitle="Orkin stored product pest addendum -- URGENT. Required before any regulatory inspection.";OIOwner="3H -- Tina Qin";OIPriority="Critical";OIStatus="Open -- Urgent"},
    @{Title="PSO-CO-001";OIRef="PSO-CO-001";OITitle="Phase 2 Change Order pending execution. Phase 2B cannot begin without signed CO.";OIOwner="3H -- Liu Hong";OIPriority="High";OIStatus="In Progress"},
    @{Title="W2-FB";OIRef="W2-FB";OITitle="W2 12-document package feedback due. Tina review pending after Apr 17 travel.";OIOwner="3H -- Tina Qin";OIPriority="High";OIStatus="In Progress"}
)
foreach ($o in $openItems) {
    $existing = Get-PnPListItem -List "PM_OpenItems" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($o.Title)</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
    if ($null -eq $existing -or $existing.Count -eq 0) {
        Add-PnPListItem -List "PM_OpenItems" -Values $o | Out-Null
        Write-Host "  + $($o.Title)" -ForegroundColor DarkGreen
    } else {
        Write-Host "  ~ $($o.Title) (exists)" -ForegroundColor DarkYellow
    }
}

# ============================================================
# PM_Budget
# ============================================================
Write-Host "`n[9] Seeding PM_Budget..." -ForegroundColor Cyan
$budgetExisting = Get-PnPListItem -List "PM_Budget" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>IMP9177-Budget</Value></Eq></Where></Query></View>" -ErrorAction SilentlyContinue
if ($null -eq $budgetExisting -or $budgetExisting.Count -eq 0) {
    Add-PnPListItem -List "PM_Budget" -Values @{
        Title="IMP9177-Budget"
        BudgetCeiling=25000
        BudgetSpent=4617
        BudgetGapsAdded=0
        BudgetInvoice="INV-IMP9177I002"
        BudgetInv1=4617
        BudgetInv2=0
        BudgetInv2Ref=""
    } | Out-Null
    Write-Host "  + IMP9177-Budget" -ForegroundColor DarkGreen
} else {
    Write-Host "  ~ IMP9177-Budget (exists)" -ForegroundColor DarkYellow
}

Write-Host "`n============================================================" -ForegroundColor Cyan
Write-Host " SEED DATA COMPLETE" -ForegroundColor Green
Write-Host " All baseline IMP9177 records loaded." -ForegroundColor Cyan
Write-Host " Web parts will now display live data from SharePoint lists." -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
