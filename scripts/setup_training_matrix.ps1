# setup_training_matrix.ps1
# Seeds QMS_TrainingMatrix with all roles x all documents (required=true by default)
# Uses correct field names: TM_RoleID, TM_DocID, TM_DocTitle, TM_DocType, TM_Required
#
# Run from SPFx root after Connect-PnPOnline

$SiteUrl = "https://adbccro.sharepoint.com/sites/IMP9177"

# Roles from QMS_Roles list
$roles = @(
    @{ ID="1"; Name="Management" },
    @{ ID="2"; Name="QA / Quality Approver" },
    @{ ID="3"; Name="R&D / QA Director" },
    @{ ID="4"; Name="Production Operator" },
    @{ ID="5"; Name="Warehouse / Receiving" },
    @{ ID="6"; Name="QC Technician" },
    @{ ID="7"; Name="Project Manager" }
)

# All 27 QMS documents
$docs = @(
    @{ ID="QM-001";      Title="Quality Manual";                        Type="QM"  },
    @{ ID="SOP-QMS-001"; Title="Management Responsibility";             Type="SOP" },
    @{ ID="SOP-QMS-002"; Title="Document Control";                      Type="SOP" },
    @{ ID="SOP-QMS-003"; Title="Change Control";                        Type="SOP" },
    @{ ID="SOP-SUP-001"; Title="Supplier Qualification";                Type="SOP" },
    @{ ID="SOP-SUP-002"; Title="Receiving Inspection";                  Type="SOP" },
    @{ ID="SOP-FS-001";  Title="Allergen Control";                      Type="SOP" },
    @{ ID="SOP-FS-002";  Title="Equipment Cleaning";                    Type="SOP" },
    @{ ID="SOP-FS-003";  Title="Facility Sanitation";                   Type="SOP" },
    @{ ID="SOP-FS-004";  Title="Environmental Monitoring";              Type="SOP" },
    @{ ID="SOP-PC-001";  Title="Pest Sighting Response";                Type="SOP" },
    @{ ID="SOP-PRD-108"; Title="Finished Product Release Management";   Type="SOP" },
    @{ ID="SOP-PRD-432"; Title="Finished Product Specifications";       Type="SOP" },
    @{ ID="SOP-FRS-549"; Title="Product Specification Sheet";           Type="SOP" },
    @{ ID="SOP-RCL-321"; Title="Product Recall Procedure";              Type="SOP" },
    @{ ID="FPS-001";     Title="Lychee VD3 Gummy Finished Product Spec";Type="FPS" },
    @{ ID="FM-001";      Title="Master Document Log";                   Type="FM"  },
    @{ ID="FM-002";      Title="Change Request Form";                   Type="FM"  },
    @{ ID="FM-003";      Title="Document Change Order Form";            Type="FM"  },
    @{ ID="FM-004";      Title="Approved Supplier List";                Type="FM"  },
    @{ ID="FM-005";      Title="Receiving Log";                         Type="FM"  },
    @{ ID="FM-006";      Title="Raw Material Specification Sheet";      Type="FM"  },
    @{ ID="FM-007";      Title="Material Hold Label";                   Type="FM"  },
    @{ ID="FM-027";      Title="QU/QS Designation Record";              Type="FM"  },
    @{ ID="FM-030";      Title="Finished Product Spec Sheet Template";  Type="FM"  },
    @{ ID="FM-ALG";      Title="Allergen Status Record";                Type="FM"  },
    @{ ID="FM-008";      Title="Supplier CoA Requirements Checklist";   Type="FM"  }
)

Write-Host "Loading existing QMS_TrainingMatrix rows..."
$existing = Get-PnPListItem -List "QMS_TrainingMatrix" -PageSize 500 |
    ForEach-Object { "$($_.FieldValues.TM_RoleID)|$($_.FieldValues.TM_DocID)" }
$existingSet = @{}
foreach ($e in $existing) { $existingSet[$e] = $true }
Write-Host "  Found $($existingSet.Count) existing rows"

$added   = 0
$skipped = 0

foreach ($role in $roles) {
    foreach ($doc in $docs) {
        $key = "$($role.ID)|$($doc.ID)"
        if ($existingSet.ContainsKey($key)) {
            $skipped++
            continue
        }
        try {
            Add-PnPListItem -List "QMS_TrainingMatrix" -Values @{
                Title         = "$($role.Name) — $($doc.ID)"
                TM_RoleID     = $role.ID
                TM_RoleName   = $role.Name
                TM_DocID      = $doc.ID
                TM_DocTitle   = $doc.Title
                TM_DocType    = $doc.Type
                TM_Required   = $true
            } | Out-Null
            Write-Host "  [OK] $($role.Name) — $($doc.ID)"
            $added++
        } catch {
            Write-Host "  [ERR] $($role.Name) — $($doc.ID) : $_" -ForegroundColor Red
        }
    }
}

Write-Host ""
Write-Host "Done. Added: $added  Skipped (already exist): $skipped"
