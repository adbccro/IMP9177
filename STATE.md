# IMP9177 QMS Portal — Project State
Last updated: 2026-04-26 (session4)

## Deployment
- SP site: https://adbccro.sharepoint.com/sites/IMP9177
- App: imp-9177-spfx.sppkg (built, not yet deployed this session — requires PnP interactive auth)
- Last deployed build: session3 (Tasks 0-13)
- Last committed build: session4 (commit 5f7a776)

## Main File
`src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts`

## SharePoint Lists (all in QMS_ namespace)
| List | Key Fields |
|------|------------|
| QMS_DCOs | Id, Title, DCO_Title, DCO_Phase, DCO_Originator, DCO_Docs, DCO_DocumentIDs, DCO_CRLink, DCO_EffectiveDate, DCO_DocsLastUpdated |
| QMS_DCOApprovals | Id, Appr_DCOID, Appr_Name, Appr_Email, Appr_Status, Appr_SignedDate, Appr_Role |
| QMS_Approvers | Id, Title, Appr_Name, Approver_Email, Appr_DocType, Appr_Active |
| QMS_RoutingHistory | Id, Title, RH_DCOID, RH_EventType, RH_Stage, RH_Actor, RH_Note, RH_Timestamp |
| QMS_Employees | Id, Title, Emp_FullName, Emp_Email, Emp_Dept, Emp_Status, Emp_PortalRole |
| QMS_Config | Id, Title (key), Cfg_Value |
| QMS_Documents | Id, Title, Doc_Title, Doc_Status, Doc_Category, Doc_RevisionLevel |
| QMS_CRs | Id, Title, CR_Title, CR_Status, CR_Description |
| QMS_TrainingMatrix | Id, TM_EmpID, TM_DocID, TM_Required |
| QMS_TrainingCompletions | Id, TC_EmpID, TC_DocID, TC_CompletedDate, TC_Score, TC_Status |

## Roles (Emp_PortalRole)
- PM — full access (all tabs visible)
- ADB User — full access (all tabs visible)
- Change Analyst — can view/edit DCOs, no admin/publish/config tabs
- Observer — read-only, no admin/publish/config tabs
- External — read-only

## Feature Status (session4 complete)

### Implemented
- [x] Identity resolution: `_currentUser` from QMS_Employees matched by SP user email
- [x] Amber warning banner for users not in QMS_Employees
- [x] `_applyPermissions()` — hides publish/administration/config tabs for non-PM/ADB-User
- [x] Approver auto-assign on DCO submit — email-dedup, blocks if no approvers found
- [x] Inline "Approver For" section per employee row in Administration tab (with Edit/Save checkboxes)
- [x] AI P&J — live-fetch API key from QMS_Config (key: `anthropicApiKey`), `anthropic-dangerous-direct-browser-access: true` header
- [x] DCO edit CR Link — multi-select dropdown from `_data.crs`
- [x] Submit optimistic update — button → "Submitted" (green, disabled), modal closes after 1.5s
- [x] Doc open handler uses `_auditLog()` for DocumentOpened events
- [x] Dashboard draft/active count separation; db-bc1 counts current-user-involved DCOs
- [x] Training tab in DCO modal — lazy-loads QMS_TrainingMatrix + QMS_TrainingCompletions
  - Per-doc progress bar (% complete)
  - Per-employee status table (Status, Completed date, Score)

### Scripts
- [x] `scripts/cleanup_session4.ps1` — reset DCO-0001 to Draft, delete approvals/history, deduplicate employees
- [x] `scripts/setup_training_matrix.ps1` — seed QMS_TrainingMatrix per active employee x 26 documents

## Known Limitations / Not Yet Implemented
- Training tab shows employees by `TM_EmpID` (Title field value from QMS_TrainingMatrix). Full names require joining with QMS_Employees.
- `_drmT` (doc ID → title map) used in training tab but may be sparse — falls back to doc ID if no title found.
- `DCO_DocumentIDs` field used in training tab filter; must be populated for training to load. If only `DCO_Docs` is populated, training tab shows "No documents associated."
- Separate "Approver Assignments" panel in Administration > Approvers sub-tab still exists alongside the new inline section.
- PDF download in Effective DCOs still uses a direct inline SP REST call (not `_auditLog`) for the download audit event.

## Build System
- Toolchain: heft 0.x, TypeScript 5.8.3, Webpack 5, SPFx 1.18
- Clean: `./node_modules/.bin/heft clean`
- Build: `./node_modules/.bin/heft build --production`
- Package: `./node_modules/.bin/heft package-solution --production`
- Output: `sharepoint/solution/imp-9177-spfx.sppkg`

## Git
- Remote: https://github.com/adbccro/IMP9177.git
- Branch: main
- Last commit: 5f7a776
