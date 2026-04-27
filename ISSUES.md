# IMP9177 QMS Portal — Issue Tracker
Last updated: 2026-04-26 (session4)

## Open Issues

### ISSUE-001 — Training tab: employee names show ID not full name
**Status:** Open  
**Priority:** Low  
**Location:** `openDCOInline` training tab lazy-load block  
**Description:** The training table renders `TM_EmpID` (the employee ID string from QMS_TrainingMatrix) rather than the employee's full name. A join against `this._data.employees` by `Emp_FullName`/`Title` match is needed to display human-readable names.  
**Fix:** In the `_rows11.map(...)` lambda, look up `this._data.employees.find((e:any) => e.Title === r.TM_EmpID)?.Emp_FullName || r.TM_EmpID`.

### ISSUE-002 — Training tab: requires DCO_DocumentIDs field populated
**Status:** Open  
**Priority:** Medium  
**Location:** `openDCOInline` training tab lazy-load block  
**Description:** The training filter uses `dco.DCO_DocumentIDs`. If only `DCO_Docs` is populated on a DCO item the training tab will show "No documents associated with this DCO." Needs fallback to `DCO_Docs` if `DCO_DocumentIDs` is empty.  
**Fix:** `const _raw = dco.DCO_DocumentIDs || dco.DCO_Docs || '';`

### ISSUE-003 — Duplicate Approver Assignments UI in Administration tab
**Status:** Open  
**Priority:** Low  
**Location:** `_renderAdministration()`  
**Description:** The standalone "Approver Assignments" panel in the Approvers sub-tab was not removed when the new inline "Approver For" section per employee was added in session4. Both UIs exist simultaneously, which may confuse users.  
**Fix:** Remove the standalone panel HTML and its save/wire logic from the Approvers sub-tab section.

### ISSUE-004 — AI P&J: anthropic-dangerous-direct-browser-access may be blocked by CSP
**Status:** Open  
**Priority:** Medium  
**Location:** AI P&J button handler  
**Description:** Direct browser calls to api.anthropic.com require the SharePoint site's Content Security Policy to allow connections to that host. If the CSP blocks it, the fetch will fail silently or throw a network error. The current error handler shows a toast, but the root cause is CSP not the API key.  
**Fix:** Verify CSP headers at tenant/site level. Alternatively route through a Power Automate flow or Azure Function proxy.

### ISSUE-005 — Submit blocks if DCO has no document types matching any active approver
**Status:** Open  
**Priority:** Medium  
**Location:** Submit handler, Task 3 approver auto-assign  
**Description:** If all documents in the DCO are of type "OTHER" (no QM/SOP/FM/FPS prefix) and no approvers have been manually added, the submit is blocked with a toast. This is correct behaviour but may surprise users adding non-standard doc IDs.  
**Fix:** Document the convention, or add a fallback approver type "ALL" that matches any DCO.

### ISSUE-006 — PDF download audit trail not using _auditLog()
**Status:** Open  
**Priority:** Low  
**Location:** `openDCOInline` — `_dlWrap.querySelectorAll("[data-dl-url]")` handler  
**Description:** The DOCX/PDF download buttons in Effective DCOs still post directly to QMS_RoutingHistory via `spHttpClient.post` rather than going through `_auditLog()`. Inconsistent with the rest of the audit trail.  
**Fix:** Replace the inline spHttpClient.post with `this._auditLog('download', dcoId, ...)`.

## Resolved Issues

### RESOLVED-001 — _pubFM used before declaration (TS2448)
**Session:** session4  
**Description:** `_pubFM` was referenced in the `_auditLog()` call that was placed before the `const _pubFM` declaration in the same function scope. TypeScript strict mode flagged this.  
**Fix:** Moved the `_auditLog()` call to after `_pubFM` is declared, and changed to use the already-computed `docFileName` variable.

### RESOLVED-002 — unknown[] not assignable to string[] (TS2322)
**Session:** session4  
**Description:** `[...new Set(...map(...))]` spread returned `unknown[]` because TypeScript could not infer the Set generic parameter.  
**Fix:** Changed to `Array.from(new Set<string>(...))` with explicit type parameter.

### RESOLVED-003 — QMS_Approvers not loaded in _data on initial load
**Session:** session4  
**Description:** `_data.approvers` was undefined because `QMS_Approvers` was only fetched on-demand in `_renderApprovers()` rather than in the global `_loadAll()` Promise.all.  
**Fix:** Added `approvers` to the `_loadAll()` Promise.all with fields `Id,Title,Appr_Name,Approver_Email,Appr_DocType,Appr_Active`.
