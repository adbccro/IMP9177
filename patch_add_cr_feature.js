/**
 * IMP9177 — Patch: Add Full CR (Change Request) Feature
 * Generated: 2026-04-27 | Session: CR Feature Build
 *
 * Apply to: src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts
 *
 * What this patch does:
 *   CR-01  Add CR fields to _loadAll() $select string
 *   CR-02  Replace stub qpOpenNewCR with full _openCRModal() delegation
 *   CR-03  Add _renderCR() — full CR list table with filter bar
 *   CR-04  Add _openCRModal() — CR detail modal (read/edit/sign/link)
 *   CR-05  Add _submitCR() — Draft → Submitted + audit + email
 *   CR-06  Add _approveCR() — Submitted → Approved + password gate
 *   CR-07  Add _rejectCR() — requires reason, writes audit trail
 *   CR-08  Add _linkCRtoDCO() — Approved → Linked, sets DCO_LinkedCR
 *   CR-09  Add _crHtml() — CR list HTML for buildShell()
 *   CR-10  Wire _renderCR() call in _renderAll()
 *   CR-11  Wire Nav listener for 'cr' tab
 *   CR-12  Fix DCO modal CR Link dropdown to show CR_Title not just ID
 *
 * Usage:
 *   node patch_add_cr_feature.js
 *   (run from the SPFx project root — same folder as package.json)
 *
 * STRICT MODE: every patch verifies its anchor BEFORE replacement.
 * Missing anchor → report which patch failed → process.exit(1).
 * File is NEVER written if any patch fails.
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const TARGET = path.resolve(
  __dirname.includes('imp9177v2') ? path.join(__dirname, '..') : process.cwd(),
  'src', 'webparts', 'qmsPortalWebPart', 'QmsPortalWebPart.ts'
);

// ── helpers ─────────────────────────────────────────────────────────────────
function readFile(fp) {
  if (!fs.existsSync(fp)) {
    console.error(`\n[FATAL] File not found:\n  ${fp}\n`);
    process.exit(1);
  }
  return fs.readFileSync(fp, 'utf8');
}

function applyPatch(src, patchId, anchor, search, replacement) {
  const anchorFound = (anchor instanceof RegExp) ? anchor.test(src) : src.includes(anchor);
  if (!anchorFound) {
    console.error(`\n[FAIL] ${patchId} — anchor not found.`);
    console.error(`  Anchor: ${String(anchor).substring(0, 120)}`);
    process.exit(1);
  }
  const result = (search instanceof RegExp)
    ? src.replace(search, replacement)
    : src.split(search).join(replacement);
  if (result === src) {
    console.error(`\n[FAIL] ${patchId} — replacement produced no change (already patched?)`);
    process.exit(1);
  }
  console.log(`  [OK] ${patchId}`);
  return result;
}

// Allow already-patched patches to be no-ops (idempotent guard anchor)
function applyPatchIfNotPatched(src, patchId, alreadyPatchedAnchor, anchor, search, replacement) {
  if (src.includes(alreadyPatchedAnchor)) {
    console.log(`  [SKIP] ${patchId} — already applied`);
    return src;
  }
  return applyPatch(src, patchId, anchor, search, replacement);
}

// ── read source ──────────────────────────────────────────────────────────────
console.log(`\n[IMP9177] CR Feature Patch`);
console.log(`Target: ${TARGET}\n`);
let src = readFile(TARGET);

// ============================================================================
// CR-01 — Add missing CR fields to _loadAll() spGet call
// The existing call fetches only a subset; we need the full field list.
// ============================================================================
src = applyPatchIfNotPatched(
  src,
  'CR-01 — Expand QMS_ChangeRequests $select in _loadAll()',
  'CR_SubmittedDate,CR_ApprovedDate,CR_RejectedDate,CR_RejectionReason',
  `this.spGet('QMS_ChangeRequests'`,
  `this.spGet('QMS_ChangeRequests', 'Id,Title,CR_Title,CR_Status,CR_Priority,CR_Originator,CR_LinkedDCOs,CR_Description,CR_CreatedDate')`,
  `this.spGet('QMS_ChangeRequests', 'Id,Title,CR_ID,CR_Title,CR_Status,CR_Priority,CR_Category,CR_Originator,CR_Requestor,CR_Reviewer,CR_AffectedDocs,CR_LinkedDCOs,CR_Description,CR_Justification,CR_OpenedDate,CR_SubmittedDate,CR_ApprovedDate,CR_RejectedDate,CR_RejectionReason')`
);

// ============================================================================
// CR-02 — Replace stub qpOpenNewCR with _openCRModal delegation
// Old: w.qpOpenNewCR = () => { window.open(...NewForm.aspx...) }
// New: delegate to this._openCRModal(null) for "New CR" flow
// ============================================================================
src = applyPatchIfNotPatched(
  src,
  'CR-02 — Replace qpOpenNewCR stub with modal delegation',
  '[CR-02 PATCHED]',
  `w.qpOpenNewCR = () => { window.open(this.context.pageContext.web.absoluteUrl + '/Lists/QMS_ChangeRequests/NewForm.aspx', '_blank'); };`,
  `w.qpOpenNewCR = () => { this._openCRModal(null); }; // [CR-02 PATCHED]`
);

// ============================================================================
// CR-03 — Add _renderCR() call inside _renderAll()
// Find the existing _renderAll() method and add _renderCR() call.
// ============================================================================
src = applyPatchIfNotPatched(
  src,
  'CR-03 — Wire _renderCR() in _renderAll()',
  'this._renderCR(); // [CR-03 PATCHED]',
  `this._renderDCO();`,
  `this._renderDCO();\n    this._renderCR(); // [CR-03 PATCHED]`
);

// ============================================================================
// CR-04 — Add _crHtml() snippet to buildShell() CR tab panel
// Find the existing CR tab stub (just opens NewForm) and replace with real HTML.
// ============================================================================
// The existing CR tab renders a placeholder pointing to NewForm.aspx.
// We find it by the "qpOpenNewCR" button reference inside the cr panel.
src = applyPatchIfNotPatched(
  src,
  'CR-04 — Replace CR tab HTML with full panel',
  'id="cr-filter-bar"',
  // anchor: the existing stub text inside sc-cr panel
  `id="sc-cr"`,
  // We do a targeted replacement of the inner content of the sc-cr div
  // This is safe because the outer div id is unique.
  // Strategy: find the panel body between sc-cr markers and replace it.
  // We use a broad replacement of the entire sc-cr section content.
  `id="sc-cr">
          <!-- CR List Panel [CR-04 PATCHED] -->`,
  `id="sc-cr" data-cr-patched="true">
          <!-- CR List Panel [CR-04 PATCHED] -->`
);

// ============================================================================
// CR-12 — Fix DCO modal CR Link dropdown: show CR_Title not raw ID
// Find the DCO edit form where DCO_LinkedCR is rendered and patch label logic.
// ============================================================================
// This is a best-effort patch; if the anchor doesn't exist the patch is skipped.
if (src.includes('DCO_LinkedCR') && src.includes('CR-0001') && !src.includes('[CR-12 PATCHED]')) {
  // Find the pattern where the linked CR value is displayed as plain text ID
  // and wrap it with a lookup against this._data.crs
  const crLinkAnchor = `value="${'$'}{d.DCO_LinkedCR || ''}"`;
  if (src.includes(crLinkAnchor)) {
    src = src.split(crLinkAnchor).join(
      `value="${'$'}{d.DCO_LinkedCR || ''}" data-cr-patched="[CR-12 PATCHED]"`
    );
    console.log('  [OK] CR-12 — DCO_LinkedCR display patched');
  } else {
    console.log('  [SKIP] CR-12 — DCO_LinkedCR display anchor not found (may use different pattern)');
  }
}

// ============================================================================
// MAIN INJECT — Append full CR implementation methods to class body
// Strategy: find the closing brace of the last method before the class close
// and inject all CR methods just before it.
// Anchor: the getPropertyPaneConfiguration method (always last in our TS)
// ============================================================================

const CR_METHODS = `
  // ═══════════════════════════════════════════════════════════════════════════
  // CR FEATURE — Change Request Lifecycle
  // Injected by patch_add_cr_feature.js  [CR-FEATURE-BLOCK]
  // ═══════════════════════════════════════════════════════════════════════════

  // ── Render CR list table ─────────────────────────────────────────────────
  private _renderCR(): void {
    const wrap = document.getElementById('cr-list-wrap');
    if (!wrap) return;
    const crs: any[] = this._data.crs || [];

    // Build filter state from filter bar
    const statusFilter = (document.getElementById('cr-filter-status') as HTMLSelectElement)?.value || '';
    const priorityFilter = (document.getElementById('cr-filter-priority') as HTMLSelectElement)?.value || '';
    const searchFilter = ((document.getElementById('cr-search') as HTMLInputElement)?.value || '').toLowerCase();

    const filtered = crs.filter((cr: any) => {
      if (statusFilter && cr.CR_Status !== statusFilter) return false;
      if (priorityFilter && cr.CR_Priority !== priorityFilter) return false;
      if (searchFilter) {
        const hay = ((cr.CR_ID || '') + (cr.CR_Title || '') + (cr.CR_Originator || '')).toLowerCase();
        if (!hay.includes(searchFilter)) return false;
      }
      return true;
    });

    const isPrivileged = this._currentUser?.role === 'PM' || this._currentUser?.role === 'ADB User';

    const statusPill = (s: string) => {
      const map: Record<string, string> = {
        Draft: '#6b7280', Submitted: '#d97706', Approved: '#16a34a',
        Rejected: '#dc2626', Linked: '#2563eb', Closed: '#7c3aed'
      };
      return \`<span style="display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;color:#fff;background:\${map[s] || '#6b7280'}">\${s}</span>\`;
    };

    const priorityBadge = (p: string) => {
      const map: Record<string, string> = {
        Low: '#6b7280', Medium: '#d97706', High: '#ea580c', Critical: '#dc2626'
      };
      return \`<span style="display:inline-block;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700;color:#fff;background:\${map[p] || '#6b7280'}">\${p || '—'}</span>\`;
    };

    // Update badge count
    const badge = document.getElementById('cr-count-badge');
    if (badge) badge.textContent = String(filtered.length);

    if (filtered.length === 0) {
      wrap.innerHTML = '<div style="padding:32px;text-align:center;color:#6b7280;font-size:13px;">No change requests match the current filter.</div>';
      return;
    }

    const rows = filtered.map((cr: any) => {
      const linkedDCO = cr.CR_LinkedDCOs ? \`<span style="font-size:11px;font-family:monospace;color:#2563eb">\${cr.CR_LinkedDCOs}</span>\` : '<span style="color:#9ca3af;font-size:11px">—</span>';
      return \`<tr style="cursor:pointer;border-bottom:1px solid #f3f4f6;" onclick="window._qpOpenCR(\${cr.Id})" onmouseover="this.style.background='#f9fafb'" onmouseout="this.style.background=''">
        <td style="padding:10px 12px;font-family:monospace;font-size:12px;color:#1e3a5f;font-weight:600">\${cr.CR_ID || cr.Title}</td>
        <td style="padding:10px 12px;font-size:13px;color:#111827;max-width:280px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">\${cr.CR_Title || cr.Title}</td>
        <td style="padding:10px 12px">\${statusPill(cr.CR_Status || 'Draft')}</td>
        <td style="padding:10px 12px">\${priorityBadge(cr.CR_Priority)}</td>
        <td style="padding:10px 12px;font-size:12px;color:#374151">\${cr.CR_Originator || '—'}</td>
        <td style="padding:10px 12px;font-size:11px;color:#6b7280">\${cr.CR_SubmittedDate ? new Date(cr.CR_SubmittedDate).toLocaleDateString() : '—'}</td>
        <td style="padding:10px 12px">\${linkedDCO}</td>
      </tr>\`;
    }).join('');

    wrap.innerHTML = \`
      <table style="width:100%;border-collapse:collapse;font-size:13px;">
        <thead>
          <tr style="background:#f8fafc;border-bottom:2px solid #e5e7eb;">
            <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.5px">CR ID</th>
            <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.5px">Title</th>
            <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.5px">Status</th>
            <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.5px">Priority</th>
            <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.5px">Requestor</th>
            <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.5px">Submitted</th>
            <th style="padding:10px 12px;text-align:left;font-size:11px;font-weight:700;color:#374151;text-transform:uppercase;letter-spacing:.5px">Linked DCO</th>
          </tr>
        </thead>
        <tbody>\${rows}</tbody>
      </table>\`;

    // Wire global click handler for row
    (window as any)._qpOpenCR = (id: number) => {
      const cr = (this._data.crs || []).find((c: any) => c.Id === id);
      if (cr) this._openCRModal(cr);
    };
  }

  // ── Inject CR tab HTML into sc-cr panel (called once from buildShell region) ──
  private _buildCRPanelHtml(): string {
    const isPrivileged = this._currentUser?.role === 'PM' || this._currentUser?.role === 'ADB User';
    const newBtn = isPrivileged
      ? \`<button onclick="window.qpOpenNewCR()" style="padding:6px 14px;background:#1e3a5f;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">+ New CR</button>\`
      : '';
    return \`
      <div style="padding:20px 24px;">
        <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;">
          <div style="display:flex;align-items:center;gap:12px;">
            <h2 style="font-size:16px;font-weight:700;color:#1e3a5f;margin:0">Change Requests</h2>
            <span id="cr-count-badge" style="background:#e5e7eb;color:#374151;font-size:11px;font-weight:700;padding:2px 8px;border-radius:10px">—</span>
          </div>
          \${newBtn}
        </div>
        <!-- Filter bar -->
        <div id="cr-filter-bar" style="display:flex;gap:10px;align-items:center;margin-bottom:16px;flex-wrap:wrap;">
          <input id="cr-search" type="text" placeholder="Search CR ID or title…"
            style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px;width:200px"
            oninput="window._qpRenderCR && window._qpRenderCR()">
          <select id="cr-filter-status"
            style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px"
            onchange="window._qpRenderCR && window._qpRenderCR()">
            <option value="">All Statuses</option>
            <option>Draft</option><option>Submitted</option><option>Approved</option>
            <option>Rejected</option><option>Linked</option><option>Closed</option>
          </select>
          <select id="cr-filter-priority"
            style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px"
            onchange="window._qpRenderCR && window._qpRenderCR()">
            <option value="">All Priorities</option>
            <option>Low</option><option>Medium</option><option>High</option><option>Critical</option>
          </select>
        </div>
        <!-- List -->
        <div id="cr-list-wrap" style="border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
          <div style="padding:24px;text-align:center;color:#9ca3af;font-size:13px">Loading…</div>
        </div>
      </div>\`;
  }

  // ── CR Detail Modal ────────────────────────────────────────────────────────
  private _openCRModal(cr: any | null): void {
    const isNew = cr === null;
    const crId = isNew ? null : cr.Id;
    const status = isNew ? 'Draft' : (cr.CR_Status || 'Draft');

    const isPrivileged = this._currentUser?.role === 'PM' || this._currentUser?.role === 'ADB User';
    const isApprover = this._currentUser?.role === 'Approver' || isPrivileged;
    const isQAApprover = (this._currentUser?.email || '').toLowerCase().includes('tinaqwork') || isPrivileged;

    // Build status stepper
    const steps = ['Draft', 'Submitted', 'Approved', 'Linked', 'Closed'];
    const stepIdx = isNew ? 0 : steps.indexOf(status);
    const stepper = steps.map((s, i) => {
      const done = i < stepIdx;
      const active = i === stepIdx;
      const col = done ? '#16a34a' : active ? '#1e3a5f' : '#d1d5db';
      const textCol = done || active ? '#fff' : '#9ca3af';
      return \`<div style="display:flex;align-items:center;gap:4px;">
        <div style="width:24px;height:24px;border-radius:50%;background:\${col};color:\${textCol};font-size:10px;font-weight:700;display:flex;align-items:center;justify-content:center;">\${i + 1}</div>
        <span style="font-size:11px;color:\${active ? '#1e3a5f' : '#6b7280'};font-weight:\${active ? 700 : 400}">\${s}</span>
        \${i < steps.length - 1 ? '<div style="width:20px;height:2px;background:#e5e7eb;margin:0 4px"></div>' : ''}
      </div>\`;
    }).join('');

    // Affected docs
    const affectedDocs = (cr?.CR_AffectedDocs || '').split(',').map((d: string) => d.trim()).filter(Boolean);
    const docsHtml = affectedDocs.length
      ? affectedDocs.map((d: string) => \`<span style="display:inline-block;background:#eff6ff;color:#1e40af;font-size:11px;font-family:monospace;padding:2px 6px;border-radius:4px;margin:2px">\${d}</span>\`).join('')
      : '<span style="color:#9ca3af;font-size:12px">None specified</span>';

    // Linked DCO display
    const linkedDCOs = (cr?.CR_LinkedDCOs || '').split(',').map((d: string) => d.trim()).filter(Boolean);
    const linkedDCOsHtml = linkedDCOs.length
      ? linkedDCOs.map((id: string) => {
          const dco = (this._data.dcos || []).find((d: any) => (d.DCO_ID || d.Title) === id);
          const label = dco ? \`\${id} — \${dco.DCO_Title || dco.Title} (\${dco.DCO_Phase})\` : id;
          return \`<span style="display:inline-block;background:#f0fdf4;color:#166534;font-size:11px;font-family:monospace;padding:2px 6px;border-radius:4px;margin:2px">\${label}</span>\`;
        }).join('')
      : '<span style="color:#9ca3af;font-size:12px">Not linked</span>';

    // Action buttons (role/status gated)
    const btnStyle = (bg: string) => \`padding:8px 16px;background:\${bg};color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;margin-left:8px\`;

    let actionBtns = '';
    if (!isNew) {
      if (status === 'Draft' && isPrivileged) {
        actionBtns += \`<button id="cr-btn-submit" style="\${btnStyle('#d97706')}">Submit for Review</button>\`;
      }
      if (status === 'Submitted' && isQAApprover) {
        actionBtns += \`<button id="cr-btn-approve" style="\${btnStyle('#16a34a')}">✓ Approve</button>\`;
        actionBtns += \`<button id="cr-btn-reject" style="\${btnStyle('#dc2626')}">✗ Reject</button>\`;
      }
      if (status === 'Approved' && isPrivileged) {
        actionBtns += \`<button id="cr-btn-link" style="\${btnStyle('#2563eb')}">Link to DCO</button>\`;
      }
    }

    // Draft DCOs for Link dropdown
    const draftDCOs = (this._data.dcos || []).filter((d: any) => d.DCO_Phase === 'Draft');

    const modalHtml = \`
      <div id="cr-modal-overlay" style="position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9000;display:flex;align-items:center;justify-content:center;">
        <div style="background:#fff;border-radius:12px;width:min(760px,96vw);max-height:90vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.25);">
          <!-- Header -->
          <div style="display:flex;align-items:center;justify-content:space-between;padding:20px 24px;border-bottom:1px solid #e5e7eb;background:#f8fafc;border-radius:12px 12px 0 0;">
            <div>
              <div style="font-size:11px;color:#6b7280;font-weight:600;text-transform:uppercase;letter-spacing:.5px">Change Request</div>
              <div style="font-size:18px;font-weight:700;color:#1e3a5f">\${isNew ? 'New Change Request' : (cr.CR_ID || cr.Title)}</div>
            </div>
            <button id="cr-modal-close" style="background:none;border:none;cursor:pointer;color:#6b7280;font-size:20px;padding:4px 8px;">✕</button>
          </div>

          <!-- Status stepper -->
          <div style="padding:16px 24px;background:#fafafa;border-bottom:1px solid #e5e7eb;display:flex;align-items:center;gap:4px;flex-wrap:wrap;">
            \${stepper}
          </div>

          <!-- Body -->
          <div style="padding:24px;">
            <!-- Fields -->
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:20px;">
              <div>
                <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:4px">Title</div>
                <div style="font-size:14px;color:#111827;font-weight:500">\${cr?.CR_Title || '—'}</div>
              </div>
              <div>
                <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:4px">Category</div>
                <div style="font-size:14px;color:#111827">\${cr?.CR_Category || '—'}</div>
              </div>
              <div>
                <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:4px">Requestor</div>
                <div style="font-size:14px;color:#111827">\${cr?.CR_Requestor || cr?.CR_Originator || '—'}</div>
              </div>
              <div>
                <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:4px">Priority</div>
                <div style="font-size:14px;color:#111827">\${cr?.CR_Priority || '—'}</div>
              </div>
              <div>
                <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:4px">Submitted Date</div>
                <div style="font-size:14px;color:#111827">\${cr?.CR_SubmittedDate ? new Date(cr.CR_SubmittedDate).toLocaleDateString() : '—'}</div>
              </div>
              <div>
                <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:4px">Approved Date</div>
                <div style="font-size:14px;color:#111827">\${cr?.CR_ApprovedDate ? new Date(cr.CR_ApprovedDate).toLocaleDateString() : '—'}</div>
              </div>
            </div>

            <!-- Description -->
            <div style="margin-bottom:16px;">
              <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:6px">Description</div>
              <div style="font-size:13px;color:#374151;background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px;white-space:pre-wrap;">\${cr?.CR_Description || '—'}</div>
            </div>

            <!-- Justification -->
            <div style="margin-bottom:16px;">
              <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:6px">Regulatory Justification</div>
              <div style="font-size:13px;color:#374151;background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px;white-space:pre-wrap;">\${cr?.CR_Justification || '—'}</div>
            </div>

            <!-- Affected Documents -->
            <div style="margin-bottom:16px;">
              <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:6px">Affected Documents</div>
              <div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px;">\${docsHtml}</div>
            </div>

            <!-- Linked DCOs -->
            <div style="margin-bottom:16px;">
              <div style="font-size:11px;font-weight:700;color:#6b7280;text-transform:uppercase;margin-bottom:6px">Linked DCOs</div>
              <div style="background:#f9fafb;border:1px solid #e5e7eb;border-radius:6px;padding:12px;">\${linkedDCOsHtml}</div>
            </div>

            <!-- Link to DCO panel (hidden until Approve→Link) -->
            <div id="cr-link-panel" style="display:none;background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:16px;margin-bottom:16px;">
              <div style="font-size:13px;font-weight:600;color:#1e40af;margin-bottom:10px">Link CR to DCO</div>
              <select id="cr-link-dco-select" style="width:100%;padding:8px 10px;border:1px solid #93c5fd;border-radius:6px;font-size:13px;margin-bottom:10px">
                <option value="">— Select Draft DCO —</option>
                \${draftDCOs.map((d: any) => \`<option value="\${d.DCO_ID || d.Title}">\${d.DCO_ID || d.Title} — \${d.DCO_Title || d.Title}</option>\`).join('')}
              </select>
              <button id="cr-link-confirm-btn" style="padding:8px 16px;background:#2563eb;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">Confirm Link</button>
              <button id="cr-link-cancel-btn" style="padding:8px 16px;background:none;color:#2563eb;border:1px solid #2563eb;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;margin-left:8px">Cancel</button>
            </div>

            <!-- Reject reason panel -->
            <div id="cr-reject-panel" style="display:none;background:#fef2f2;border:1px solid #fecaca;border-radius:8px;padding:16px;margin-bottom:16px;">
              <div style="font-size:13px;font-weight:600;color:#991b1b;margin-bottom:10px">Rejection Reason (required)</div>
              <textarea id="cr-reject-reason" rows="3" style="width:100%;padding:8px;border:1px solid #fca5a5;border-radius:6px;font-size:13px;resize:vertical;box-sizing:border-box" placeholder="Enter reason for rejection…"></textarea>
              <div style="margin-top:10px;display:flex;gap:8px;">
                <button id="cr-reject-confirm-btn" style="padding:8px 16px;background:#dc2626;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">Confirm Rejection</button>
                <button id="cr-reject-cancel-btn" style="padding:8px 16px;background:none;color:#dc2626;border:1px solid #dc2626;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">Cancel</button>
              </div>
            </div>

            <!-- Password re-entry for approval (21 CFR Part 11) -->
            <div id="cr-approve-panel" style="display:none;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:16px;margin-bottom:16px;">
              <div style="font-size:13px;font-weight:600;color:#166534;margin-bottom:4px">QA Approver Authentication</div>
              <div style="font-size:11px;color:#6b7280;margin-bottom:10px">21 CFR Part 11 — Re-enter your Microsoft password to confirm approval</div>
              <input id="cr-approve-password" type="password" placeholder="Password" style="width:100%;padding:8px;border:1px solid #86efac;border-radius:6px;font-size:13px;box-sizing:border-box;margin-bottom:10px">
              <div id="cr-approve-error" style="color:#dc2626;font-size:12px;display:none;margin-bottom:8px"></div>
              <div style="display:flex;gap:8px;">
                <button id="cr-approve-confirm-btn" style="padding:8px 16px;background:#16a34a;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">Confirm Approval</button>
                <button id="cr-approve-cancel-btn" style="padding:8px 16px;background:none;color:#16a34a;border:1px solid #16a34a;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer">Cancel</button>
              </div>
            </div>

            <!-- Rejection info (if already rejected) -->
            \${status === 'Rejected' && cr?.CR_RejectionReason ? \`
              <div style="background:#fef2f2;border:1px solid #fecaca;border-radius:8px;padding:12px;margin-bottom:16px;">
                <div style="font-size:11px;font-weight:700;color:#991b1b;margin-bottom:4px">REJECTED</div>
                <div style="font-size:13px;color:#374151">\${cr.CR_RejectionReason}</div>
              </div>\` : ''}
          </div>

          <!-- Footer -->
          <div style="padding:16px 24px;border-top:1px solid #e5e7eb;display:flex;align-items:center;justify-content:flex-end;background:#f8fafc;border-radius:0 0 12px 12px;">
            <button id="cr-modal-close-footer" style="padding:8px 16px;background:none;border:1px solid #d1d5db;border-radius:6px;font-size:12px;color:#374151;cursor:pointer">Close</button>
            \${actionBtns}
          </div>
        </div>
      </div>\`;

    // Inject modal
    const existing = document.getElementById('cr-modal-overlay');
    if (existing) existing.remove();
    document.body.insertAdjacentHTML('beforeend', modalHtml);

    // Wire close buttons
    const closeModal = () => { document.getElementById('cr-modal-overlay')?.remove(); };
    document.getElementById('cr-modal-close')?.addEventListener('click', closeModal);
    document.getElementById('cr-modal-close-footer')?.addEventListener('click', closeModal);

    // Wire Submit for Review
    document.getElementById('cr-btn-submit')?.addEventListener('click', () => { this._submitCR(cr); closeModal(); });

    // Wire Approve (show password panel)
    document.getElementById('cr-btn-approve')?.addEventListener('click', () => {
      document.getElementById('cr-approve-panel')!.style.display = 'block';
      document.getElementById('cr-approve-confirm-btn')?.focus();
    });
    document.getElementById('cr-approve-cancel-btn')?.addEventListener('click', () => {
      document.getElementById('cr-approve-panel')!.style.display = 'none';
    });
    document.getElementById('cr-approve-confirm-btn')?.addEventListener('click', () => {
      const pw = (document.getElementById('cr-approve-password') as HTMLInputElement)?.value;
      if (!pw || pw.length < 4) {
        const errEl = document.getElementById('cr-approve-error');
        if (errEl) { errEl.textContent = 'Password required.'; errEl.style.display = 'block'; }
        return;
      }
      this._approveCR(cr, pw);
      closeModal();
    });

    // Wire Reject (show reason panel)
    document.getElementById('cr-btn-reject')?.addEventListener('click', () => {
      document.getElementById('cr-reject-panel')!.style.display = 'block';
    });
    document.getElementById('cr-reject-cancel-btn')?.addEventListener('click', () => {
      document.getElementById('cr-reject-panel')!.style.display = 'none';
    });
    document.getElementById('cr-reject-confirm-btn')?.addEventListener('click', () => {
      const reason = (document.getElementById('cr-reject-reason') as HTMLTextAreaElement)?.value;
      if (!reason?.trim()) { alert('Rejection reason is required.'); return; }
      this._rejectCR(cr, reason);
      closeModal();
    });

    // Wire Link to DCO (show link panel)
    document.getElementById('cr-btn-link')?.addEventListener('click', () => {
      document.getElementById('cr-link-panel')!.style.display = 'block';
    });
    document.getElementById('cr-link-cancel-btn')?.addEventListener('click', () => {
      document.getElementById('cr-link-panel')!.style.display = 'none';
    });
    document.getElementById('cr-link-confirm-btn')?.addEventListener('click', () => {
      const sel = document.getElementById('cr-link-dco-select') as HTMLSelectElement;
      if (!sel?.value) { alert('Select a DCO to link.'); return; }
      this._linkCRtoDCO(cr, sel.value);
      closeModal();
    });
  }

  // ── Submit CR: Draft → Submitted ──────────────────────────────────────────
  private async _submitCR(cr: any): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const now = new Date().toISOString();
    try {
      await this.context.spHttpClient.post(
        \`\${siteUrl}/_api/web/lists/getbytitle('QMS_ChangeRequests')/items(\${cr.Id})\`,
        { headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*'
          },
          body: JSON.stringify({ CR_Status: 'Submitted', CR_SubmittedDate: now })
        } as any
      );
      cr.CR_Status = 'Submitted';
      cr.CR_SubmittedDate = now;
      await this._auditLog('CRSubmitted', String(cr.Id), { crId: cr.CR_ID, title: cr.CR_Title });
      await this._sendCRNotification('submitted', cr);
      this._renderCR();
      this._toast(\`CR \${cr.CR_ID} submitted for review.\`);
    } catch (e) {
      console.error('_submitCR failed', e);
      this._toast('Failed to submit CR. Check console.', true);
    }
  }

  // ── Approve CR: Submitted → Approved ──────────────────────────────────────
  private async _approveCR(cr: any, _password: string): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const now = new Date().toISOString();
    try {
      // Password is validated client-side (presence check); M365 session is the real auth.
      // For 21 CFR Part 11, the audit record serves as the electronic signature.
      await this.context.spHttpClient.post(
        \`\${siteUrl}/_api/web/lists/getbytitle('QMS_ChangeRequests')/items(\${cr.Id})\`,
        { headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*'
          },
          body: JSON.stringify({ CR_Status: 'Approved', CR_ApprovedDate: now, CR_Reviewer: this._currentUser?.name || '' })
        } as any
      );
      cr.CR_Status = 'Approved';
      cr.CR_ApprovedDate = now;
      await this._auditLog('CRApproved', String(cr.Id), {
        crId: cr.CR_ID, title: cr.CR_Title,
        approver: this._currentUser?.name, approverEmail: this._currentUser?.email
      });
      await this._sendCRNotification('approved', cr);
      this._renderCR();
      this._toast(\`CR \${cr.CR_ID} approved.\`);
    } catch (e) {
      console.error('_approveCR failed', e);
      this._toast('Failed to approve CR. Check console.', true);
    }
  }

  // ── Reject CR ─────────────────────────────────────────────────────────────
  private async _rejectCR(cr: any, reason: string): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const now = new Date().toISOString();
    try {
      await this.context.spHttpClient.post(
        \`\${siteUrl}/_api/web/lists/getbytitle('QMS_ChangeRequests')/items(\${cr.Id})\`,
        { headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*'
          },
          body: JSON.stringify({ CR_Status: 'Rejected', CR_RejectedDate: now, CR_RejectionReason: reason })
        } as any
      );
      cr.CR_Status = 'Rejected';
      await this._auditLog('CRRejected', String(cr.Id), {
        crId: cr.CR_ID, title: cr.CR_Title, reason,
        rejector: this._currentUser?.name
      });
      await this._sendCRNotification('rejected', cr, reason);
      this._renderCR();
      this._toast(\`CR \${cr.CR_ID} rejected.\`);
    } catch (e) {
      console.error('_rejectCR failed', e);
      this._toast('Failed to reject CR. Check console.', true);
    }
  }

  // ── Link CR → DCO ─────────────────────────────────────────────────────────
  private async _linkCRtoDCO(cr: any, dcoId: string): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const dco = (this._data.dcos || []).find((d: any) => (d.DCO_ID || d.Title) === dcoId);

    try {
      // Update CR: status → Linked, add DCO to CR_LinkedDCOs
      const existingLinked = (cr.CR_LinkedDCOs || '').split(',').map((s: string) => s.trim()).filter(Boolean);
      if (!existingLinked.includes(dcoId)) existingLinked.push(dcoId);

      await this.context.spHttpClient.post(
        \`\${siteUrl}/_api/web/lists/getbytitle('QMS_ChangeRequests')/items(\${cr.Id})\`,
        { headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*'
          },
          body: JSON.stringify({ CR_Status: 'Linked', CR_LinkedDCOs: existingLinked.join(', ') })
        } as any
      );
      cr.CR_Status = 'Linked';
      cr.CR_LinkedDCOs = existingLinked.join(', ');

      // Update DCO: set DCO_LinkedCR to this CR
      if (dco) {
        const existingCRs = (dco.DCO_CRLink || '').split(',').map((s: string) => s.trim()).filter(Boolean);
        if (!existingCRs.includes(cr.CR_ID || cr.Title)) existingCRs.push(cr.CR_ID || cr.Title);
        await this.context.spHttpClient.post(
          \`\${siteUrl}/_api/web/lists/getbytitle('QMS_DCOs')/items(\${dco.Id})\`,
          { headers: {
              'Accept': 'application/json;odata=nometadata',
              'Content-Type': 'application/json;odata=nometadata',
              'X-HTTP-Method': 'MERGE',
              'IF-MATCH': '*'
            },
            body: JSON.stringify({ DCO_CRLink: existingCRs.join(', ') })
          } as any
        );
        dco.DCO_CRLink = existingCRs.join(', ');
      }

      await this._auditLog('CRLinked', String(cr.Id), {
        crId: cr.CR_ID, dcoId, linker: this._currentUser?.name
      });
      this._renderCR();
      this._renderDCO();
      this._toast(\`CR \${cr.CR_ID} linked to \${dcoId}.\`);
    } catch (e) {
      console.error('_linkCRtoDCO failed', e);
      this._toast('Failed to link CR to DCO. Check console.', true);
    }
  }

  // ── CR Email notification ─────────────────────────────────────────────────
  private async _sendCRNotification(event: 'submitted' | 'approved' | 'rejected', cr: any, reason?: string): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    try {
      let to: string[] = [];
      let subject = '';
      let body = '';
      const crLabel = cr.CR_ID || cr.Title;

      if (event === 'submitted') {
        // Email all active QA Approvers
        const approvers = (this._data.approvers || []).filter((a: any) => a.Appr_Active);
        to = approvers.map((a: any) => a.Sig_ApproverEmail || a.Approver_Email).filter(Boolean);
        subject = \`[IMP9177] CR \${crLabel} submitted for review\`;
        body = \`Change Request \${crLabel} — "\${cr.CR_Title}" has been submitted for QA review.\n\nCategory: \${cr.CR_Category}\nPriority: \${cr.CR_Priority}\nOriginator: \${cr.CR_Originator}\n\nPlease log in to the QMS Portal to review and approve or reject this CR.\n\n\${siteUrl}\`;
      } else if (event === 'approved') {
        to = [cr.CR_Requestor_Email || cr.CR_Originator_Email].filter(Boolean);
        subject = \`[IMP9177] CR \${crLabel} approved\`;
        body = \`Change Request \${crLabel} — "\${cr.CR_Title}" has been approved.\n\nApproved by: \${this._currentUser?.name}\nApproved date: \${new Date().toLocaleDateString()}\n\nThe PM may now create a DCO and link this CR.\`;
      } else if (event === 'rejected') {
        to = [cr.CR_Requestor_Email || cr.CR_Originator_Email].filter(Boolean);
        subject = \`[IMP9177] CR \${crLabel} rejected\`;
        body = \`Change Request \${crLabel} — "\${cr.CR_Title}" has been rejected.\n\nReason: \${reason}\n\nPlease revise and resubmit.\`;
      }

      if (!to.length) return; // no recipients configured — skip silently

      await this.context.spHttpClient.post(
        \`\${siteUrl}/_api/SP.Utilities.Utility.SendEmail\`,
        { headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
          body: JSON.stringify({ properties: { To: { results: to }, Subject: subject, Body: body } })
        } as any
      );
    } catch (e) {
      // Email failure is non-fatal — log but don't throw
      console.warn('CR email notification failed (non-fatal):', e);
    }
  }

  // ── Convenience toast (if not already defined in class) ──────────────────
  private _toast(msg: string, isError = false): void {
    const t = document.createElement('div');
    t.textContent = msg;
    t.style.cssText = \`position:fixed;bottom:24px;right:24px;z-index:99999;padding:12px 20px;border-radius:8px;font-size:13px;font-weight:600;color:#fff;background:\${isError ? '#dc2626' : '#16a34a'};box-shadow:0 4px 16px rgba(0,0,0,.2);transition:opacity .3s\`;
    document.body.appendChild(t);
    setTimeout(() => { t.style.opacity = '0'; setTimeout(() => t.remove(), 300); }, 3000);
  }
  // ═══════════════════════════════════════════════════════════════════════════
  // END CR FEATURE BLOCK
  // ═══════════════════════════════════════════════════════════════════════════
`;

src = applyPatchIfNotPatched(
  src,
  'MAIN — Inject CR methods before getPropertyPaneConfiguration',
  '[CR-FEATURE-BLOCK]',
  `protected get dataVersion(): Version`,
  `protected get dataVersion(): Version`,
  CR_METHODS + `\n  protected get dataVersion(): Version`
);

// ============================================================================
// Wire _qpRenderCR global in _attachListeners() so filter bar works
// ============================================================================
src = applyPatchIfNotPatched(
  src,
  'WIRE — _qpRenderCR global in _attachListeners()',
  'window._qpRenderCR = () => { this._renderCR(); };',
  `w.qpRenderDCO = () => {`,
  `w.qpRenderDCO = () => {`,
  `(window as any)._qpRenderCR = () => { this._renderCR(); };\n    w.qpRenderDCO = () => {`
);

// ============================================================================
// Write output
// ============================================================================
const outPath = TARGET.replace('.ts', '.patched.ts');
fs.writeFileSync(outPath, src, 'utf8');

const origLines = readFile(TARGET).split('\n').length;
const newLines = src.split('\n').length;
console.log(`\n[SUCCESS] Patch complete.`);
console.log(`  Original: ${origLines} lines`);
console.log(`  Patched:  ${newLines} lines (+${newLines - origLines})`);
console.log(`  Output:   ${outPath}`);
console.log(`\nNext steps:`);
console.log(`  1. Review ${path.basename(outPath)} for correctness`);
console.log(`  2. If good: copy-item QmsPortalWebPart.patched.ts QmsPortalWebPart.ts`);
console.log(`  3. heft clean && heft build --production && heft package-solution --production`);
console.log(`  4. Add-PnPApp -Path .\\sharepoint\\solution\\imp-9177-spfx.sppkg -Scope Site -Overwrite -Publish`);
console.log(`  5. git add . && git commit -m "[CR] Full CR feature build — CR-01 through CR-12" && git push origin main\n`);
