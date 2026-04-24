// patch_phase_aware.js
// Makes DCO document pane phase-aware:
//   In Review  → show Opened/Unopened review gate (existing behaviour)
//   Effective  → show DCO Report download + per-doc View/Download buttons
// Run from SPFx root: node patch_phase_aware.js

const fs = require('fs');
const path = require('path');

const FILE = path.join(__dirname, '..', 'imp9177-spfx', 'src', 'webparts', 'qmsPortalWebPart', 'QmsPortalWebPart.ts');

if (!fs.existsSync(FILE)) {
  // Try relative path (if run from SPFx root)
  const FILE2 = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
  if (fs.existsSync(FILE2)) {
    module.exports = FILE2;
  }
}

const filePath = fs.existsSync('./src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts')
  ? './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts'
  : FILE;

console.log('Reading:', filePath);
let src = fs.readFileSync(filePath, 'utf8');

let changed = 0;

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 1 — Insert phase flags + download rows HTML before const docRowsHtml
// ─────────────────────────────────────────────────────────────────────────────
const P1_OLD = `const docRowsHtml = displayDocs.map((docId: string, idx2: number) => {`;

const P1_NEW = `// ── Phase flags ──
      const isEffective = ['Implemented','Awaiting Training','Effective'].includes(dco.DCO_Phase || '');
      const isInReview   = !isEffective;

      // ── Official zone file map ──
      const _SP   = 'https://adbccro.sharepoint.com/sites/IMP9177';
      const _OFD  = _SP + '/Shared%20Documents/Official/QMS/Documents';
      const _OFF  = _SP + '/Shared%20Documents/Official/QMS/Forms';
      const _FIDS = ['FM-001','FM-002','FM-003','FM-004','FM-005','FM-006','FM-007','FM-008','FM-027','FM-030','FM-ALG'];
      const _FM: Record<string,string> = {
        'QM-001':'QM-001_Quality_Manual_RevA.docx',
        'SOP-QMS-001':'SOP-QMS-001_RevA_Management_Responsibility.docx',
        'SOP-QMS-002':'SOP-QMS-002_RevA_Document_Control.docx',
        'SOP-QMS-003':'SOP-QMS-003_RevA_Change_Control.docx',
        'SOP-PRD-108':'SOP-PRD-108_RevA.docx',
        'SOP-PRD-432':'SOP-PRD-432_RevA.docx',
        'SOP-FRS-549':'SOP-FRS-549_RevA.docx',
        'SOP-SUP-001':'SOP-SUP-001_RevA_Supplier_Qualification_FINAL.docx',
        'SOP-SUP-002':'SOP-SUP-002_RevA_Receiving_Inspection_FINAL.docx',
        'SOP-FS-001':'SOP-FS-001_RevA_Allergen_Control_FINAL.docx',
        'SOP-FS-002':'SOP-FS-002_RevA_Equipment_Cleaning_FINAL.docx',
        'SOP-FS-003':'SOP-FS-003_RevA_Facility_Sanitation_FINAL.docx',
        'SOP-FS-004':'SOP-FS-004_RevA_Environmental_Monitoring_FINAL.docx',
        'SOP-PC-001':'SOP-PC-001_RevA_Pest_Sighting_Response.docx',
        'FM-001':'FM-001_Master_Document_Log_RevA.docx',
        'FM-002':'FM-002_Change_Request_Form_RevA.docx',
        'FM-003':'FM-003_Document_Change_Order_RevA.docx',
        'FM-004':'FM-004_Supplier_Evaluation_Form_RevA.docx',
        'FM-005':'FM-005_Ingredient_Approval_Form_RevA.docx',
        'FM-006':'FM-006_Material_Receipt_Log_RevA.docx',
        'FM-007':'FM-007_CoA_Review_Checklist_RevA.docx',
        'FM-008':'FM-008_Supplier_CoA_Requirements_Checklist_RevA.docx',
        'FM-027':'FM-027_QU_QS_Designation_Record_RevA.docx',
        'FM-030':'FM-030_Finished_Product_Spec_Sheet_RevA.docx',
        'FM-ALG':'FM-ALG_Allergen_Status_Record_RevA.docx',
      };

      // ── Download pane HTML (Effective phase only) ──
      const dlDocsHtml = isEffective ? (
        // DCO Report banner at top
        \`<div style="display:flex;align-items:center;gap:12px;padding:11px 14px;background:var(--n);border-radius:7px;margin-bottom:10px">
          <div style="flex:1">
            <div style="font-size:12px;font-weight:700;color:#fff">&#128196; DCO Completion Report — \${dcoId}</div>
            <div style="font-size:10px;color:rgba(255,255,255,.6);margin-top:2px">21 CFR Part 11 · Signatures · Training Compliance · Routing History</div>
          </div>
          <button id="dl-rpt-btn-\${dcoId}" style="font-size:11px;font-weight:700;padding:6px 14px;border-radius:5px;background:#fff;color:var(--n);border:none;cursor:pointer;white-space:nowrap">&#11015; Download PDF</button>
        </div>\` +
        // Per-document rows
        displayDocs.map((docId: string) => {
          const _rev  = docRevMap[docId] || 'Rev A';
          const _name = docNameMap[docId] || docId;
          const _fn   = _FM[docId] || (docId + '_RevA.docx');
          const _isF  = _FIDS.includes(docId);
          const _base = _isF ? 'Shared Documents/Official/QMS/Forms' : 'Shared Documents/Official/QMS/Documents';
          const _view = _SP + '/_layouts/15/WopiFrame.aspx?sourcedoc=' + encodeURIComponent('/sites/IMP9177/' + _base + '/' + _fn) + '&action=view';
          const _dl   = (_isF ? _OFF : _OFD) + '/' + encodeURIComponent(_fn);
          return \`<div style="display:flex;align-items:center;gap:10px;padding:8px 12px;border-bottom:1px solid var(--s1)">
            <div style="width:26px;height:26px;border-radius:5px;background:var(--g1);display:flex;align-items:center;justify-content:center;font-size:13px;flex-shrink:0">&#10003;</div>
            <div style="flex:1;min-width:0">
              <div style="font-size:11px;font-weight:700;font-family:var(--mono);color:var(--b)">\${docId}</div>
              <div style="font-size:11px;color:var(--s5);white-space:nowrap;overflow:hidden;text-overflow:ellipsis">\${_name}</div>
            </div>
            <div style="display:flex;align-items:center;gap:5px;flex-shrink:0">
              <span style="font-size:10px;font-family:var(--mono);font-weight:700;padding:1px 5px;border-radius:3px;background:var(--b1);color:var(--b)">\${_rev}</span>
              <span style="font-size:9px;padding:1px 5px;border-radius:3px;background:var(--g1);color:var(--g);font-weight:700">Effective Apr 24, 2026</span>
              <a href="\${_view}" target="_blank" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid var(--b);color:var(--b);background:var(--w);text-decoration:none">View &#8599;</a>
              <a href="\${_dl}" target="_blank" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid var(--s2);color:var(--s7);background:var(--s1);text-decoration:none">&#11015; DOCX</a>
            </div>
          </div>\`;
        }).join('')
      ) : '';

      const docRowsHtml = displayDocs.map((docId: string, idx2: number) => {`;

if (!src.includes(P1_OLD)) { console.error('PATCH 1 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P1_OLD, P1_NEW);
changed++;
console.log('✓ Patch 1 — phase flags + download rows inserted');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 2 — Replace hardcoded doc-review-section + sign-gate in body template
//           with conditional versions
// ─────────────────────────────────────────────────────────────────────────────
const P2_OLD = `        <div class="doc-review-section">
          <div class="doc-review-hdr">
            <div class="doc-review-title">Documents — open each before signing</div>
            <span class="doc-req-badge pend" id="dreq-badge-\${dcoId}">0 of \${displayDocs.length} opened</span>
          </div>
          \${docRowsHtml}
        </div>
        <div class="sign-gate-box" id="sgate-\${dcoId}">
          <div class="sign-gate-msg pend" id="sgate-msg-\${dcoId}">Open all documents before signing</div>
          <div class="sign-gate-check" id="sgate-docs-\${dcoId}">
            <div class="sgchk pend" id="sgate-chk-\${dcoId}">✕</div>
            <span id="sgate-txt-\${dcoId}">0 of \${displayDocs.length} documents opened</span>
          </div>
          <div class="sign-gate-check ok"><div class="sgchk ok">✓</div><span>M365 identity verified</span></div>
          <div class="sign-gate-check ok"><div class="sgchk ok">✓</div><span>Required approver on this DCO</span></div>
        </div>`;

const P2_NEW = `        \${isInReview ? \`
        <div class="doc-review-section">
          <div class="doc-review-hdr">
            <div class="doc-review-title">Documents — open each before signing</div>
            <span class="doc-req-badge pend" id="dreq-badge-\${dcoId}">0 of \${displayDocs.length} opened</span>
          </div>
          \${docRowsHtml}
        </div>
        <div class="sign-gate-box" id="sgate-\${dcoId}">
          <div class="sign-gate-msg pend" id="sgate-msg-\${dcoId}">Open all documents before signing</div>
          <div class="sign-gate-check" id="sgate-docs-\${dcoId}">
            <div class="sgchk pend" id="sgate-chk-\${dcoId}">✕</div>
            <span id="sgate-txt-\${dcoId}">0 of \${displayDocs.length} documents opened</span>
          </div>
          <div class="sign-gate-check ok"><div class="sgchk ok">✓</div><span>M365 identity verified</span></div>
          <div class="sign-gate-check ok"><div class="sgchk ok">✓</div><span>Required approver on this DCO</span></div>
        </div>\` : ''}`;

if (!src.includes(P2_OLD)) { console.error('PATCH 2 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P2_OLD, P2_NEW);
changed++;
console.log('✓ Patch 2 — review gate made conditional on isInReview');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 3 — Inject dlDocsHtml into the docs pane and wire report button
// ─────────────────────────────────────────────────────────────────────────────
const P3_OLD = `      const docsPaneEl = d.getElementById('mdco-pane-docs-' + dcoId);
      const docReviewEl = d.querySelector('.doc-review-section');
      const gateEl = d.getElementById('sgate-' + dcoId);
      if (docsPaneEl && docReviewEl) docsPaneEl.appendChild(docReviewEl);
      if (docsPaneEl && gateEl) docsPaneEl.appendChild(gateEl);`;

const P3_NEW = `      const docsPaneEl = d.getElementById('mdco-pane-docs-' + dcoId);
      if (isEffective && docsPaneEl) {
        const _dlWrap = d.createElement('div');
        _dlWrap.innerHTML = dlDocsHtml;
        docsPaneEl.appendChild(_dlWrap);
        // Wire DCO report download button
        const _rptBtn = d.getElementById('dl-rpt-btn-' + dcoId);
        if (_rptBtn) _rptBtn.addEventListener('click', () => (this as any)._generateDCOReport(dcoId));
      }
      const docReviewEl = d.querySelector('.doc-review-section');
      const gateEl = d.getElementById('sgate-' + dcoId);
      if (docsPaneEl && docReviewEl) docsPaneEl.appendChild(docReviewEl);
      if (docsPaneEl && gateEl) docsPaneEl.appendChild(gateEl);`;

if (!src.includes(P3_OLD)) { console.error('PATCH 3 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P3_OLD, P3_NEW);
changed++;
console.log('✓ Patch 3 — download pane wired into docs tab');

// ─────────────────────────────────────────────────────────────────────────────
// Write
// ─────────────────────────────────────────────────────────────────────────────
fs.writeFileSync(filePath, src, 'utf8');
console.log(`\n✅ All ${changed}/3 patches applied. Ready to build.`);
