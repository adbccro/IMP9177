// patch_fix_downloads2.js
// Fix 1: DOCX — trigger download via parent window postMessage (bypasses iframe sandbox)
// Fix 2: DCO Report — expose _generateDCOReport on parent window so iframe can call it directly
// Run from SPFx root: node patch_fix_downloads2.js

const fs = require('fs');
const filePath = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
console.log('Reading:', filePath);
let src = fs.readFileSync(filePath, 'utf8');
let changed = 0;

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 1 — Expose _generateDCOReport on window so iframe can call it via parent
//           Add to _attachListeners after the existing reg() calls
// ─────────────────────────────────────────────────────────────────────────────
const P1_OLD = `    // Flush any queued onclick calls that fired before listeners were ready
    if (w._qpFlush) w._qpFlush();`;

const P1_NEW = `    // Flush any queued onclick calls that fired before listeners were ready
    if (w._qpFlush) w._qpFlush();

    // Expose download helpers on iframe window so buttons can call them directly
    w._qpDownloadDocx = (downloadUrl: string, fileName: string) => {
      // Open download.aspx in parent window (not iframe) to avoid sandbox blocking
      const a = window.document.createElement('a');
      a.href = downloadUrl;
      a.download = fileName;
      a.target = '_blank';
      window.document.body.appendChild(a);
      a.click();
      window.document.body.removeChild(a);
    };
    w._qpGenerateReport = (dcoId: string) => {
      this._generateDCOReport(dcoId);
    };`;

if (!src.includes(P1_OLD)) { console.error('PATCH 1 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P1_OLD, P1_NEW);
changed++;
console.log('✓ Patch 1 — _qpDownloadDocx and _qpGenerateReport exposed on iframe window');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 2 — Fix DOCX <a> tag to call _qpDownloadDocx instead of href navigation
// ─────────────────────────────────────────────────────────────────────────────
const P2_OLD = `              <a href="\${_dl}" target="_blank" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid var(--s2);color:var(--s7);background:var(--s1);text-decoration:none">&#11015; DOCX</a>`;

const P2_NEW = `              <button onclick="window._qpDownloadDocx('\${_dl}','\${docId}_RevA.docx')" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid var(--s2);color:var(--s7);background:var(--s1);cursor:pointer;white-space:nowrap">&#11015; DOCX</button>`;

if (!src.includes(P2_OLD)) { console.error('PATCH 2 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P2_OLD, P2_NEW);
changed++;
console.log('✓ Patch 2 — DOCX button calls _qpDownloadDocx via parent window');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 3 — Fix DCO Report button to call _qpGenerateReport via window
// ─────────────────────────────────────────────────────────────────────────────
const P3_OLD = `          <button id="dl-rpt-btn-\${dcoId}" style="font-size:11px;font-weight:700;padding:6px 14px;border-radius:5px;background:#fff;color:var(--n);border:none;cursor:pointer;white-space:nowrap">&#11015; Download PDF</button>`;

const P3_NEW = `          <button onclick="window._qpGenerateReport('\${dcoId}')" style="font-size:11px;font-weight:700;padding:6px 14px;border-radius:5px;background:#fff;color:var(--n);border:none;cursor:pointer;white-space:nowrap">&#11015; Download PDF</button>`;

if (!src.includes(P3_OLD)) { console.error('PATCH 3 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P3_OLD, P3_NEW);
changed++;
console.log('✓ Patch 3 — DCO Report button calls _qpGenerateReport via window');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 4 — Remove old rptBtn wiring (now handled by onclick above)
// ─────────────────────────────────────────────────────────────────────────────
const P4_OLD = `        // Wire DCO report button — query inside _dlWrap directly (tab may be hidden)
        const _rptBtn = _dlWrap.querySelector('#dl-rpt-btn-' + dcoId) as HTMLElement;
        if (_rptBtn) _rptBtn.addEventListener('click', () => { (this as any)._generateDCOReport(dcoId); });`;

const P4_NEW = `        // DCO report button uses onclick="window._qpGenerateReport()" — no wiring needed here`;

if (!src.includes(P4_OLD)) { console.error('PATCH 4 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P4_OLD, P4_NEW);
changed++;
console.log('✓ Patch 4 — removed stale rptBtn addEventListener');

// ─────────────────────────────────────────────────────────────────────────────
// Write
// ─────────────────────────────────────────────────────────────────────────────
fs.writeFileSync(filePath, src, 'utf8');
console.log(`\n✅ All ${changed}/4 patches applied. Ready to build.`);
