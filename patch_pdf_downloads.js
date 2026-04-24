// patch_pdf_downloads.js
// 1. Fix DCO Report button — query inside _dlWrap directly (not getElementById)
// 2. Add PDF download button per document using SharePoint's native ?download=1&format=pdf
// Run from SPFx root: node patch_pdf_downloads.js

const fs = require('fs');
const filePath = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
console.log('Reading:', filePath);
let src = fs.readFileSync(filePath, 'utf8');
let changed = 0;

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 1 — Fix report button wiring: query inside _dlWrap, not getElementById
// ─────────────────────────────────────────────────────────────────────────────
const P1_OLD = `      if (isEffective && docsPaneEl) {
        const _dlWrap = d.createElement('div');
        _dlWrap.innerHTML = dlDocsHtml;
        docsPaneEl.appendChild(_dlWrap);
        // Wire DCO report download button
        const _rptBtn = d.getElementById('dl-rpt-btn-' + dcoId);
        if (_rptBtn) _rptBtn.addEventListener('click', () => (this as any)._generateDCOReport(dcoId));
      }`;

const P1_NEW = `      if (isEffective && docsPaneEl) {
        const _dlWrap = d.createElement('div');
        _dlWrap.innerHTML = dlDocsHtml;
        docsPaneEl.appendChild(_dlWrap);
        // Wire DCO report button — query inside _dlWrap directly (tab may be hidden)
        const _rptBtn = _dlWrap.querySelector('#dl-rpt-btn-' + dcoId) as HTMLElement;
        if (_rptBtn) _rptBtn.addEventListener('click', () => { (this as any)._generateDCOReport(dcoId); });
        // Wire PDF download buttons inside _dlWrap
        _dlWrap.querySelectorAll('[data-pdf-url]').forEach((btn: Element) => {
          btn.addEventListener('click', (e: Event) => {
            e.preventDefault();
            const pdfUrl = (btn as HTMLElement).getAttribute('data-pdf-url') || '';
            const docId  = (btn as HTMLElement).getAttribute('data-doc-id') || '';
            if (pdfUrl) w.open(pdfUrl, '_blank');
            if (w.qpToast) w.qpToast('Opening PDF: ' + docId);
          });
        });
      }`;

if (!src.includes(P1_OLD)) { console.error('PATCH 1 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P1_OLD, P1_NEW);
changed++;
console.log('✓ Patch 1 — report button wiring fixed');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 2 — Add ⬇ PDF button to each document row
//           Uses SharePoint's native PDF conversion:
//           /_layouts/15/Doc.aspx?sourcedoc=<path>&action=download&DefaultItemOpen=1
//           Combined with the MS Graph format=pdf endpoint as fallback label
// ─────────────────────────────────────────────────────────────────────────────
const P2_OLD = `              <a href="\${_dl}" target="_blank" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid var(--s2);color:var(--s7);background:var(--s1);text-decoration:none">&#11015; DOCX</a>`;

const P2_NEW = `              <a href="\${_dl}" target="_blank" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid var(--s2);color:var(--s7);background:var(--s1);text-decoration:none">&#11015; DOCX</a>
              <button data-pdf-url="\${_pdfUrl}" data-doc-id="\${docId}" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid #dc2626;color:#dc2626;background:#fef2f2;cursor:pointer">&#11015; PDF</button>`;

// Also need to add _pdfUrl variable to the map function
const P2_VAR_OLD = `          const _view = _SP + '/_layouts/15/Doc.aspx?sourcedoc=' + encodeURIComponent('/sites/IMP9177/' + _base + '/' + _fn) + '&action=view';
          const _dl   = (_isF ? _OFF : _OFD) + '/' + encodeURIComponent(_fn);`;

const P2_VAR_NEW = `          const _view   = _SP + '/_layouts/15/Doc.aspx?sourcedoc=' + encodeURIComponent('/sites/IMP9177/' + _base + '/' + _fn) + '&action=view';
          const _dl    = (_isF ? _OFF : _OFD) + '/' + encodeURIComponent(_fn);
          // PDF conversion via SharePoint native endpoint (M365 converts DOCX→PDF server-side)
          const _pdfUrl = _SP + '/_layouts/15/Doc.aspx?sourcedoc=' + encodeURIComponent('/sites/IMP9177/' + _base + '/' + _fn) + '&action=download&DefaultItemOpen=1&format=pdf';`;

if (!src.includes(P2_VAR_OLD)) { console.error('PATCH 2a FAILED — var anchor not found'); process.exit(1); }
src = src.replace(P2_VAR_OLD, P2_VAR_NEW);

if (!src.includes(P2_OLD)) { console.error('PATCH 2b FAILED — button anchor not found'); process.exit(1); }
src = src.replace(P2_OLD, P2_NEW);
changed++;
console.log('✓ Patch 2 — PDF download button added per document');

// ─────────────────────────────────────────────────────────────────────────────
// Write
// ─────────────────────────────────────────────────────────────────────────────
fs.writeFileSync(filePath, src, 'utf8');
console.log(`\n✅ All ${changed}/2 patches applied. Ready to build.`);
