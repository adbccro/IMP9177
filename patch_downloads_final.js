// patch_downloads_final.js
// Fixes all download issues in the DCO effective document pane:
//   1. DOCX download  — use _layouts/15/download.aspx (authenticated)
//   2. PDF download   — fetch file item via SP REST, then call Graph content?format=pdf
//   3. DCO Report PDF — fix timing with proper jsPDF load wait + fallback load trigger
// Run from SPFx root: node patch_downloads_final.js

const fs = require('fs');
const filePath = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
console.log('Reading:', filePath);
let src = fs.readFileSync(filePath, 'utf8');
let changed = 0;

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 1 — Fix DOCX download URL and PDF button in download rows
//           Replace the entire _view/_dl/_pdfUrl block with correct URLs
// ─────────────────────────────────────────────────────────────────────────────
const P1_OLD = `          const _view   = _SP + '/_layouts/15/Doc.aspx?sourcedoc=' + encodeURIComponent('/sites/IMP9177/' + _base + '/' + _fn) + '&action=view';
          const _dl    = (_isF ? _OFF : _OFD) + '/' + encodeURIComponent(_fn);
          // PDF conversion via SharePoint native endpoint (M365 converts DOCX→PDF server-side)
          const _pdfUrl = _SP + '/_layouts/15/Doc.aspx?sourcedoc=' + encodeURIComponent('/sites/IMP9177/' + _base + '/' + _fn) + '&action=download&DefaultItemOpen=1&format=pdf';`;

const P1_NEW = `          const _view   = _SP + '/_layouts/15/Doc.aspx?sourcedoc=' + encodeURIComponent('/sites/IMP9177/' + _base + '/' + _fn) + '&action=view';
          // Authenticated DOCX download via download.aspx
          const _dl    = _SP + '/_layouts/15/download.aspx?SourceUrl=' + encodeURIComponent('/sites/IMP9177/' + _base + '/' + _fn);
          // PDF: use Graph API endpoint — fetched via SPFx REST in click handler
          const _pdfPath = '/sites/IMP9177/' + _base + '/' + _fn;`;

if (!src.includes(P1_OLD)) { console.error('PATCH 1 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P1_OLD, P1_NEW);
changed++;
console.log('✓ Patch 1 — DOCX download URL fixed to download.aspx');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 2 — Fix PDF button: replace data-pdf-url with data-pdf-path
//           The click handler will use SPFx REST to get item ID then Graph PDF
// ─────────────────────────────────────────────────────────────────────────────
const P2_OLD = `              <button data-pdf-url="\${_pdfUrl}" data-doc-id="\${docId}" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid #dc2626;color:#dc2626;background:#fef2f2;cursor:pointer">&#11015; PDF</button>`;

const P2_NEW = `              <button data-pdf-path="\${_pdfPath}" data-doc-id="\${docId}" style="font-size:10px;font-weight:600;padding:3px 8px;border-radius:4px;border:1px solid #dc2626;color:#dc2626;background:#fef2f2;cursor:pointer;white-space:nowrap">&#11015; PDF</button>`;

if (!src.includes(P2_OLD)) { console.error('PATCH 2 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P2_OLD, P2_NEW);
changed++;
console.log('✓ Patch 2 — PDF button attribute updated');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 3 — Fix PDF click handler: use Graph API with SPFx token
// ─────────────────────────────────────────────────────────────────────────────
const P3_OLD = `        // Wire PDF download buttons inside _dlWrap
        _dlWrap.querySelectorAll('[data-pdf-url]').forEach((btn: Element) => {
          btn.addEventListener('click', (e: Event) => {
            e.preventDefault();
            const pdfUrl = (btn as HTMLElement).getAttribute('data-pdf-url') || '';
            const docId  = (btn as HTMLElement).getAttribute('data-doc-id') || '';
            if (pdfUrl) w.open(pdfUrl, '_blank');
            if (w.qpToast) w.qpToast('Opening PDF: ' + docId);
          });
        });`;

const P3_NEW = `        // Wire PDF download buttons — fetch via Graph API with SPFx auth token
        _dlWrap.querySelectorAll('[data-pdf-path]').forEach((btn: Element) => {
          btn.addEventListener('click', async (e: Event) => {
            e.preventDefault();
            const pdfPath = (btn as HTMLElement).getAttribute('data-pdf-path') || '';
            const docId   = (btn as HTMLElement).getAttribute('data-doc-id') || '';
            if (!pdfPath) return;
            if (w.qpToast) w.qpToast('Preparing PDF for ' + docId + '...');
            (btn as HTMLElement).textContent = '⏳ PDF';
            try {
              // Step 1: get file item ID via SharePoint REST
              const base = this.context.pageContext.web.absoluteUrl;
              const metaUrl = base + "/_api/web/GetFileByServerRelativeUrl('" + pdfPath.replace(/'/g, "''") + "')?$select=UniqueId,ListItemAllFields/Id&$expand=ListItemAllFields";
              const metaResp = await this.context.spHttpClient.get(metaUrl, SPHttpClient.configurations.v1);
              const meta = await metaResp.json();
              const uniqueId = meta?.UniqueId;
              if (!uniqueId) throw new Error('File not found: ' + pdfPath);
              // Step 2: build Graph PDF URL using site-relative approach
              // Graph: /sites/{hostname},{siteId},{webId}/drive/items/{uniqueId}/content?format=pdf
              const siteId = this.context.pageContext.site.id.toString();
              const webId  = this.context.pageContext.web.id.toString();
              const hostname = window.location.hostname;
              const graphUrl = 'https://graph.microsoft.com/v1.0/sites/' + hostname + ',' + siteId + ',' + webId + '/drive/items/' + uniqueId + '/content?format=pdf';
              // Step 3: fetch with auth token via aadHttpClientFactory
              const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
              const token = await tokenProvider.getToken('https://graph.microsoft.com');
              const pdfResp = await fetch(graphUrl, { headers: { 'Authorization': 'Bearer ' + token } });
              if (!pdfResp.ok) throw new Error('Graph PDF failed: ' + pdfResp.status);
              const blob = await pdfResp.blob();
              const url = URL.createObjectURL(blob);
              const a = document.createElement('a');
              a.href = url;
              a.download = docId + '_RevA.pdf';
              a.click();
              URL.revokeObjectURL(url);
              if (w.qpToast) w.qpToast('PDF downloaded: ' + docId);
              (btn as HTMLElement).innerHTML = '&#11015; PDF';
            } catch (err) {
              console.error('PDF download failed:', err);
              // Fallback: open Word Online print view (user can print to PDF)
              const viewUrl = 'https://adbccro.sharepoint.com/sites/IMP9177/_layouts/15/Doc.aspx?sourcedoc=' + encodeURIComponent(pdfPath) + '&action=view';
              w.open(viewUrl, '_blank');
              if (w.qpToast) w.qpToast('Opening in Word Online — use File > Print > Save as PDF');
              (btn as HTMLElement).innerHTML = '&#11015; PDF';
            }
          });
        });`;

if (!src.includes(P3_OLD)) { console.error('PATCH 3 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P3_OLD, P3_NEW);
changed++;
console.log('✓ Patch 3 — PDF handler uses Graph API with SPFx auth token');

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 4 — Fix DCO Report button: wait for jsPDF to load before firing
// ─────────────────────────────────────────────────────────────────────────────
const P4_OLD = `  private _generateDCOReport(dcoId:string):void{const w=this._iframe?.contentWindow as any;if(!w||!w.jspdf){if(w?.qpToast)w.qpToast('PDF library loading — try again in a moment');return;}`;

const P4_NEW = `  private _generateDCOReport(dcoId:string):void{const w=this._iframe?.contentWindow as any;if(!w)return;
    // If jsPDF not loaded yet, inject it dynamically and retry
    if(!w.jspdf){
      const d=this._iframe?.contentDocument;
      if(!d){return;}
      if(w.qpToast)w.qpToast('Loading PDF engine...');
      // Check if script tag already exists (may still be loading)
      const existing = d.querySelector('script[src*="jspdf"]');
      if(!existing){
        const s=d.createElement('script');
        s.src='https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js';
        s.onload=()=>{setTimeout(()=>this._generateDCOReport(dcoId),200);};
        d.head.appendChild(s);
      } else {
        // Script exists, wait for it to finish loading
        setTimeout(()=>this._generateDCOReport(dcoId),500);
      }
      return;
    }`;

if (!src.includes(P4_OLD)) { console.error('PATCH 4 FAILED — anchor not found'); process.exit(1); }
src = src.replace(P4_OLD, P4_NEW);
changed++;
console.log('✓ Patch 4 — DCO Report button auto-loads jsPDF if not ready');

// ─────────────────────────────────────────────────────────────────────────────
// Write
// ─────────────────────────────────────────────────────────────────────────────
fs.writeFileSync(filePath, src, 'utf8');
console.log(`\n✅ All ${changed}/4 patches applied. Ready to build.`);
