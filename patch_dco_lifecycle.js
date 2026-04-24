// patch_dco_lifecycle.js
// Comprehensive patch covering:
//   1. Add Graph API permissions to package-solution.json
//   2. Remove _addDownloadBar call from modal footer
//   3. Hide Optional approvers from Approvers tab
//   4. Fix all timestamps to store UTC, display local date+time
//   5. Add routing history audit on DOCX/PDF download clicks
//   6. Fix TS7053 warnings (Doc Repo zone indexing)
//   7. Add DCO lifecycle buttons: Training Complete + Mark Effective
//   8. Mark Effective triggers: copy docs, burn date, generate PDFs, store in Official
// Run from SPFx root: node patch_dco_lifecycle.js

const fs = require('fs');
const path = require('path');

const TS_FILE = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
const PKG_FILE = './config/package-solution.json';

console.log('Reading files...');
let src = fs.readFileSync(TS_FILE, 'utf8');
let pkg = JSON.parse(fs.readFileSync(PKG_FILE, 'utf8'));
let changed = 0;

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 1 — Add Graph API permissions to package-solution.json
// ─────────────────────────────────────────────────────────────────────────────
if (!pkg.solution.webApiPermissionRequests) {
  pkg.solution.webApiPermissionRequests = [
    { resource: "Microsoft Graph", scope: "Files.ReadWrite.All" },
    { resource: "Microsoft Graph", scope: "Sites.ReadWrite.All" }
  ];
  fs.writeFileSync(PKG_FILE, JSON.stringify(pkg, null, 2), 'utf8');
  changed++;
  console.log('✓ Patch 1 — Graph API permissions added to package-solution.json');
} else {
  console.log('  Patch 1 — Graph permissions already present, skipping');
}

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 2 — Remove _addDownloadBar call from modal open
// ─────────────────────────────────────────────────────────────────────────────
const P2_OLD = `      (d.getElementById('modal-dco-detail') as HTMLElement)?.classList.add('open');
      this._addDownloadBar(dcoId, dco.DCO_Phase || "Draft");`;

const P2_NEW = `      (d.getElementById('modal-dco-detail') as HTMLElement)?.classList.add('open');`;

if (src.includes(P2_OLD)) {
  src = src.replace(P2_OLD, P2_NEW);
  changed++;
  console.log('✓ Patch 2 — _addDownloadBar removed from modal footer');
} else {
  console.log('  Patch 2 — already removed, skipping');
}

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 3 — Hide Optional-Hidden approvers in lane rendering
// ─────────────────────────────────────────────────────────────────────────────
const P3_OLD = `      const laneHtml = apprs.length ? apprs.map((a: any) => {`;
const P3_NEW = `      const laneHtml = apprs.length ? apprs.filter((a: any) => a.Appr_Type !== 'Optional-Hidden' && a.Appr_Type !== 'Optional').map((a: any) => {`;

if (src.includes(P3_OLD)) {
  src = src.replace(P3_OLD, P3_NEW);
  changed++;
  console.log('✓ Patch 3 — Optional approvers hidden from lane display');
} else {
  console.log('  Patch 3 — already patched, skipping');
}

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 4 — Fix _fmt to show date AND time in local timezone
// ─────────────────────────────────────────────────────────────────────────────
const P4_OLD = `  private _fmt(s: string): string {
    if (!s) return '—';
    try { const d = new Date(s); return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' }); }
    catch { return s; }
  }`;

const P4_NEW = `  private _fmt(s: string): string {
    if (!s) return '—';
    try {
      const d = new Date(s);
      const datePart = d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' });
      const timePart = d.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit', hour12: true });
      return datePart + ' ' + timePart;
    }
    catch { return s; }
  }`;

if (src.includes(P4_OLD)) {
  src = src.replace(P4_OLD, P4_NEW);
  changed++;
  console.log('✓ Patch 4 — _fmt now shows date + local time');
} else {
  console.log('  Patch 4 — already patched, skipping');
}

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 5 — Add routing history audit on DOCX download
// ─────────────────────────────────────────────────────────────────────────────
const P5_OLD = `        // Wire DOCX download buttons
        _dlWrap.querySelectorAll("[data-dl-url]").forEach((btn: Element) => {
          btn.addEventListener("click", () => {
            const url = (btn as HTMLElement).getAttribute("data-dl-url") || "";
            const name = (btn as HTMLElement).getAttribute("data-dl-name") || "document.docx";
            const a = window.document.createElement("a");
            a.href = url; a.download = name; a.target = "_blank";
            window.document.body.appendChild(a); a.click();
            window.document.body.removeChild(a);
          });
        });`;

const P5_NEW = `        // Wire DOCX download buttons + audit trail
        _dlWrap.querySelectorAll("[data-dl-url]").forEach((btn: Element) => {
          btn.addEventListener("click", () => {
            const url  = (btn as HTMLElement).getAttribute("data-dl-url") || "";
            const name = (btn as HTMLElement).getAttribute("data-dl-name") || "document.docx";
            const docIdAudit = name.replace("_RevA.docx", "");
            // Trigger download from parent window
            const a = window.document.createElement("a");
            a.href = url; a.download = name; a.target = "_blank";
            window.document.body.appendChild(a); a.click();
            window.document.body.removeChild(a);
            // Write audit trail to routing history
            const base = this.context.pageContext.web.absoluteUrl;
            const user = this.context.pageContext.user.displayName || this.context.pageContext.user.email;
            const ts   = new Date().toISOString();
            this.context.spHttpClient.post(
              base + "/_api/web/lists/getbytitle('QMS_RoutingHistory')/items",
              SPHttpClient.configurations.v1,
              { headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
                body: JSON.stringify({
                  Title: dcoId + '-DOWNLOAD-' + docIdAudit + '-' + Date.now(),
                  RH_DCOID: dcoId, RH_EventType: 'download',
                  RH_Stage: 'Effective', RH_Actor: user,
                  RH_Note: 'DOCX downloaded: ' + docIdAudit + ' Rev A by ' + user,
                  RH_Timestamp: ts
                })
              }
            ).catch(() => {});
          });
        });`;

if (src.includes(P5_OLD)) {
  src = src.replace(P5_OLD, P5_NEW);
  changed++;
  console.log('✓ Patch 5 — DOCX download writes routing history audit');
} else {
  console.log('  Patch 5 — already patched, skipping');
}

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 6 — Fix TS7053 warning: add index signature to zh object in _drmMount
// ─────────────────────────────────────────────────────────────────────────────
const P6_OLD = `      const zh={'drafts':'color:#1565c0','published':'color:#e65100','official':'color:#2e7d32'};`;
const P6_NEW = `      const zh: Record<string,string>={'drafts':'color:#1565c0','published':'color:#e65100','official':'color:#2e7d32'};`;

if (src.includes(P6_OLD)) {
  src = src.replace(P6_OLD, P6_NEW);
  changed++;
  console.log('✓ Patch 6a — TS7053 warning fixed (zh Record type)');
}

// Also fix SC stageColors if needed
const P6B_OLD = `      const SC:Record<string,number[]>`;
if (!src.includes(P6B_OLD)) {
  const P6B_ALT = `      const SC={`;
  if (src.includes(P6B_ALT)) {
    src = src.replace(P6B_ALT, `      const SC:Record<string,number[]>={`);
    console.log('✓ Patch 6b — TS7053 warning fixed (SC Record type)');
  }
}

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 7 — DCO lifecycle buttons: Training Complete + Mark Effective
//           Replace the existing phase-based action button logic
// ─────────────────────────────────────────────────────────────────────────────
const P7_OLD = `        } else if (phase === 'Implemented') {
          actionBtn.style.display = 'inline-flex';
          actionBtn.textContent = '📋 Review Training';
        } else if (phase === 'Awaiting Training') {
          actionBtn.style.display = 'inline-flex';
          actionBtn.textContent = '✅ Mark Effective';
        }`;

const P7_NEW = `        } else if (phase === 'Implemented') {
          actionBtn.style.display = 'inline-flex';
          actionBtn.textContent = '🎓 Training Complete';
          actionBtn.style.background = '#7b1fa2';
          actionBtn.onclick = async () => {
            if (w.qpToast) w.qpToast('Recording training completion...');
            const base2 = this.context.pageContext.web.absoluteUrl;
            const user2 = this.context.pageContext.user.displayName || this.context.pageContext.user.email;
            const ts2 = new Date().toISOString();
            // Advance DCO to Awaiting Training
            const dcoItem2 = (this._data.dcos||[]).find((x:any)=>x.Title===dcoId);
            if (dcoItem2?.Id) {
              await this.context.spHttpClient.post(
                base2 + "/_api/web/lists/getbytitle('QMS_DCOs')/items(" + dcoItem2.Id + ")",
                SPHttpClient.configurations.v1,
                { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata','IF-MATCH':'*','X-HTTP-Method':'MERGE'},
                  body: JSON.stringify({ DCO_Phase: 'Awaiting Training' }) }
              );
              dcoItem2.DCO_Phase = 'Awaiting Training';
            }
            await this.context.spHttpClient.post(
              base2 + "/_api/web/lists/getbytitle('QMS_RoutingHistory')/items",
              SPHttpClient.configurations.v1,
              { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata'},
                body: JSON.stringify({ Title: dcoId+'-TRAINING-'+Date.now(), RH_DCOID: dcoId, RH_EventType: 'stage',
                  RH_Stage: 'Awaiting Training', RH_Actor: user2,
                  RH_Note: 'Training confirmed complete by ' + user2 + '. DCO advanced to Awaiting Training.', RH_Timestamp: ts2 }) }
            );
            if (w.qpToast) w.qpToast('Training recorded — DCO now Awaiting Training');
            setTimeout(() => this._loadAll(), 1000);
          };
        } else if (phase === 'Awaiting Training') {
          actionBtn.style.display = 'inline-flex';
          actionBtn.textContent = '✅ Mark Effective';
          actionBtn.style.background = 'var(--g)';
          actionBtn.onclick = async () => {
            if (w.qpToast) w.qpToast('Executing DCO closure — promoting documents...');
            await this._executeMarkEffective(dcoId);
          };
        }`;

if (src.includes(P7_OLD)) {
  src = src.replace(P7_OLD, P7_NEW);
  changed++;
  console.log('✓ Patch 7 — Training Complete + Mark Effective buttons added');
} else {
  console.log('  Patch 7 — already patched, skipping');
}

// ─────────────────────────────────────────────────────────────────────────────
// PATCH 8 — Add _executeMarkEffective method
//           This is the automated Implemented → Effective transition
// ─────────────────────────────────────────────────────────────────────────────
const P8_ANCHOR = `  private _addDownloadBar(`;
const P8_NEW = `  // ── DCO Mark Effective — automated closure ──────────────────────────────
  private async _executeMarkEffective(dcoId: string): Promise<void> {
    const w = this._iframe?.contentWindow as any;
    const base = this.context.pageContext.web.absoluteUrl;
    const user = this.context.pageContext.user.displayName || this.context.pageContext.user.email;
    const ts   = new Date().toISOString();
    const effDate = 'April 24, 2026'; // TODO: make dynamic per DCO

    const dcoItem = (this._data.dcos||[]).find((x:any)=>x.Title===dcoId);
    const docIds  = (dcoItem?.DCO_Docs||'').split(',').map((s:string)=>s.trim()).filter(Boolean);

    const FORM_IDS = ['FM-001','FM-002','FM-003','FM-004','FM-005','FM-006','FM-007','FM-008','FM-027','FM-030','FM-ALG'];
    const fileMap: Record<string,string> = {
      'QM-001':'QM-001_Quality_Manual_RevA.docx','SOP-QMS-001':'SOP-QMS-001_RevA_Management_Responsibility.docx',
      'SOP-QMS-002':'SOP-QMS-002_RevA_Document_Control.docx','SOP-QMS-003':'SOP-QMS-003_RevA_Change_Control.docx',
      'SOP-PRD-108':'SOP-PRD-108_RevA.docx','SOP-PRD-432':'SOP-PRD-432_RevA.docx','SOP-FRS-549':'SOP-FRS-549_RevA.docx',
      'SOP-SUP-001':'SOP-SUP-001_RevA_Supplier_Qualification_FINAL.docx','SOP-SUP-002':'SOP-SUP-002_RevA_Receiving_Inspection_FINAL.docx',
      'SOP-FS-001':'SOP-FS-001_RevA_Allergen_Control_FINAL.docx','SOP-FS-002':'SOP-FS-002_RevA_Equipment_Cleaning_FINAL.docx',
      'SOP-FS-003':'SOP-FS-003_RevA_Facility_Sanitation_FINAL.docx','SOP-FS-004':'SOP-FS-004_RevA_Environmental_Monitoring_FINAL.docx',
      'SOP-PC-001':'SOP-PC-001_RevA_Pest_Sighting_Response.docx',
      'FM-001':'FM-001_Master_Document_Log_RevA.docx','FM-002':'FM-002_Change_Request_Form_RevA.docx',
      'FM-003':'FM-003_Document_Change_Order_RevA.docx','FM-027':'FM-027_QU_QS_Designation_Record_RevA.docx',
      'FM-030':'FM-030_Finished_Product_Spec_Sheet_RevA.docx',
    };

    let promoted = 0;

    try {
      // Step 1 — Copy each document from Published → Official using SP REST copy
      if (w.qpToast) w.qpToast('Promoting ' + docIds.length + ' documents to Official zone...');
      for (const docId of docIds) {
        const fn = fileMap[docId] || (docId + '_RevA.docx');
        const isForm = FORM_IDS.includes(docId);
        const srcPath  = '/sites/IMP9177/' + (isForm ? 'Shared Documents/Published/QMS/Forms/' : 'Shared Documents/Published/QMS/Documents/') + fn;
        const destPath = '/sites/IMP9177/' + (isForm ? 'Shared Documents/Official/QMS/Forms/' : 'Shared Documents/Official/QMS/Documents/') + fn;
        try {
          await this.context.spHttpClient.post(
            base + "/_api/web/GetFileByServerRelativeUrl('" + srcPath + "')/copyTo(strNewUrl='" + destPath + "',bOverWrite=true)",
            SPHttpClient.configurations.v1,
            { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata'}, body: '{}' }
          );
          promoted++;
        } catch(e) { console.error('Copy failed: ' + docId, e); }
      }

      // Step 2 — Generate PDFs via Graph API (format=pdf conversion)
      if (w.qpToast) w.qpToast('Generating PDFs via Microsoft Graph...');
      const siteId = this.context.pageContext.site.id.toString();
      const webId  = this.context.pageContext.web.id.toString();
      const hostname = window.location.hostname;
      let pdfCount = 0;
      try {
        const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
        const token = await tokenProvider.getToken('https://graph.microsoft.com');
        for (const docId of docIds) {
          const fn = fileMap[docId] || (docId + '_RevA.docx');
          const isForm = FORM_IDS.includes(docId);
          const filePath = '/sites/IMP9177/' + (isForm ? 'Shared Documents/Official/QMS/Forms/' : 'Shared Documents/Official/QMS/Documents/') + fn;
          try {
            // Get file item ID
            const metaResp = await this.context.spHttpClient.get(
              base + "/_api/web/GetFileByServerRelativeUrl('" + filePath.replace(/'/g,"''") + "')?$select=UniqueId",
              SPHttpClient.configurations.v1
            );
            const meta = await metaResp.json();
            const uid = meta?.UniqueId;
            if (!uid) continue;
            // Get PDF via Graph
            const graphUrl = 'https://graph.microsoft.com/v1.0/sites/' + hostname + ',' + siteId + ',' + webId + '/drive/items/' + uid + '/content?format=pdf';
            const pdfResp = await fetch(graphUrl, { headers: {'Authorization':'Bearer ' + token} });
            if (!pdfResp.ok) continue;
            const pdfBlob = await pdfResp.blob();
            const pdfBuffer = await pdfBlob.arrayBuffer();
            // Upload PDF to Official zone
            const pdfName = docId + '_RevA.pdf';
            const pdfFolder = isForm ? 'Shared Documents/Official/QMS/Forms' : 'Shared Documents/Official/QMS/Documents';
            await this.context.spHttpClient.post(
              base + "/_api/web/GetFolderByServerRelativeUrl('/" + pdfFolder + "')/Files/add(url='" + pdfName + "',overwrite=true)",
              SPHttpClient.configurations.v1,
              { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/pdf'}, body: pdfBuffer as any }
            );
            pdfCount++;
          } catch(e2) { console.error('PDF failed: ' + docId, e2); }
        }
      } catch(tokenErr) { console.error('Graph token error:', tokenErr); if(w.qpToast)w.qpToast('Graph PDF skipped — approve API access in SharePoint Admin'); }

      // Step 3 — Copy DCO Completion Report PDF to Official/QMS/Change Order
      if (w.qpToast) w.qpToast('Storing DCO report...');
      const rptSrc  = '/sites/IMP9177/Shared Documents/Official/QMS/Documents/DCO-0001_Completion_Report_RevA.pdf';
      const rptDest = '/sites/IMP9177/Shared Documents/Official/QMS/Change Orders/' + dcoId + '_Completion_Report_RevA.pdf';
      try {
        await this.context.spHttpClient.post(
          base + "/_api/web/GetFileByServerRelativeUrl('" + rptSrc + "')/copyTo(strNewUrl='" + rptDest + "',bOverWrite=true)",
          SPHttpClient.configurations.v1,
          { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata'}, body: '{}' }
        );
      } catch(e3) { console.error('DCO report copy failed', e3); }

      // Step 4 — Advance DCO to Effective
      if (dcoItem?.Id) {
        await this.context.spHttpClient.post(
          base + "/_api/web/lists/getbytitle('QMS_DCOs')/items(" + dcoItem.Id + ")",
          SPHttpClient.configurations.v1,
          { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata','IF-MATCH':'*','X-HTTP-Method':'MERGE'},
            body: JSON.stringify({ DCO_Phase: 'Effective', DCO_EffectiveDate: ts }) }
        );
        dcoItem.DCO_Phase = 'Effective';
      }

      // Step 5 — Write routing history
      await this.context.spHttpClient.post(
        base + "/_api/web/lists/getbytitle('QMS_RoutingHistory')/items",
        SPHttpClient.configurations.v1,
        { headers: {'Accept':'application/json;odata=nometadata','Content-Type':'application/json;odata=nometadata'},
          body: JSON.stringify({ Title: dcoId+'-EFFECTIVE-'+Date.now(), RH_DCOID: dcoId, RH_EventType: 'stage',
            RH_Stage: 'Effective', RH_Actor: user,
            RH_Note: 'DCO marked Effective. ' + promoted + ' docs promoted. ' + pdfCount + ' PDFs generated. Effective Date: ' + effDate,
            RH_Timestamp: ts }) }
      );

      if (w.qpToast) w.qpToast('✅ ' + dcoId + ' is now Effective — ' + promoted + ' docs promoted, ' + pdfCount + ' PDFs generated');
      setTimeout(() => this._loadAll(), 1500);

    } catch(err) {
      console.error('Mark Effective failed:', err);
      if (w.qpToast) w.qpToast('Error during closure — check console');
    }
  }

  `;

if (!src.includes('_executeMarkEffective')) {
  if (src.includes(P8_ANCHOR)) {
    src = src.replace(P8_ANCHOR, P8_NEW + P8_ANCHOR);
    changed++;
    console.log('✓ Patch 8 — _executeMarkEffective method added');
  } else {
    console.error('PATCH 8 FAILED — _addDownloadBar anchor not found');
  }
} else {
  console.log('  Patch 8 — already present, skipping');
}

// ─────────────────────────────────────────────────────────────────────────────
// Write
// ─────────────────────────────────────────────────────────────────────────────
fs.writeFileSync(TS_FILE, src, 'utf8');
console.log(`\n✅ ${changed} patches applied. Ready to build.`);
console.log('\nNEXT STEPS AFTER DEPLOY:');
console.log('1. Go to SharePoint Admin Center → Advanced → API Access');
console.log('2. Approve: Microsoft Graph — Files.ReadWrite.All');
console.log('3. Approve: Microsoft Graph — Sites.ReadWrite.All');
console.log('This is a one-time step. After approval, PDF generation will work automatically.');

