# patch_and_deploy.ps1
# Patches Imp9177DashboardWebPart.ts with the Document Repository dashboard,
# then builds and deploys to SharePoint.
# Run from: C:\Users\andre\ADBCCRO Dropbox\Andre Butler\IMP9177\imp9177-spfx

Set-StrictMode -Off
$ErrorActionPreference = "Stop"

$webPartFile = ".\src\webparts\imp9177Dashboard\Imp9177DashboardWebPart.ts"

Write-Host "Reading source file..." -ForegroundColor Cyan
$src = Get-Content $webPartFile -Raw

# -------------------------------------------------------------------------
# GUARD: don't patch twice
# -------------------------------------------------------------------------
if ($src -match "DOCREPO_BODY") {
  Write-Host "Already patched. Skipping patch step." -ForegroundColor Yellow
} else {

  Write-Host "Patching CSS block..." -ForegroundColor Cyan

  # -------------------------------------------------------------------------
  # 1. Inject Document Repository CSS at end of CSS const (before closing backtick)
  # -------------------------------------------------------------------------
  $drmCss = @'

/* ---- Document Repository Matrix ---- */
.drm-tb{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap;}
.drm-tb select,.drm-tb input{font-size:12px;padding:3px 8px;border:1px solid var(--s2);border-radius:4px;height:28px;background:var(--w);color:var(--s7);}
.drm-tb input{min-width:160px;}
.drm-live{font-size:10px;padding:2px 8px;border-radius:10px;background:#D1FAE5;color:#065F46;border:1px solid #A7F3D0;font-weight:600;}
.drm-sbar{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap;}
.drm-pill{font-size:11px;padding:2px 9px;border-radius:10px;border:1px solid var(--s2);background:var(--s0);color:var(--s5);}
.drm-pill strong{font-weight:700;color:var(--n);}
.drm-pill.warn{border-color:#FCA5A5;color:#B91C1C;background:#FEE2E2;}
.drm-flow{display:grid;grid-template-columns:100px 144px 1fr 52px 1fr 52px 1fr;align-items:center;margin-bottom:4px;}
.drm-fblank{grid-column:span 2;}
.drm-fzone{text-align:center;font-size:10px;color:var(--s5);line-height:1.4;}
.drm-fzone small{font-size:9px;opacity:.7;display:block;}
.drm-farr{display:flex;flex-direction:column;align-items:center;gap:2px;}
.drm-fline{width:100%;height:1px;background:var(--s2);position:relative;}
.drm-fline::after{content:'';position:absolute;right:-1px;top:-3px;border-left:4px solid var(--s2);border-top:3px solid transparent;border-bottom:3px solid transparent;}
.drm-flbl{font-size:9px;color:var(--s5);text-align:center;line-height:1.2;}
.drm-table{width:100%;border-collapse:collapse;table-layout:fixed;}
.drm-table col.drm-cid{width:108px;}
.drm-table col.drm-ctitle{width:150px;}
.drm-table col.drm-czone{width:calc((100% - 258px)/3);}
.drm-table thead th{font-size:11px;font-weight:700;padding:6px 9px;text-align:left;border-bottom:2px solid var(--s2);background:var(--s0);}
.drm-zh{text-align:center!important;font-size:12px;}
.drm-zh-d{color:#1E56A0;border-bottom-color:#1E56A0!important;}
.drm-zh-p{color:#92400E;border-bottom-color:#D97706!important;}
.drm-zh-o{color:#065F46;border-bottom-color:#059669!important;}
.drm-table tbody tr{border-bottom:1px solid var(--s1);cursor:pointer;}
.drm-table tbody tr:hover td{background:var(--b0);}
.drm-table tbody tr.drm-sel td{background:#DBEAFE;}
.drm-table tbody tr td{padding:0;vertical-align:top;}
.drm-grp td{padding:4px 9px 3px!important;font-size:10px;font-weight:700;color:var(--s5);background:var(--s0);border-top:1px solid var(--s2);cursor:default!important;text-transform:uppercase;letter-spacing:.4px;}
.drm-grp:hover td{background:var(--s0)!important;}
.drm-ic{padding:7px 9px;}
.drm-docid{font-size:11px;font-weight:700;font-family:monospace;color:#1E56A0;white-space:nowrap;}
.drm-badge{font-size:9px;padding:1px 4px;border-radius:2px;margin-top:2px;display:inline-block;font-weight:700;}
.drm-SOP{background:#DBEAFE;color:#1E56A0;}
.drm-FM{background:#FEF3C7;color:#92400E;}
.drm-QM{background:#D1FAE5;color:#065F46;}
.drm-FPS{background:#EDE9FE;color:#5B21B6;}
.drm-tc{padding:7px 9px;font-size:11px;color:var(--s5);line-height:1.4;vertical-align:middle!important;}
.drm-zc{padding:6px 8px;border-left:1px solid var(--s1);vertical-align:top;}
.drm-zc.drm-empty{background:repeating-linear-gradient(45deg,transparent,transparent 5px,var(--s1) 5px,var(--s1) 5.5px);opacity:.4;}
.drm-vc{display:flex;flex-direction:column;gap:3px;}
.drm-vb{display:flex;align-items:center;gap:4px;flex-wrap:wrap;}
.drm-rev{font-size:10px;font-family:monospace;font-weight:700;padding:1px 5px;border-radius:2px;}
.drm-rA{background:#DBEAFE;color:#1E56A0;}
.drm-rB{background:#D1FAE5;color:#065F46;}
.drm-rC{background:#FEF3C7;color:#92400E;}
.drm-ver{font-size:10px;font-family:monospace;padding:1px 4px;border-radius:2px;background:var(--s1);color:var(--s5);}
.drm-ver.major{background:#D1FAE5;color:#065F46;font-weight:700;}
.drm-dco{font-size:9px;font-family:monospace;padding:1px 4px;border-radius:2px;white-space:nowrap;}
.drm-blocked{background:#FEE2E2;color:#B91C1C;}
.drm-open{background:#FEF3C7;color:#92400E;}
.drm-signed{background:#D1FAE5;color:#065F46;}
.drm-dco-na{background:var(--s1);color:var(--s5);}
.drm-mod{font-size:9px;color:var(--s5);}
.drm-path{font-size:9px;color:var(--s5);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:130px;}
.drm-warn{font-size:9px;color:#B91C1C;}
.drm-detail{margin-top:10px;border:1px solid var(--s2);border-radius:8px;background:var(--w);overflow:hidden;}
.drm-dhead{padding:8px 14px;background:var(--s0);border-bottom:1px solid var(--s2);display:flex;align-items:center;gap:10px;}
.drm-dhead-id{font-family:monospace;font-size:13px;font-weight:700;color:var(--n);}
.drm-dhead-title{font-size:12px;color:var(--s5);flex:1;}
.drm-dhead-close{cursor:pointer;color:var(--s5);font-size:16px;padding:0 4px;border:none;background:transparent;font-weight:700;}
.drm-dbody{padding:10px 14px;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;}
.drm-dzone{border:1px solid var(--s2);border-radius:6px;padding:10px;}
.drm-dzone-name{font-size:11px;font-weight:700;margin-bottom:7px;padding-bottom:5px;border-bottom:1px solid var(--s2);}
.drm-dzone-name.drafts{color:#1E56A0;}.drm-dzone-name.published{color:#92400E;}.drm-dzone-name.official{color:#065F46;}
.drm-dfield{font-size:11px;margin-bottom:4px;display:flex;gap:5px;align-items:flex-start;}
.drm-dlbl{color:var(--s5);min-width:58px;flex-shrink:0;font-size:10px;padding-top:1px;}
.drm-dval{color:var(--s7);word-break:break-all;font-size:10px;font-family:monospace;}
.drm-absent{font-size:11px;color:var(--s5);font-style:italic;}
.drm-open-sp{display:inline-block;margin-top:5px;font-size:10px;color:#1E56A0;text-decoration:none;padding:2px 7px;border:1px solid var(--s2);border-radius:3px;background:var(--w);}
.drm-open-sp:hover{background:var(--b0);}
.drm-spinner{display:flex;align-items:center;gap:10px;padding:24px 0;color:var(--s5);font-size:13px;}
.drm-spin{width:20px;height:20px;border:2px solid var(--s2);border-top-color:var(--b);border-radius:50%;animation:spin .8s linear infinite;flex-shrink:0;}
@keyframes spin{to{transform:rotate(360deg);}}
'@

  # Insert DRM CSS before the closing backtick of the CSS const
  $src = $src -replace '(@media\(max-width:768px\)\{[^`]+\})\s*`', "`$1$drmCss``"

  # -------------------------------------------------------------------------
  # 2. Add DOCREPO_BODY const after PM_BODY const
  # -------------------------------------------------------------------------
  $docrepoBody = @'

const DOCREPO_BODY = `
<div class="panel full">
  <div class="ph">
    <div class="pt">&#128218; Document Repository Matrix</div>
    <span id="drm-live" class="drm-live">Loading...</span>
  </div>
  <div style="padding:16px 18px;">
    <div class="drm-tb">
      <select id="drm-tf"><option value="all">All types</option><option value="SOP">SOP</option><option value="FM">FM</option><option value="QM">QM</option><option value="FPS">FPS</option></select>
      <select id="drm-df"><option value="all">All DCOs</option><option value="DCO-0001">DCO-0001</option><option value="DCO-0002">DCO-0002</option><option value="none">No DCO</option></select>
      <select id="drm-zf"><option value="all">All zones</option><option value="drafts">Has Drafts</option><option value="published">Has Published</option><option value="drafts-only">Drafts only</option><option value="both">Both zones</option></select>
      <input id="drm-qf" type="text" placeholder="Search ID or title...">
    </div>
    <div class="drm-sbar" id="drm-sbar"></div>
    <div class="drm-flow">
      <div class="drm-fblank"></div>
      <div class="drm-fzone">Drafts<small>QMS/.../Drafts</small></div>
      <div class="drm-farr"><div class="drm-fline"></div><div class="drm-flbl">major check-in</div></div>
      <div class="drm-fzone">Published<small>Published/QMS/...</small></div>
      <div class="drm-farr"><div class="drm-fline"></div><div class="drm-flbl">DCO sign-off</div></div>
      <div class="drm-fzone">Official<small>Official/QMS (empty)</small></div>
    </div>
    <div id="drm-loading" class="drm-spinner"><div class="drm-spin"></div><span>Loading document matrix from SharePoint...</span></div>
    <div id="drm-table-wrap" style="display:none">
      <table class="drm-table">
        <colgroup><col class="drm-cid"><col class="drm-ctitle"><col class="drm-czone"><col class="drm-czone"><col class="drm-czone"></colgroup>
        <thead><tr>
          <th>Doc ID</th><th>Title</th>
          <th class="drm-zh drm-zh-d">Drafts</th>
          <th class="drm-zh drm-zh-p">Published</th>
          <th class="drm-zh drm-zh-o">Official</th>
        </tr></thead>
        <tbody id="drm-tbody"></tbody>
      </table>
    </div>
    <div id="drm-detail" style="display:none"></div>
  </div>
</div>`;

'@

  # Insert after the closing of PM_BODY
  $src = $src -replace '(const PM_BODY = `[\s\S]+?`;\s*\n)', "`$1$docrepoBody"

  # -------------------------------------------------------------------------
  # 3. Add DOCREPO branch in render() method
  # -------------------------------------------------------------------------
  $docrepoRenderBranch = @'
 } else if (dashboard === 'DOCREPO') {
      html = makeHtml('DOCS', 'Document Repository',
        'IMP9177 \u00b7 3H Pharmaceuticals LLC \u00b7 Live document matrix across Drafts / Published / Official',
        DOCREPO_BODY,
        'IMP9177 Document Repository \u00b7 ADB Consulting & CRO Inc. \u00b7 Live data from SharePoint file libraries');
'@

  # Insert before the final else block in render()
  $src = $src -replace '(\s*\} else \{\s*\n\s*html = makeHtml\(''PM'')', "$docrepoRenderBranch`$1"

  # -------------------------------------------------------------------------
  # 4. Add DOCREPO branch in _fetchAll
  # -------------------------------------------------------------------------
  $docrepoFetchBranch = @'

    if (dashboard === 'DOCREPO') {
      return {};
    }
'@
  $src = $src -replace '(private async _fetchAll\(dashboard: string\): Promise<Record<string, any\[\]>> \{)', "`$1$docrepoFetchBranch"

  # -------------------------------------------------------------------------
  # 5. Add DOCREPO branch in _renderAll / _renderFallback
  # -------------------------------------------------------------------------
  $docrepoRenderAll = @'

    if (dashboard === 'DOCREPO') {
      this._loadDocumentMatrix();
      return;
    }
'@
  $src = $src -replace '(private _renderAll\(dashboard: string, data: Record<string, any\[\]>\): void \{)', "`$1$docrepoRenderAll"

  $docrepoFallback = @'

    if (dashboard === 'DOCREPO') {
      this._loadDocumentMatrix();
      return;
    }
'@
  $src = $src -replace '(private _renderFallback\(dashboard: string\): void \{)', "`$1$docrepoFallback"

  # -------------------------------------------------------------------------
  # 6. Add PropertyPane option for DOCREPO
  # -------------------------------------------------------------------------
  $src = $src -replace "(key: 'PM', text: 'PM Tracker' \})", "`$1,`n                { key: 'DOCREPO', text: 'Document Repository' }"

  # -------------------------------------------------------------------------
  # 7. Inject _loadDocumentMatrix method before the closing brace of the class
  # The method fetches SP file metadata and renders the DRM table via contentDocument
  # -------------------------------------------------------------------------
  $drmMethod = @'

  // ---------------------------------------------------------------------------
  // Document Repository Matrix -- data fetch + render
  // ---------------------------------------------------------------------------
  private _TITLES: Record<string, string> = {
    'QM-001':'Quality Manual','SOP-QMS-001':'Management Responsibility','SOP-QMS-002':'Document Control',
    'SOP-QMS-003':'Change Control','SOP-SUP-001':'Supplier Qualification','SOP-SUP-002':'Receiving Inspection',
    'SOP-FS-001':'Allergen Control','SOP-FS-002':'Equipment Cleaning','SOP-FS-003':'Facility Sanitation',
    'SOP-FS-004':'Environmental Monitoring','SOP-PC-001':'Pest Sighting Response',
    'SOP-PRD-108':'Finished Product Release','SOP-PRD-432':'Finished Product Spec & Testing',
    'SOP-FRS-549':'Product Specification Sheet','SOP-RCL-321':'Recall Procedure',
    'FPS-001':'Lychee VD3 Gummy Spec','FM-001':'Master Document Log','FM-002':'Change Request Form',
    'FM-003':'Document Change Order Form','FM-004':'Approved Supplier List','FM-005':'Receiving Log',
    'FM-006':'Raw Material Spec Sheet','FM-007':'Material Hold Label','FM-027':'QU/QS Designation Record',
    'FM-030':'Finished Product Spec Sheet','FM-ALG':'Allergen Status Record'
  };
  private _GROUP_MAP: Record<string, string> = {
    'QM-':'Quality Manual','SOP-QMS-':'SOPs -- QMS Core','SOP-SUP-':'SOPs -- Supplier & Receiving',
    'SOP-FS-':'SOPs -- Food Safety','SOP-PC-':'SOPs -- Pest Control',
    'SOP-PRD-':'SOPs -- Production','SOP-FRS-':'SOPs -- Production','SOP-RCL-':'SOPs -- Production',
    'FPS-':'Finished Product Specs','FM-00':'Forms -- Change Control','FM-0':'Forms -- Supplier & Receiving',
    'FM-2':'Forms -- Quality Unit','FM-ALG':'Forms -- Allergen'
  };
  private _DCO_STATUS: Record<string, string> = { 'DCO-0001':'blocked','DCO-0002':'open' };

  private _docGroup(id: string): string {
    for (const prefix of Object.keys(this._GROUP_MAP)) {
      if (id.startsWith(prefix)) return this._GROUP_MAP[prefix];
    }
    return 'Other';
  }
  private _extractDocId(name: string): string {
    const m = name.match(/^([A-Z]{2,5}-[A-Z]{2,5}-\d{3}|[A-Z]{2,5}-\d{3}|FM-[A-Z]{2,5}|QM-\d{3}|FPS-\d{3})/i);
    return m ? m[1].toUpperCase() : name.replace(/\.[^.]+$/, '').toUpperCase();
  }
  private _extractRev(name: string): string {
    const m = name.match(/_Rev([A-Z])/i); return m ? m[1].toUpperCase() : 'A';
  }
  private _inferDco(id: string): string | null {
    const dco1 = ['QM-001','SOP-QMS-001','SOP-QMS-002','SOP-QMS-003','SOP-PRD-108','SOP-PRD-432','SOP-FRS-549','FM-001','FM-002','FM-003','FM-027','FM-030'];
    if (dco1.includes(id)) return 'DCO-0001';
    if (id.startsWith('SOP-SUP-')||id.startsWith('SOP-FS-')||id.startsWith('SOP-PC-')||id.startsWith('FM-0')||id==='FM-ALG') return 'DCO-0002';
    return null;
  }

  private _loadDocumentMatrix(): void {
    const base = this.context.pageContext.web.absoluteUrl;
    const folders: Record<string, string[]> = {
      drafts: ['Shared%20Documents/QMS/Documents/Drafts','Shared%20Documents/QMS/Forms/Drafts'],
      published: ['Shared%20Documents/Published/QMS/Documents','Shared%20Documents/Published/QMS/Forms','Shared%20Documents/Published/QMS/Quality%20Manual'],
      official: ['Shared%20Documents/Official/QMS/Documents','Shared%20Documents/Official/QMS/Forms']
    };

    const fetchFolder = (folder: string): Promise<any[]> => {
      const url = `${base}/_api/web/GetFolderByServerRelativeUrl('${decodeURIComponent(folder)}')/Files?$select=Name,ServerRelativeUrl,TimeLastModified,CheckOutType,CheckedOutByUser/Title,MajorVersion,MinorVersion&$expand=CheckedOutByUser`;
      return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
        .then((r: SPHttpClientResponse) => r.ok ? r.json() : { value: [] })
        .then((d: any) => (d.value || []).filter((f: any) => f.Name.toLowerCase().endsWith('.docx') && !f.ServerRelativeUrl.toLowerCase().includes('/archive/')))
        .catch(() => []);
    };

    Promise.all([
      Promise.all(folders.drafts.map(fetchFolder)).then(a => ([] as any[]).concat(...a)),
      Promise.all(folders.published.map(fetchFolder)).then(a => ([] as any[]).concat(...a)),
      Promise.all(folders.official.map(fetchFolder)).then(a => ([] as any[]).concat(...a))
    ]).then(([draftFiles, pubFiles, offFiles]) => {
      const docMap = new Map<string, any>();
      const SP_SITE = 'https://adbccro.sharepoint.com/sites/IMP9177';

      const addFile = (f: any, zone: string): void => {
        const id = this._extractDocId(f.Name);
        if (!docMap.has(id)) {
          docMap.set(id, { id, title: this._TITLES[id] || id, type: id.startsWith('QM-') ? 'QM' : id.startsWith('SOP-') ? 'SOP' : id.startsWith('FM-') ? 'FM' : id.startsWith('FPS-') ? 'FPS' : 'DOC', group: this._docGroup(id), zones: { drafts: null, published: null, official: null } });
        }
        const doc = docMap.get(id);
        const major = f.MajorVersion || 1, minor = f.MinorVersion || 0;
        const ver = minor === 0 ? `${major}.0` : `${major}.${minor}`;
        const dco = this._inferDco(id);
        const dcoStatus = dco ? (this._DCO_STATUS[dco] || 'open') : 'n/a';
        const pathParts = (f.ServerRelativeUrl || '').split('/');
        const folderLabel = pathParts.slice(4, -1).join('/') || zone;
        doc.zones[zone] = {
          rev: this._extractRev(f.Name), ver, verType: minor === 0 ? 'major' : 'minor',
          dco, dcoStatus, checkedOut: f.CheckOutType > 0 ? (f.CheckedOutByUser && f.CheckedOutByUser.Title ? f.CheckedOutByUser.Title : 'Unknown') : null,
          file: f.Name, path: folderLabel, modified: (f.TimeLastModified || '').substring(0, 10),
          webUrl: `${SP_SITE}${f.ServerRelativeUrl}`
        };
      };

      draftFiles.forEach((f: any) => addFile(f, 'drafts'));
      pubFiles.forEach((f: any) => addFile(f, 'published'));
      offFiles.forEach((f: any) => addFile(f, 'official'));

      const GROUP_ORDER = ['Quality Manual','SOPs -- QMS Core','SOPs -- Supplier & Receiving','SOPs -- Food Safety','SOPs -- Pest Control','SOPs -- Production','Finished Product Specs','Forms -- Change Control','Forms -- Supplier & Receiving','Forms -- Quality Unit','Forms -- Allergen','Other'];
      const docs = [...docMap.values()].sort((a, b) => {
        const ga = GROUP_ORDER.indexOf(a.group), gb = GROUP_ORDER.indexOf(b.group);
        return ga !== gb ? ga - gb : a.id.localeCompare(b.id);
      });

      this._renderDrmTable(docs);
    }).catch(() => {
      this._renderDrmError();
    });
  }

  private _renderDrmTable(allDocs: any[]): void {
    const d = this.iframe && this.iframe.contentDocument;
    if (!d) return;

    const loading = d.getElementById('drm-loading');
    const tableWrap = d.getElementById('drm-table-wrap');
    const liveEl = d.getElementById('drm-live');
    if (loading) loading.style.display = 'none';
    if (tableWrap) tableWrap.style.display = '';
    if (liveEl) { liveEl.textContent = 'Live \u00b7 SP \u00b7 ' + new Date().toLocaleDateString('en-US', { month:'short', day:'numeric', year:'numeric' }); }

    let filtered = [...allDocs];
    let selected: string | null = null;

    const dcoClass = (ds: string): string => ({ blocked:'drm-blocked', open:'drm-open', signed:'drm-signed', 'n/a':'drm-dco-na' }[ds] || 'drm-dco-na');

    const zoneCell = (z: any): string => {
      if (!z) return '<td class="drm-zc drm-empty"></td>';
      const vi = z.verType === 'major' ? '&#9679;' : '&#9675;';
      const vc = z.verType === 'major' ? 'major' : '';
      const co = z.checkedOut ? `<div style="font-size:9px;color:#92400E;">&#9679; ${z.checkedOut}</div>` : `<div style="font-size:9px;color:#065F46;">&#9675; Checked in</div>`;
      return `<td class="drm-zc"><div class="drm-vc"><div class="drm-vb"><span class="drm-rev drm-r${z.rev}">Rev ${z.rev}</span><span class="drm-ver ${vc}">${vi} ${z.ver}</span>${z.dco ? `<span class="drm-dco ${dcoClass(z.dcoStatus)}">${z.dco}</span>` : '<span class="drm-dco drm-dco-na">--</span>'}</div>${co}<span class="drm-mod">${z.modified}</span><span class="drm-path" title="${z.path}">${z.path}</span></div></td>`;
    };

    const renderSummary = (): void => {
      const sbar = d.getElementById('drm-sbar');
      if (!sbar) return;
      const tot = filtered.length;
      const both = filtered.filter((x: any) => x.zones.drafts && x.zones.published).length;
      const donly = filtered.filter((x: any) => x.zones.drafts && !x.zones.published).length;
      const ponly = filtered.filter((x: any) => !x.zones.drafts && x.zones.published).length;
      const bl = filtered.filter((x: any) => ['drafts','published','official'].some((zk: string) => x.zones[zk] && x.zones[zk].dcoStatus === 'blocked')).length;
      const co = filtered.filter((x: any) => ['drafts','published','official'].some((zk: string) => x.zones[zk] && x.zones[zk].checkedOut)).length;
      sbar.innerHTML = `<span class="drm-pill"><strong>${tot}</strong> docs</span><span class="drm-pill"><strong>${both}</strong> in both zones</span><span class="drm-pill"><strong>${donly}</strong> drafts only</span><span class="drm-pill"><strong>${ponly}</strong> published only</span><span class="drm-pill"><strong>0</strong> in Official</span>${bl ? `<span class="drm-pill warn"><strong>${bl}</strong> DCO blocked</span>` : ''}${co ? `<span class="drm-pill warn"><strong>${co}</strong> checked out</span>` : ''}`;
    };

    const renderTable = (): void => {
      const tbody = d.getElementById('drm-tbody');
      if (!tbody) return;
      const groups = [...new Set(filtered.map((x: any) => x.group))];
      let html = '';
      groups.forEach((g: any) => {
        html += `<tr class="drm-grp"><td colspan="5">${g}</td></tr>`;
        filtered.filter((x: any) => x.group === g).forEach((doc: any) => {
          html += `<tr class="${selected === doc.id ? 'drm-sel' : ''}" data-id="${doc.id}"><td class="drm-ic"><div class="drm-docid">${doc.id}</div><span class="drm-badge drm-${doc.type}">${doc.type}</span></td><td class="drm-tc">${doc.title}</td>${zoneCell(doc.zones.drafts)}${zoneCell(doc.zones.published)}${zoneCell(doc.zones.official)}</tr>`;
        });
      });
      tbody.innerHTML = html;
      tbody.querySelectorAll('tr[data-id]').forEach((row: Element) => {
        row.addEventListener('click', () => {
          const id = (row as HTMLElement).dataset['id']!;
          selected = selected === id ? null : id;
          renderTable();
          if (selected) showDetail(id); else hideDetail();
        });
      });
    };

    const showDetail = (id: string): void => {
      const panel = d.getElementById('drm-detail');
      if (!panel) return;
      const doc = allDocs.find((x: any) => x.id === id);
      if (!doc) return;
      panel.style.display = '';
      const zkeys = ['drafts','published','official'];
      const znames: Record<string, string> = { drafts:'Drafts', published:'Published', official:'Official' };
      const zonesHtml = zkeys.map((zk: string) => {
        const z = doc.zones[zk];
        if (!z) return `<div class="drm-dzone"><div class="drm-dzone-name ${zk}">${znames[zk]}</div><div class="drm-absent">Not present</div></div>`;
        return `<div class="drm-dzone"><div class="drm-dzone-name ${zk}">${znames[zk]}</div><div class="drm-dfield"><span class="drm-dlbl">Revision</span><span class="drm-rev drm-r${z.rev}">Rev ${z.rev}</span></div><div class="drm-dfield"><span class="drm-dlbl">Version</span><span class="drm-ver ${z.verType === 'major' ? 'major' : ''}">${z.verType === 'major' ? '&#9679;' : '&#9675;'} ${z.ver}</span></div><div class="drm-dfield"><span class="drm-dlbl">DCO</span><span class="drm-dco ${dcoClass(z.dcoStatus)}">${z.dco || 'none'} &middot; ${z.dcoStatus}</span></div><div class="drm-dfield"><span class="drm-dlbl">Checkout</span><span style="font-size:10px;color:${z.checkedOut ? '#92400E' : '#065F46'}">${z.checkedOut ? '&#9888; ' + z.checkedOut : '&#10003; Checked in'}</span></div><div class="drm-dfield"><span class="drm-dlbl">Modified</span><span class="drm-dval">${z.modified}</span></div><div class="drm-dfield" style="align-items:flex-start"><span class="drm-dlbl">File</span><span class="drm-dval">${z.file}</span></div><a class="drm-open-sp" href="${z.webUrl}" target="_blank">Open in SharePoint &#8599;</a></div>`;
      }).join('');
      panel.innerHTML = `<div class="drm-dhead"><span class="drm-dhead-id">${doc.id}</span><span class="drm-dhead-title">${doc.title}</span><button class="drm-dhead-close" id="drm-close">&#x2715;</button></div><div class="drm-dbody">${zonesHtml}</div>`;
      const closeBtn = d.getElementById('drm-close');
      if (closeBtn) closeBtn.addEventListener('click', () => { selected = null; hideDetail(); renderTable(); });
    };

    const hideDetail = (): void => {
      const panel = d.getElementById('drm-detail');
      if (panel) { panel.style.display = 'none'; panel.innerHTML = ''; }
    };

    const applyFilters = (): void => {
      const tf = (d.getElementById('drm-tf') as HTMLSelectElement)?.value || 'all';
      const df = (d.getElementById('drm-df') as HTMLSelectElement)?.value || 'all';
      const zf = (d.getElementById('drm-zf') as HTMLSelectElement)?.value || 'all';
      const qf = ((d.getElementById('drm-qf') as HTMLInputElement)?.value || '').toLowerCase();
      filtered = allDocs.filter((doc: any) => {
        const allZ = ['drafts','published','official'].map((k: string) => doc.zones[k]).filter(Boolean);
        const mt = tf === 'all' || doc.type === tf;
        const mq = !qf || doc.id.toLowerCase().includes(qf) || doc.title.toLowerCase().includes(qf);
        const md = df === 'all' || (df === 'none' && allZ.every((z: any) => !z.dco)) || (df !== 'none' && allZ.some((z: any) => z.dco === df));
        const hasd = !!doc.zones.drafts, hasp = !!doc.zones.published;
        const mz = zf === 'all' || (zf === 'drafts' && hasd) || (zf === 'published' && hasp) || (zf === 'drafts-only' && hasd && !hasp) || (zf === 'both' && hasd && hasp);
        return mt && mq && md && mz;
      });
      selected = null;
      hideDetail();
      renderSummary();
      renderTable();
    };

    ['drm-tf','drm-df','drm-zf'].forEach((id: string) => {
      const el = d.getElementById(id);
      if (el) el.addEventListener('change', applyFilters);
    });
    const qEl = d.getElementById('drm-qf');
    if (qEl) qEl.addEventListener('input', applyFilters);

    renderSummary();
    renderTable();
  }

  private _renderDrmError(): void {
    const d = this.iframe && this.iframe.contentDocument;
    if (!d) return;
    const loading = d.getElementById('drm-loading');
    if (loading) loading.innerHTML = '<span style="color:#B91C1C">Error loading document matrix. Check SharePoint permissions and try refreshing.</span>';
  }

'@

  # Insert before the SPFx boilerplate comment near the end of the class
  $src = $src -replace '(  // ---------------------------------------------------------------------------\s*\n  // SPFx boilerplate)', "$drmMethod`$1"

  Write-Host "Writing patched file..." -ForegroundColor Cyan
  [System.IO.File]::WriteAllText((Resolve-Path $webPartFile), $src, [System.Text.Encoding]::UTF8)
  Write-Host "Patch applied successfully." -ForegroundColor Green
}

# -------------------------------------------------------------------------
# BUILD
# -------------------------------------------------------------------------
Write-Host "`nBuilding..." -ForegroundColor Cyan
Remove-Item -Recurse -Force ".\temp" -ErrorAction SilentlyContinue

$bundleResult = & cmd /c "node_modules\.bin\gulp bundle --ship 2>&1"
Write-Host $bundleResult
if ($LASTEXITCODE -ne 0) { Write-Host "Bundle FAILED" -ForegroundColor Red; exit 1 }

$packageResult = & cmd /c "node_modules\.bin\gulp package-solution --ship 2>&1"
Write-Host $packageResult
if ($LASTEXITCODE -ne 0) { Write-Host "Package FAILED" -ForegroundColor Red; exit 1 }

Write-Host "`nBuild complete." -ForegroundColor Green

# -------------------------------------------------------------------------
# DEPLOY
# -------------------------------------------------------------------------
Write-Host "`nDeploying to SharePoint..." -ForegroundColor Cyan
Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive
$app = Add-PnPApp -Path ".\sharepoint\solution\imp-9177-spfx.sppkg" -Scope Site -Overwrite -Publish
Write-Host "Deployed. New App ID: $($app.Id)" -ForegroundColor Green
Write-Host "`nDone. Go to the portal, open Web Part settings, and select 'Document Repository' from the dashboard dropdown." -ForegroundColor Cyan
