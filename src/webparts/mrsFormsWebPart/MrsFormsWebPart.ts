import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IMrsFormsWebPartProps {}

export default class MrsFormsWebPart extends BaseClientSideWebPart<IMrsFormsWebPartProps> {

  private _docs: any[] = [];
  private _gaps: any[] = [];
  private _docFilter: string = 'all';
  private _gapFilter: string = 'all';
  private _docSearch: string = '';
  private _gapSearch: string = '';
  private _modalData: any = null;
  private _modalList: string = '';

  public render(): void {
    this.domElement.innerHTML = this._getShell();
    this._loadData();
    this._bindGlobalEvents();
  }

  private _getShell(): string {
    return `
<div id="mrsRoot" style="font-family:'DM Sans',Segoe UI,sans-serif;background:#f4f6f9;min-height:100vh;color:#1a2332">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600;700&display=swap');
    #mrsRoot *{box-sizing:border-box;margin:0;padding:0}
    #mrsRoot .ph{background:linear-gradient(135deg,#0a3259 0%,#0f4c81 60%,#1a6bb5 100%);padding:28px 32px 0;position:relative;overflow:hidden}
    #mrsRoot .ph::before{content:'';position:absolute;top:-40px;right:-40px;width:200px;height:200px;border-radius:50%;background:rgba(255,255,255,.04)}
    #mrsRoot .ph-eye{font-family:'DM Mono',monospace;font-size:10px;letter-spacing:2px;color:rgba(255,255,255,.5);text-transform:uppercase;margin-bottom:6px}
    #mrsRoot .ph-title{font-size:22px;font-weight:700;color:#fff;letter-spacing:-.3px}
    #mrsRoot .ph-sub{font-size:13px;color:rgba(255,255,255,.65);margin-top:4px}
    #mrsRoot .ph-ts{font-family:'DM Mono',monospace;font-size:10px;color:rgba(255,255,255,.4);text-align:right;white-space:nowrap;margin-top:4px}
    #mrsRoot .ph-top{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;position:relative;z-index:1}
    #mrsRoot .kpi-strip{background:#0a3259;border-top:1px solid rgba(255,255,255,.08);padding:0 32px;display:flex;overflow-x:auto;margin:16px -32px 0}
    #mrsRoot .kpi-c{padding:14px 24px 14px 0;margin-right:24px;border-right:1px solid rgba(255,255,255,.08);min-width:110px;flex-shrink:0}
    #mrsRoot .kpi-c:last-child{border-right:none}
    #mrsRoot .kpi-v{font-family:'DM Mono',monospace;font-size:24px;font-weight:500;color:#fff;line-height:1}
    #mrsRoot .kpi-v.red{color:#ef9a9a} #mrsRoot .kpi-v.green{color:#a5d6a7} #mrsRoot .kpi-v.amber{color:#ffcc80}
    #mrsRoot .kpi-l{font-size:10px;color:rgba(255,255,255,.45);letter-spacing:.8px;text-transform:uppercase;margin-top:4px}
    #mrsRoot .kpi-load{animation:mrsPulse 1.2s ease-in-out infinite;color:rgba(255,255,255,.2)}
    @keyframes mrsPulse{0%,100%{opacity:.3}50%{opacity:1}}
    #mrsRoot .alert{margin:16px 32px 0;padding:12px 16px;border-radius:8px;border-left:4px solid #c62828;background:#fde8e8;color:#c62828;display:none;align-items:center;gap:10px;font-size:13px;font-weight:500}
    #mrsRoot .alert.show{display:flex}
    #mrsRoot .content{padding:20px 32px 32px}
    #mrsRoot .sec-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;margin-top:24px}
    #mrsRoot .sec-title{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#0f4c81;display:flex;align-items:center;gap:8px}
    #mrsRoot .sec-title::before{content:'';width:3px;height:14px;background:#0f4c81;border-radius:2px;display:block}
    #mrsRoot .sec-cnt{font-family:'DM Mono',monospace;font-size:11px;color:#5a6a7e;background:#e0e6ed;padding:2px 8px;border-radius:10px}
    #mrsRoot .fbar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center}
    #mrsRoot .fbtn{font-family:'DM Sans',sans-serif;font-size:11px;font-weight:600;padding:5px 12px;border-radius:20px;border:1px solid #e0e6ed;background:#fff;color:#5a6a7e;cursor:pointer;transition:all .15s}
    #mrsRoot .fbtn:hover,#mrsRoot .fbtn.active{background:#0f4c81;color:#fff;border-color:#0f4c81}
    #mrsRoot .sbox{font-family:'DM Sans',sans-serif;font-size:12px;padding:5px 12px;border-radius:20px;border:1px solid #e0e6ed;background:#fff;color:#1a2332;outline:none;min-width:180px;margin-left:auto}
    #mrsRoot .sbox:focus{border-color:#0f4c81}
    #mrsRoot .tcard{background:#fff;border-radius:10px;border:1px solid #e0e6ed;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.05)}
    #mrsRoot table{width:100%;border-collapse:collapse;font-size:12.5px}
    #mrsRoot thead th{background:#f8fafc;padding:10px 14px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;border-bottom:1px solid #e0e6ed;white-space:nowrap;cursor:pointer;user-select:none}
    #mrsRoot thead th:hover{color:#0f4c81}
    #mrsRoot tbody tr{border-bottom:1px solid #f0f4f8;transition:background .1s;cursor:pointer}
    #mrsRoot tbody tr:last-child{border-bottom:none}
    #mrsRoot tbody tr:hover{background:#f8fafc}
    #mrsRoot td{padding:10px 14px;vertical-align:middle}
    #mrsRoot .cid{font-family:'DM Mono',monospace;font-size:11px;color:#0f4c81;font-weight:500}
    #mrsRoot .cmut{font-size:11px;color:#5a6a7e}
    #mrsRoot .cdate{font-family:'DM Mono',monospace;font-size:11px;color:#5a6a7e;white-space:nowrap}
    #mrsRoot .pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:10.5px;font-weight:600;white-space:nowrap}
    #mrsRoot .pr{background:#fde8e8;color:#c62828} #mrsRoot .pg{background:#e8f5e9;color:#2e7d32}
    #mrsRoot .pa{background:#fff3e0;color:#e65100} #mrsRoot .pb{background:#e3f2fd;color:#0d47a1}
    #mrsRoot .pz{background:#f0f0f0;color:#616161}
    #mrsRoot .ws{display:inline-block;padding:2px 8px;border-radius:4px;font-family:'DM Mono',monospace;font-size:10px;font-weight:700;letter-spacing:.5px}
    #mrsRoot .wsW0{background:#f3e5f5;color:#4a148c} #mrsRoot .wsW1{background:#e3f2fd;color:#0d47a1} #mrsRoot .wsW2{background:#e8f5e9;color:#2e7d32}
    #mrsRoot .gsb{display:inline-block;padding:2px 8px;border-radius:4px;font-family:'DM Mono',monospace;font-size:10px;font-weight:700}
    #mrsRoot .gsPSO{background:#e3f2fd;color:#0d47a1} #mrsRoot .gsSV{background:#fce4ec;color:#880e4f}
    #mrsRoot .gsCOA{background:#fff3e0;color:#e65100} #mrsRoot .gsPC{background:#e8f5e9;color:#1b5e20} #mrsRoot .gsREC{background:#f3e5f5;color:#4a148c}
    #mrsRoot .btn-new{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:7px 16px;border-radius:6px;background:#0f4c81;color:#fff;border:none;cursor:pointer;display:inline-flex;align-items:center;gap:6px;transition:background .15s;text-decoration:none}
    #mrsRoot .btn-new:hover{background:#1a6bb5}
    #mrsRoot .load-row td{padding:24px;text-align:center;color:#5a6a7e;font-size:12px;font-family:'DM Mono',monospace}
    #mrsRoot .modal-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:9999;align-items:center;justify-content:center;padding:20px}
    #mrsRoot .modal-ov.open{display:flex}
    #mrsRoot .modal{background:#fff;border-radius:12px;width:100%;max-width:600px;max-height:85vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
    #mrsRoot .modal-hdr{padding:20px 24px 16px;border-bottom:1px solid #e0e6ed;display:flex;justify-content:space-between;align-items:flex-start;position:sticky;top:0;background:#fff;z-index:1}
    #mrsRoot .modal-title{font-size:15px;font-weight:700;color:#0f4c81}
    #mrsRoot .modal-sub{font-family:'DM Mono',monospace;font-size:10px;color:#5a6a7e;margin-top:2px}
    #mrsRoot .modal-x{background:none;border:none;font-size:22px;cursor:pointer;color:#5a6a7e;padding:0 4px;line-height:1}
    #mrsRoot .modal-body{padding:20px 24px}
    #mrsRoot .fg{margin-bottom:14px}
    #mrsRoot .fl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;margin-bottom:4px}
    #mrsRoot .fv{font-size:13px;color:#1a2332;line-height:1.5;padding:8px 10px;background:#f8fafc;border-radius:6px;border:1px solid #e0e6ed}
    #mrsRoot .fv.mono{font-family:'DM Mono',monospace;font-size:12px}
    #mrsRoot .modal-ft{padding:16px 24px;border-top:1px solid #e0e6ed;display:flex;justify-content:flex-end;gap:8px}
    #mrsRoot .btn-sec{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:8px 18px;border-radius:6px;background:#f0f4f8;color:#1a2332;border:1px solid #e0e6ed;cursor:pointer}
    #mrsRoot .btn-pri{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:8px 18px;border-radius:6px;background:#0f4c81;color:#fff;border:none;cursor:pointer;text-decoration:none;display:inline-block}
  </style>

  <div class="ph">
    <div class="ph-top">
      <div>
        <div class="ph-eye">IMP9177 · 3H Pharmaceuticals LLC</div>
        <div class="ph-title">Document Control &amp; Gap Registry</div>
        <div class="ph-sub" id="mrsSub">Loading live data...</div>
      </div>
      <div><div class="ph-ts" id="mrsTs">Fetching...</div></div>
    </div>
    <div class="kpi-strip">
      <div class="kpi-c"><div class="kpi-v" id="kTD"><span class="kpi-load">—</span></div><div class="kpi-l">Total Docs</div></div>
      <div class="kpi-c"><div class="kpi-v red" id="kPD"><span class="kpi-load">—</span></div><div class="kpi-l">Past Due</div></div>
      <div class="kpi-c"><div class="kpi-v green" id="kW2"><span class="kpi-load">—</span></div><div class="kpi-l">W2 Delivered</div></div>
      <div class="kpi-c"><div class="kpi-v" id="kTG"><span class="kpi-load">—</span></div><div class="kpi-l">Total Gaps</div></div>
      <div class="kpi-c"><div class="kpi-v red" id="kOG"><span class="kpi-load">—</span></div><div class="kpi-l">Open Gaps</div></div>
      <div class="kpi-c"><div class="kpi-v green" id="kCG"><span class="kpi-load">—</span></div><div class="kpi-l">Closed Gaps</div></div>
    </div>
  </div>

  <div class="alert" id="mrsAlert"><span>⚠️</span><span id="mrsAlertTxt"></span></div>

  <div class="content">
    <div class="sec-hdr">
      <div class="sec-title">Document Register</div>
      <div style="display:flex;align-items:center;gap:8px">
        <span class="sec-cnt" id="docCnt">—</span>
        <button class="btn-new" id="btnNewDoc">+ New Document</button>
      </div>
    </div>
    <div class="fbar" id="docFbar">
      <button class="fbtn active" data-f="all">All</button>
      <button class="fbtn" data-f="W0">W0</button>
      <button class="fbtn" data-f="W1">W1</button>
      <button class="fbtn" data-f="W2">W2</button>
      <button class="fbtn" data-f="pastdue">⚠ Past Due</button>
      <input class="sbox" id="docSearch" placeholder="Search docs...">
    </div>
    <div class="tcard">
      <table>
        <thead><tr>
          <th data-sort-docs="Title">Doc ID</th><th>WS</th><th>Rev</th>
          <th data-sort-docs="DocStatus">Status</th><th data-sort-docs="DeliveryDate">Delivered</th>
          <th>Approver</th><th>Co-Signer</th><th>Gap Refs</th>
        </tr></thead>
        <tbody id="docBody"><tr class="load-row"><td colspan="8">Loading documents...</td></tr></tbody>
      </table>
    </div>

    <div class="sec-hdr">
      <div class="sec-title">Gap Registry</div>
      <div style="display:flex;align-items:center;gap:8px">
        <span class="sec-cnt" id="gapCnt">—</span>
        <button class="btn-new" id="btnNewGap">+ New Gap</button>
      </div>
    </div>
    <div class="fbar" id="gapFbar">
      <button class="fbtn active" data-g="all">All</button>
      <button class="fbtn" data-g="GAP-PSO">GAP-PSO</button>
      <button class="fbtn" data-g="GAP-SV">GAP-SV</button>
      <button class="fbtn" data-g="GAP-COA">GAP-COA</button>
      <button class="fbtn" data-g="GAP-PC">GAP-PC</button>
      <button class="fbtn" data-g="GAP-REC">GAP-REC</button>
      <button class="fbtn" data-g="open">Open Only</button>
      <input class="sbox" id="gapSearch" placeholder="Search gaps...">
    </div>
    <div class="tcard">
      <table>
        <thead><tr>
          <th data-sort-gaps="Title">Gap ID</th><th>Series</th><th>Regulation</th>
          <th>Description</th><th data-sort-gaps="GapStatus">Status</th><th>Closing Doc</th>
          <th data-sort-gaps="ClosureDate">Closed</th>
        </tr></thead>
        <tbody id="gapBody"><tr class="load-row"><td colspan="7">Loading gaps...</td></tr></tbody>
      </table>
    </div>
  </div>

  <div class="modal-ov" id="mrsModal">
    <div class="modal">
      <div class="modal-hdr">
        <div><div class="modal-title" id="mTitle"></div><div class="modal-sub" id="mSub"></div></div>
        <button class="modal-x" id="mClose">×</button>
      </div>
      <div class="modal-body" id="mBody"></div>
      <div class="modal-ft">
        <button class="btn-sec" id="mDismiss">Close</button>
        <a class="btn-pri" id="mEdit" href="#" target="_blank">Edit in SharePoint ↗</a>
      </div>
    </div>
  </div>
</div>`;
  }

  private async _loadData(): Promise<void> {
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    try {
      const [docsRes, gapsRes]: [SPHttpClientResponse, SPHttpClientResponse] = await Promise.all([
        this.context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('MRS_Documents')/items?$select=Id,Title,DocNumber,WorkStream,RevLevel,DocStatus,DeliveryDate,DCOApprover,DCOCoSigner,LinkedGapIDs&$top=500&$orderby=Title`,
          SPHttpClient.configurations.v1
        ),
        this.context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('MRS_GapRegistry')/items?$select=Id,Title,GapSeries,RegulationRef,GapDesc,GapStatus,ClosingDocRef,ClosureDate&$top=500&$orderby=Title`,
          SPHttpClient.configurations.v1
        )
      ]);

      const [docsData, gapsData] = await Promise.all([docsRes.json(), gapsRes.json()]);
      this._docs = docsData.value || [];
      this._gaps = gapsData.value || [];

    } catch (e) {
      console.error('MRS Forms: data load failed', e);
      this._docs = [];
      this._gaps = [];
    }

    this._updateKPIs();
    this._renderDocs();
    this._renderGaps();
    this._updateTimestamp();
  }

  private _fmt(s: string): string {
    if (!s) return '—';
    try { return new Date(s).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' }); }
    catch { return s; }
  }

  private _pill(s: string): string {
    if (!s) return '<span class="pill pz">—</span>';
    const l = s.toLowerCase();
    if (l.includes('past due') || l.includes('critical') || l.includes('overdue') || l.includes('suspended') || l.includes('blocking'))
      return `<span class="pill pr">${s}</span>`;
    if (l.includes('complete') || l.includes('closed'))
      return `<span class="pill pg">${s}</span>`;
    if (l.includes('delivered') || l.includes('progress') || l.includes('partial'))
      return `<span class="pill pa">${s}</span>`;
    if (l.includes('pending') || l.includes('open') || l.includes('issued') || l.includes('new'))
      return `<span class="pill pb">${s}</span>`;
    return `<span class="pill pz">${s}</span>`;
  }

  private _updateKPIs(): void {
    const pastDue = this._docs.filter(d => (d.DocStatus || '').toLowerCase().includes('past due')).length;
    const w2del   = this._docs.filter(d => d.WorkStream === 'W2' && (d.DocStatus || '').toLowerCase().includes('delivered')).length;
    const open    = this._gaps.filter(g => !(g.GapStatus || '').toLowerCase().includes('closed')).length;

    this._set('kTD', String(this._docs.length));
    this._set('kPD', String(pastDue));
    this._set('kW2', String(w2del));
    this._set('kTG', String(this._gaps.length));
    this._set('kOG', String(open));
    this._set('kCG', String(this._gaps.length - open));
    this._set('mrsSub', `${this._docs.length} controlled documents · ${open} open gaps · ${pastDue} sign-off(s) past due`);

    if (pastDue > 0) {
      const ids = this._docs.filter(d => (d.DocStatus || '').toLowerCase().includes('past due')).map(d => d.Title).join(', ');
      this._set('mrsAlertTxt', `${pastDue} sign-off(s) PAST DUE: ${ids}`);
      const el = this.domElement.querySelector('#mrsAlert') as HTMLElement;
      if (el) el.classList.add('show');
    }
  }

  private _filteredDocs(): any[] {
    let l = this._docs;
    if (this._docFilter === 'W0') l = l.filter(d => d.WorkStream === 'W0');
    else if (this._docFilter === 'W1') l = l.filter(d => d.WorkStream === 'W1');
    else if (this._docFilter === 'W2') l = l.filter(d => d.WorkStream === 'W2');
    else if (this._docFilter === 'pastdue') l = l.filter(d => (d.DocStatus || '').toLowerCase().includes('past due'));
    if (this._docSearch) {
      const q = this._docSearch.toLowerCase();
      l = l.filter(d => (d.Title || '').toLowerCase().includes(q) || (d.DocStatus || '').toLowerCase().includes(q) || (d.LinkedGapIDs || '').toLowerCase().includes(q));
    }
    return l;
  }

  private _renderDocs(): void {
    const l = this._filteredDocs();
    this._set('docCnt', String(l.length));
    const tbody = this.domElement.querySelector('#docBody') as HTMLElement;
    if (!tbody) return;
    if (!l.length) {
      tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;padding:32px;color:#aaa">No documents match filter</td></tr>';
      return;
    }
    tbody.innerHTML = l.map(d => `
      <tr data-id="${d.Id}" data-list="doc">
        <td><span class="cid">${d.Title || '—'}</span></td>
        <td><span class="ws ws${d.WorkStream || ''}">${d.WorkStream || '—'}</span></td>
        <td><span class="cmut">${d.RevLevel || '—'}</span></td>
        <td>${this._pill(d.DocStatus)}</td>
        <td><span class="cdate">${this._fmt(d.DeliveryDate)}</span></td>
        <td><span class="cmut">${d.DCOApprover || '—'}</span></td>
        <td><span class="cmut">${d.DCOCoSigner || '—'}</span></td>
        <td><span class="cmut" style="font-size:10.5px">${d.LinkedGapIDs || '—'}</span></td>
      </tr>`).join('');
  }

  private _filteredGaps(): any[] {
    let l = this._gaps;
    if (this._gapFilter === 'open') l = l.filter(g => !(g.GapStatus || '').toLowerCase().includes('closed'));
    else if (this._gapFilter !== 'all') l = l.filter(g => g.GapSeries === this._gapFilter);
    if (this._gapSearch) {
      const q = this._gapSearch.toLowerCase();
      l = l.filter(g => (g.Title || '').toLowerCase().includes(q) || (g.GapDesc || '').toLowerCase().includes(q) || (g.RegulationRef || '').toLowerCase().includes(q));
    }
    return l;
  }

  private _renderGaps(): void {
    const l = this._filteredGaps();
    this._set('gapCnt', String(l.length));
    const tbody = this.domElement.querySelector('#gapBody') as HTMLElement;
    if (!tbody) return;
    if (!l.length) {
      tbody.innerHTML = '<tr><td colspan="7" style="text-align:center;padding:32px;color:#aaa">No gaps match filter</td></tr>';
      return;
    }
    const seriesClass = (s: string) => {
      const m: { [k: string]: string } = { 'GAP-PSO': 'gsPSO', 'GAP-SV': 'gsSV', 'GAP-COA': 'gsCOA', 'GAP-PC': 'gsPC', 'GAP-REC': 'gsREC' };
      return m[s] || '';
    };
    tbody.innerHTML = l.map(g => `
      <tr data-id="${g.Id}" data-list="gap">
        <td><span class="cid">${g.Title || '—'}</span></td>
        <td><span class="gsb ${seriesClass(g.GapSeries || '')}">${g.GapSeries || '—'}</span></td>
        <td><span class="cmut" style="font-family:'DM Mono',monospace;font-size:10.5px">${g.RegulationRef || '—'}</span></td>
        <td><span style="font-size:12px;display:block;max-width:260px">${g.GapDesc || '—'}</span></td>
        <td>${this._pill(g.GapStatus)}</td>
        <td><span class="cmut">${g.ClosingDocRef || '—'}</span></td>
        <td><span class="cdate">${this._fmt(g.ClosureDate)}</span></td>
      </tr>`).join('');
  }

  private _bindGlobalEvents(): void {
    const root = this.domElement;
    const siteUrl = this.context.pageContext.web.absoluteUrl;

    // Doc filter buttons
    root.querySelector('#docFbar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-f]') as HTMLElement;
      if (!btn) return;
      root.querySelectorAll('#docFbar .fbtn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      this._docFilter = btn.dataset.f || 'all';
      this._renderDocs();
    });

    // Gap filter buttons
    root.querySelector('#gapFbar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-g]') as HTMLElement;
      if (!btn) return;
      root.querySelectorAll('#gapFbar .fbtn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      this._gapFilter = btn.dataset.g || 'all';
      this._renderGaps();
    });

    // Search
    root.querySelector('#docSearch')?.addEventListener('input', (e) => {
      this._docSearch = (e.target as HTMLInputElement).value;
      this._renderDocs();
    });
    root.querySelector('#gapSearch')?.addEventListener('input', (e) => {
      this._gapSearch = (e.target as HTMLInputElement).value;
      this._renderGaps();
    });

    // Sort doc headers
    root.querySelectorAll('[data-sort-docs]').forEach(th => {
      th.addEventListener('click', () => {
        const f = (th as HTMLElement).dataset.sortDocs || '';
        this._docs.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || '')));
        this._renderDocs();
      });
    });
    root.querySelectorAll('[data-sort-gaps]').forEach(th => {
      th.addEventListener('click', () => {
        const f = (th as HTMLElement).dataset.sortGaps || '';
        this._gaps.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || '')));
        this._renderGaps();
      });
    });

    // Row click → modal
    root.querySelector('#docBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('tr[data-id]') as HTMLElement;
      if (!row) return;
      const id = parseInt(row.dataset.id || '0');
      const d = this._docs.find(x => x.Id === id);
      if (!d) return;
      this._set('mTitle', d.Title);
      this._set('mSub', `Document Register · ${d.WorkStream} · ${d.RevLevel}`);
      this._set('mBody', `
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(d.DocStatus)}</div></div>
        <div class="fg"><div class="fl">Work Stream</div><div class="fv"><span class="ws ws${d.WorkStream}">${d.WorkStream}</span></div></div>
        <div class="fg"><div class="fl">Revision</div><div class="fv mono">${d.RevLevel || '—'}</div></div>
        <div class="fg"><div class="fl">Delivery Date</div><div class="fv mono">${this._fmt(d.DeliveryDate)}</div></div>
        <div class="fg"><div class="fl">QA Approver</div><div class="fv">${d.DCOApprover || '—'}</div></div>
        <div class="fg"><div class="fl">Co-Signer</div><div class="fv">${d.DCOCoSigner || '—'}</div></div>
        <div class="fg"><div class="fl">Linked Gap IDs</div><div class="fv">${d.LinkedGapIDs || 'None'}</div></div>`);
      const editLink = root.querySelector('#mEdit') as HTMLAnchorElement;
      if (editLink) editLink.href = `${siteUrl}/Lists/MRS_Documents/EditForm.aspx?ID=${id}`;
      (root.querySelector('#mrsModal') as HTMLElement)?.classList.add('open');
    });

    root.querySelector('#gapBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('tr[data-id]') as HTMLElement;
      if (!row) return;
      const id = parseInt(row.dataset.id || '0');
      const g = this._gaps.find(x => x.Id === id);
      if (!g) return;
      this._set('mTitle', g.Title);
      this._set('mSub', `Gap Registry · ${g.GapSeries}`);
      this._set('mBody', `
        <div class="fg"><div class="fl">Series</div><div class="fv">${g.GapSeries || '—'}</div></div>
        <div class="fg"><div class="fl">Regulation</div><div class="fv mono">${g.RegulationRef || '—'}</div></div>
        <div class="fg"><div class="fl">Description</div><div class="fv">${g.GapDesc || '—'}</div></div>
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(g.GapStatus)}</div></div>
        <div class="fg"><div class="fl">Closing Doc</div><div class="fv">${g.ClosingDocRef || '—'}</div></div>
        <div class="fg"><div class="fl">Closure Date</div><div class="fv mono">${this._fmt(g.ClosureDate)}</div></div>`);
      const editLink = root.querySelector('#mEdit') as HTMLAnchorElement;
      if (editLink) editLink.href = `${siteUrl}/Lists/MRS_GapRegistry/EditForm.aspx?ID=${id}`;
      (root.querySelector('#mrsModal') as HTMLElement)?.classList.add('open');
    });

    // Modal close
    const closeModal = () => (root.querySelector('#mrsModal') as HTMLElement)?.classList.remove('open');
    root.querySelector('#mClose')?.addEventListener('click', closeModal);
    root.querySelector('#mDismiss')?.addEventListener('click', closeModal);
    root.querySelector('#mrsModal')?.addEventListener('click', (e) => {
      if ((e.target as HTMLElement).id === 'mrsModal') closeModal();
    });

    // New buttons
    root.querySelector('#btnNewDoc')?.addEventListener('click', () => {
      window.open(`${siteUrl}/Lists/MRS_Documents/NewForm.aspx`, '_blank');
    });
    root.querySelector('#btnNewGap')?.addEventListener('click', () => {
      window.open(`${siteUrl}/Lists/MRS_GapRegistry/NewForm.aspx`, '_blank');
    });
  }

  private _set(id: string, html: string): void {
    const el = this.domElement.querySelector(`#${id}`) as HTMLElement;
    if (el) el.innerHTML = html;
  }

  private _updateTimestamp(): void {
    this._set('mrsTs', 'Refreshed ' + new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' }));
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
