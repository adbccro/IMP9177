import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface ISupFormsWebPartProps {}

export default class SupFormsWebPart extends BaseClientSideWebPart<ISupFormsWebPartProps> {
  private _sup: any[] = [];
  private _capa: any[] = [];
  private _supFilter: string = 'all';
  private _supSearch: string = '';

  public render(): void {
    this.domElement.innerHTML = this._getShell();
    this._loadData();
    this._bindEvents();
  }

  private _getShell(): string {
    return `
<div id="supRoot" style="font-family:'DM Sans',Segoe UI,sans-serif;background:#f4f6f9;min-height:100vh;color:#1a2332">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600;700&display=swap');
    #supRoot *{box-sizing:border-box;margin:0;padding:0}
    #supRoot .ph{background:linear-gradient(135deg,#0a3259 0%,#0f4c81 60%,#1a6bb5 100%);padding:28px 32px 0;overflow:hidden}
    #supRoot .ph-top{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;position:relative;z-index:1}
    #supRoot .ph-eye{font-family:'DM Mono',monospace;font-size:10px;letter-spacing:2px;color:rgba(255,255,255,.5);text-transform:uppercase;margin-bottom:6px}
    #supRoot .ph-title{font-size:22px;font-weight:700;color:#fff}
    #supRoot .ph-sub{font-size:13px;color:rgba(255,255,255,.65);margin-top:4px}
    #supRoot .ph-ts{font-family:'DM Mono',monospace;font-size:10px;color:rgba(255,255,255,.4);text-align:right;white-space:nowrap;margin-top:4px}
    #supRoot .kpi-strip{background:#0a3259;border-top:1px solid rgba(255,255,255,.08);padding:0 32px;display:flex;overflow-x:auto;margin:16px -32px 0}
    #supRoot .kpi-c{padding:14px 24px 14px 0;margin-right:24px;border-right:1px solid rgba(255,255,255,.08);min-width:110px;flex-shrink:0}
    #supRoot .kpi-c:last-child{border-right:none}
    #supRoot .kpi-v{font-family:'DM Mono',monospace;font-size:24px;font-weight:500;color:#fff;line-height:1}
    #supRoot .kpi-v.red{color:#ef9a9a} #supRoot .kpi-v.green{color:#a5d6a7} #supRoot .kpi-v.amber{color:#ffcc80}
    #supRoot .kpi-l{font-size:10px;color:rgba(255,255,255,.45);letter-spacing:.8px;text-transform:uppercase;margin-top:4px}
    #supRoot .kpi-load{animation:supPulse 1.2s ease-in-out infinite;color:rgba(255,255,255,.2)}
    @keyframes supPulse{0%,100%{opacity:.3}50%{opacity:1}}
    #supRoot .alert{margin:16px 32px 0;padding:12px 16px;border-radius:8px;border-left:4px solid #c62828;background:#fde8e8;color:#c62828;display:none;align-items:center;gap:10px;font-size:13px;font-weight:500}
    #supRoot .alert.show{display:flex}
    #supRoot .content{padding:20px 32px 32px}
    #supRoot .sec-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;margin-top:24px}
    #supRoot .sec-title{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#0f4c81;display:flex;align-items:center;gap:8px}
    #supRoot .sec-title::before{content:'';width:3px;height:14px;background:#0f4c81;border-radius:2px;display:block}
    #supRoot .sec-cnt{font-family:'DM Mono',monospace;font-size:11px;color:#5a6a7e;background:#e0e6ed;padding:2px 8px;border-radius:10px}
    #supRoot .fbar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center}
    #supRoot .fbtn{font-family:'DM Sans',sans-serif;font-size:11px;font-weight:600;padding:5px 12px;border-radius:20px;border:1px solid #e0e6ed;background:#fff;color:#5a6a7e;cursor:pointer;transition:all .15s}
    #supRoot .fbtn:hover,#supRoot .fbtn.active{background:#0f4c81;color:#fff;border-color:#0f4c81}
    #supRoot .sbox{font-family:'DM Sans',sans-serif;font-size:12px;padding:5px 12px;border-radius:20px;border:1px solid #e0e6ed;background:#fff;outline:none;min-width:200px;margin-left:auto}
    #supRoot .sbox:focus{border-color:#0f4c81}
    #supRoot .sup-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;margin-bottom:20px}
    #supRoot .sc{background:#fff;border-radius:10px;border:1px solid #e0e6ed;padding:14px 16px;cursor:pointer;transition:all .15s;box-shadow:0 1px 3px rgba(0,0,0,.04)}
    #supRoot .sc:hover{border-color:#0f4c81;box-shadow:0 4px 12px rgba(15,76,129,.1);transform:translateY(-1px)}
    #supRoot .sc.crit{border-left:3px solid #c62828} #supRoot .sc.warn{border-left:3px solid #e65100} #supRoot .sc.ok{border-left:3px solid #2e7d32}
    #supRoot .sc-name{font-size:13px;font-weight:700;color:#1a2332;margin-bottom:4px}
    #supRoot .sc-comm{font-size:11px;color:#5a6a7e;margin-bottom:8px}
    #supRoot .tcard{background:#fff;border-radius:10px;border:1px solid #e0e6ed;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.05)}
    #supRoot table{width:100%;border-collapse:collapse;font-size:12.5px}
    #supRoot thead th{background:#f8fafc;padding:10px 14px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;border-bottom:1px solid #e0e6ed;white-space:nowrap;cursor:pointer}
    #supRoot thead th:hover{color:#0f4c81}
    #supRoot tbody tr{border-bottom:1px solid #f0f4f8;transition:background .1s;cursor:pointer}
    #supRoot tbody tr:last-child{border-bottom:none}
    #supRoot tbody tr:hover{background:#f8fafc}
    #supRoot td{padding:10px 14px;vertical-align:middle}
    #supRoot .cid{font-family:'DM Mono',monospace;font-size:11px;color:#0f4c81;font-weight:500}
    #supRoot .cmut{font-size:11px;color:#5a6a7e}
    #supRoot .cdate{font-family:'DM Mono',monospace;font-size:11px;color:#5a6a7e;white-space:nowrap}
    #supRoot .pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:10.5px;font-weight:600;white-space:nowrap}
    #supRoot .pr{background:#fde8e8;color:#c62828} #supRoot .pg{background:#e8f5e9;color:#2e7d32}
    #supRoot .pa{background:#fff3e0;color:#e65100} #supRoot .pb{background:#e3f2fd;color:#0d47a1} #supRoot .pz{background:#f0f0f0;color:#616161}
    #supRoot .chk{display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;font-size:11px}
    #supRoot .cy{background:#e8f5e9;color:#2e7d32} #supRoot .cn{background:#fde8e8;color:#c62828} #supRoot .cp{background:#fff3e0;color:#e65100}
    #supRoot .btn-new{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:7px 16px;border-radius:6px;background:#0f4c81;color:#fff;border:none;cursor:pointer;display:inline-flex;align-items:center;gap:6px;transition:background .15s}
    #supRoot .btn-new:hover{background:#1a6bb5}
    #supRoot .load-row td{padding:24px;text-align:center;color:#5a6a7e;font-size:12px;font-family:'DM Mono',monospace}
    #supRoot .modal-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:9999;align-items:center;justify-content:center;padding:20px}
    #supRoot .modal-ov.open{display:flex}
    #supRoot .modal{background:#fff;border-radius:12px;width:100%;max-width:620px;max-height:85vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
    #supRoot .modal-hdr{padding:20px 24px 16px;border-bottom:1px solid #e0e6ed;display:flex;justify-content:space-between;align-items:flex-start;position:sticky;top:0;background:#fff;z-index:1}
    #supRoot .modal-title{font-size:15px;font-weight:700;color:#0f4c81}
    #supRoot .modal-sub{font-family:'DM Mono',monospace;font-size:10px;color:#5a6a7e;margin-top:2px}
    #supRoot .modal-x{background:none;border:none;font-size:22px;cursor:pointer;color:#5a6a7e;padding:0 4px;line-height:1}
    #supRoot .modal-body{padding:20px 24px}
    #supRoot .fg{margin-bottom:14px}
    #supRoot .fl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;margin-bottom:4px}
    #supRoot .fv{font-size:13px;color:#1a2332;line-height:1.5;padding:8px 10px;background:#f8fafc;border-radius:6px;border:1px solid #e0e6ed}
    #supRoot .fv.crit{background:#fde8e8;border-color:#ef9a9a;color:#c62828;font-size:12px}
    #supRoot .modal-ft{padding:16px 24px;border-top:1px solid #e0e6ed;display:flex;justify-content:flex-end;gap:8px}
    #supRoot .btn-sec{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:8px 18px;border-radius:6px;background:#f0f4f8;color:#1a2332;border:1px solid #e0e6ed;cursor:pointer}
    #supRoot .btn-pri{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:8px 18px;border-radius:6px;background:#0f4c81;color:#fff;border:none;cursor:pointer;text-decoration:none;display:inline-block}
    #supRoot .doc-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;margin-top:8px}
    #supRoot .doc-item{display:flex;align-items:center;gap:6px;font-size:12px;color:#1a2332}
  </style>

  <div class="ph">
    <div class="ph-top">
      <div>
        <div class="ph-eye">IMP9177 · 3H Pharmaceuticals LLC</div>
        <div class="ph-title">Supplier Registry &amp; CAPA Log</div>
        <div class="ph-sub" id="supSub">Loading live data...</div>
      </div>
      <div><div class="ph-ts" id="supTs">Fetching...</div></div>
    </div>
    <div class="kpi-strip">
      <div class="kpi-c"><div class="kpi-v" id="kST"><span class="kpi-load">—</span></div><div class="kpi-l">Suppliers</div></div>
      <div class="kpi-c"><div class="kpi-v green" id="kSQ"><span class="kpi-load">—</span></div><div class="kpi-l">Qualified</div></div>
      <div class="kpi-c"><div class="kpi-v red" id="kSS"><span class="kpi-load">—</span></div><div class="kpi-l">Suspended</div></div>
      <div class="kpi-c"><div class="kpi-v amber" id="kSP"><span class="kpi-load">—</span></div><div class="kpi-l">Pending Docs</div></div>
      <div class="kpi-c"><div class="kpi-v" id="kCA"><span class="kpi-load">—</span></div><div class="kpi-l">Open CAPAs</div></div>
      <div class="kpi-c"><div class="kpi-v red" id="kNC"><span class="kpi-load">—</span></div><div class="kpi-l">Missing CoA</div></div>
    </div>
  </div>

  <div class="alert" id="supAlert"><span>🚫</span><span id="supAlertTxt"></span></div>

  <div class="content">
    <div class="sec-hdr">
      <div class="sec-title">Supplier Registry</div>
      <div style="display:flex;align-items:center;gap:8px">
        <span class="sec-cnt" id="supCnt">—</span>
        <button class="btn-new" id="btnNewSup">+ New Supplier</button>
      </div>
    </div>
    <div class="fbar" id="supFbar">
      <button class="fbtn active" data-f="all">All</button>
      <button class="fbtn" data-f="qualified">✓ Qualified</button>
      <button class="fbtn" data-f="suspended">🚫 Suspended</button>
      <button class="fbtn" data-f="pending">⏳ Pending</button>
      <button class="fbtn" data-f="noCoA">⚠ No CoA</button>
      <input class="sbox" id="supSearch" placeholder="Search suppliers...">
    </div>
    <div class="sup-grid" id="supGrid"></div>
    <div class="tcard">
      <table>
        <thead><tr>
          <th data-sort="Title">Supplier</th><th>Commodity</th>
          <th data-sort="SupplierStatus">Status</th>
          <th>CoA</th><th>SDS</th><th>Allergen</th>
          <th data-sort="QualificationDate">Qual Date</th><th data-sort="RequalificationDue">Requal Due</th>
        </tr></thead>
        <tbody id="supBody"><tr class="load-row"><td colspan="8">Loading suppliers...</td></tr></tbody>
      </table>
    </div>

    <div class="sec-hdr">
      <div class="sec-title">CAPA Log</div>
      <div style="display:flex;align-items:center;gap:8px">
        <span class="sec-cnt" id="capaCnt">—</span>
        <button class="btn-new" id="btnNewCapa">+ New CAPA</button>
      </div>
    </div>
    <div class="tcard">
      <table>
        <thead><tr>
          <th>CAPA ID</th><th>Source</th><th data-sort-capa="DateOpened">Opened</th>
          <th>Initiated By</th><th data-sort-capa="CAPAStatus">Status</th>
          <th>Corrective Actions</th><th data-sort-capa="ClosureDate">Closed</th>
        </tr></thead>
        <tbody id="capaBody"><tr class="load-row"><td colspan="7">Loading CAPAs...</td></tr></tbody>
      </table>
    </div>
  </div>

  <div class="modal-ov" id="supModal">
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
      const [supRes, capaRes]: [SPHttpClientResponse, SPHttpClientResponse] = await Promise.all([
        this.context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('MRS_SupplierRegistry')/items?$select=Id,Title,Commodity,SupplierStatus,Purchaser,CountryOfOrigin,CoAOnFile,SDSOnFile,AllergenDeclOnFile,FoodSafetyCertOnFile,IdentityTestResult,QualificationDate,RequalificationDue,QualDeficiencies,SuspensionReason&$top=500&$orderby=Title`,
          SPHttpClient.configurations.v1
        ),
        this.context.spHttpClient.get(
          `${siteUrl}/_api/web/lists/getbytitle('MRS_CAPALog')/items?$select=Id,Title,CAPASource,DateOpened,InitiatedBy,CAPAStatus,CorrectiveActions,PreventiveActions,ClosureDate,ClosedBy,LinkedDocRefs&$top=500&$orderby=Title`,
          SPHttpClient.configurations.v1
        )
      ]);
      const [supData, capaData] = await Promise.all([supRes.json(), capaRes.json()]);
      this._sup = supData.value || [];
      this._capa = capaData.value || [];
    } catch (e) {
      console.error('SUP Forms: data load failed', e);
    }
    this._updateKPIs();
    this._renderSup();
    this._renderCapa();
    this._set('supTs', 'Refreshed ' + new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' }));
  }

  private _fmt(s: string): string {
    if (!s) return '—';
    try { return new Date(s).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' }); }
    catch { return s; }
  }

  private _pill(s: string): string {
    if (!s) return '<span class="pill pz">—</span>';
    const l = s.toLowerCase();
    if (l.includes('suspended') || l.includes('no coa') || l.includes('unqualified')) return `<span class="pill pr">${s}</span>`;
    if (l.includes('qualified') && !l.includes('un')) return `<span class="pill pg">${s}</span>`;
    if (l.includes('pending') || l.includes('active')) return `<span class="pill pa">${s}</span>`;
    return `<span class="pill pz">${s}</span>`;
  }

  private _capaPill(s: string): string {
    if (!s) return '<span class="pill pz">—</span>';
    if (s.toLowerCase().includes('closed')) return `<span class="pill pg">${s}</span>`;
    if (s.toLowerCase().includes('partial')) return `<span class="pill pa">${s}</span>`;
    return `<span class="pill pb">${s}</span>`;
  }

  private _chk(v: string): string {
    if (!v || v === 'No' || v === 'None') return '<span class="chk cn">✗</span>';
    if (v === 'Partial') return '<span class="chk cp">~</span>';
    if (v === 'Yes' || v === 'Pass') return '<span class="chk cy">✓</span>';
    return `<span class="chk cp">${v}</span>`;
  }

  private _updateKPIs(): void {
    const qual = this._sup.filter(s => (s.SupplierStatus || '').toLowerCase().includes('qualified') && !(s.SupplierStatus || '').toLowerCase().includes('un')).length;
    const sus = this._sup.filter(s => (s.SupplierStatus || '').toLowerCase().includes('suspended')).length;
    const pend = this._sup.filter(s => (s.SupplierStatus || '').toLowerCase().includes('pending')).length;
    const noCoA = this._sup.filter(s => s.CoAOnFile === 'No').length;
    const openCapa = this._capa.filter(c => !(c.CAPAStatus || '').toLowerCase().includes('closed')).length;
    this._set('kST', String(this._sup.length));
    this._set('kSQ', String(qual));
    this._set('kSS', String(sus));
    this._set('kSP', String(pend));
    this._set('kCA', String(openCapa));
    this._set('kNC', String(noCoA));
    this._set('supSub', `${this._sup.length} suppliers · ${qual} qualified · ${sus} suspended · ${noCoA} missing CoA`);
    if (sus > 0 || noCoA > 0) {
      const ids = this._sup.filter(s => s.CoAOnFile === 'No' || (s.SupplierStatus || '').toLowerCase().includes('suspended')).map(s => s.Title).join(', ');
      this._set('supAlertTxt', `Supplier alert: ${ids} — suspended or missing CoA`);
      const el = this.domElement.querySelector('#supAlert') as HTMLElement;
      if (el) el.classList.add('show');
    }
  }

  private _filteredSup(): any[] {
    let l = this._sup;
    if (this._supFilter === 'qualified') l = l.filter(s => (s.SupplierStatus || '').toLowerCase().includes('qualified') && !(s.SupplierStatus || '').toLowerCase().includes('un'));
    else if (this._supFilter === 'suspended') l = l.filter(s => (s.SupplierStatus || '').toLowerCase().includes('suspended'));
    else if (this._supFilter === 'pending') l = l.filter(s => (s.SupplierStatus || '').toLowerCase().includes('pending'));
    else if (this._supFilter === 'noCoA') l = l.filter(s => s.CoAOnFile === 'No');
    if (this._supSearch) {
      const q = this._supSearch.toLowerCase();
      l = l.filter(s => (s.Title || '').toLowerCase().includes(q) || (s.Commodity || '').toLowerCase().includes(q));
    }
    return l;
  }

  private _cardClass(s: any): string {
    const l = (s.SupplierStatus || '').toLowerCase();
    if (l.includes('suspended') || s.CoAOnFile === 'No') return 'sc crit';
    if (l.includes('pending') || l.includes('active')) return 'sc warn';
    return 'sc ok';
  }

  private _renderSup(): void {
    const l = this._filteredSup();
    this._set('supCnt', String(l.length));
    const grid = this.domElement.querySelector('#supGrid') as HTMLElement;
    if (grid) {
      grid.innerHTML = l.map(s => `<div class="${this._cardClass(s)}" data-sid="${s.Id}">
        <div class="sc-name">${s.Title}</div>
        <div class="sc-comm">${s.Commodity || '—'}</div>
        ${this._pill(s.SupplierStatus)}
      </div>`).join('');
    }
    const tbody = this.domElement.querySelector('#supBody') as HTMLElement;
    if (tbody) {
      tbody.innerHTML = l.length ? l.map(s => `<tr data-sid="${s.Id}">
        <td><span class="cid">${s.Title}</span><br><span class="cmut">${s.CountryOfOrigin || '—'}</span></td>
        <td><span class="cmut">${s.Commodity || '—'}</span></td>
        <td>${this._pill(s.SupplierStatus)}</td>
        <td>${this._chk(s.CoAOnFile)}</td>
        <td>${this._chk(s.SDSOnFile)}</td>
        <td>${this._chk(s.AllergenDeclOnFile)}</td>
        <td><span class="cdate">${this._fmt(s.QualificationDate)}</span></td>
        <td><span class="cdate">${this._fmt(s.RequalificationDue)}</span></td>
      </tr>`).join('') : '<tr><td colspan="8" style="text-align:center;padding:32px;color:#aaa">No suppliers match filter</td></tr>';
    }
  }

  private _renderCapa(): void {
    this._set('capaCnt', String(this._capa.length));
    const tbody = this.domElement.querySelector('#capaBody') as HTMLElement;
    if (!tbody) return;
    tbody.innerHTML = this._capa.length ? this._capa.map(c => `<tr data-cid="${c.Id}">
      <td><span class="cid">${c.Title}</span></td>
      <td><span class="cmut">${c.CAPASource || '—'}</span></td>
      <td><span class="cdate">${this._fmt(c.DateOpened)}</span></td>
      <td><span class="cmut">${c.InitiatedBy || '—'}</span></td>
      <td>${this._capaPill(c.CAPAStatus)}</td>
      <td><span style="font-size:11.5px;display:block;max-width:250px">${(c.CorrectiveActions || '—').substring(0, 100)}${(c.CorrectiveActions || '').length > 100 ? '…' : ''}</span></td>
      <td><span class="cdate">${this._fmt(c.ClosureDate)}</span></td>
    </tr>`).join('') : '<tr><td colspan="7" style="text-align:center;padding:32px;color:#aaa">No CAPAs on file</td></tr>';
  }

  private _showSupModal(id: number): void {
    const s = this._sup.find(x => x.Id === id);
    if (!s) return;
    const crit = s.CoAOnFile === 'No' || (s.SupplierStatus || '').toLowerCase().includes('suspended');
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    this._set('mTitle', s.Title);
    this._set('mSub', `Supplier Registry · ${s.Commodity}`);
    this._set('mBody', `
      <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(s.SupplierStatus)}</div></div>
      <div class="fg"><div class="fl">Commodity</div><div class="fv">${s.Commodity || '—'}</div></div>
      <div class="fg"><div class="fl">Country of Origin</div><div class="fv">${s.CountryOfOrigin || '—'}</div></div>
      <div class="fg"><div class="fl">Purchaser</div><div class="fv">${s.Purchaser || '—'}</div></div>
      <div class="fg"><div class="fl">Documentation</div>
        <div class="doc-grid">
          <div class="doc-item">${this._chk(s.CoAOnFile)} CoA</div>
          <div class="doc-item">${this._chk(s.SDSOnFile)} SDS</div>
          <div class="doc-item">${this._chk(s.AllergenDeclOnFile)} Allergen</div>
          <div class="doc-item">${this._chk(s.FoodSafetyCertOnFile)} Food Safety</div>
          <div class="doc-item">${this._chk(s.IdentityTestResult)} Identity</div>
        </div>
      </div>
      <div class="fg"><div class="fl">Qual Date</div><div class="fv">${this._fmt(s.QualificationDate)}</div></div>
      <div class="fg"><div class="fl">Requal Due</div><div class="fv">${this._fmt(s.RequalificationDue)}</div></div>
      ${s.QualDeficiencies ? `<div class="fg"><div class="fl">Deficiencies</div><div class="fv ${crit ? 'crit' : ''}">${s.QualDeficiencies}</div></div>` : ''}
      ${s.SuspensionReason ? `<div class="fg"><div class="fl">Suspension Reason</div><div class="fv crit">${s.SuspensionReason}</div></div>` : ''}`);
    const editLink = this.domElement.querySelector('#mEdit') as HTMLAnchorElement;
    if (editLink) editLink.href = `${siteUrl}/Lists/MRS_SupplierRegistry/EditForm.aspx?ID=${id}`;
    (this.domElement.querySelector('#supModal') as HTMLElement)?.classList.add('open');
  }

  private _showCapaModal(id: number): void {
    const c = this._capa.find(x => x.Id === id);
    if (!c) return;
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    this._set('mTitle', c.Title);
    this._set('mSub', 'CAPA Log');
    this._set('mBody', `
      <div class="fg"><div class="fl">Source</div><div class="fv">${c.CAPASource || '—'}</div></div>
      <div class="fg"><div class="fl">Status</div><div class="fv">${this._capaPill(c.CAPAStatus)}</div></div>
      <div class="fg"><div class="fl">Opened</div><div class="fv">${this._fmt(c.DateOpened)}</div></div>
      <div class="fg"><div class="fl">Initiated By</div><div class="fv">${c.InitiatedBy || '—'}</div></div>
      <div class="fg"><div class="fl">Corrective Actions</div><div class="fv" style="font-size:12px">${c.CorrectiveActions || '—'}</div></div>
      <div class="fg"><div class="fl">Preventive Actions</div><div class="fv" style="font-size:12px">${c.PreventiveActions || '—'}</div></div>
      <div class="fg"><div class="fl">Linked Docs</div><div class="fv">${c.LinkedDocRefs || '—'}</div></div>
      <div class="fg"><div class="fl">Closure Date</div><div class="fv">${this._fmt(c.ClosureDate)}</div></div>`);
    const editLink = this.domElement.querySelector('#mEdit') as HTMLAnchorElement;
    if (editLink) editLink.href = `${siteUrl}/Lists/MRS_CAPALog/EditForm.aspx?ID=${id}`;
    (this.domElement.querySelector('#supModal') as HTMLElement)?.classList.add('open');
  }

  private _bindEvents(): void {
    const root = this.domElement;
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    root.querySelector('#supFbar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-f]') as HTMLElement;
      if (!btn) return;
      root.querySelectorAll('#supFbar .fbtn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active');
      this._supFilter = btn.dataset.f || 'all';
      this._renderSup();
    });
    root.querySelector('#supSearch')?.addEventListener('input', (e) => {
      this._supSearch = (e.target as HTMLInputElement).value;
      this._renderSup();
    });
    root.querySelectorAll('[data-sort]').forEach(th => {
      th.addEventListener('click', () => {
        const f = (th as HTMLElement).dataset.sort || '';
        this._sup.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || '')));
        this._renderSup();
      });
    });
    root.querySelectorAll('[data-sort-capa]').forEach(th => {
      th.addEventListener('click', () => {
        const f = (th as HTMLElement).dataset.sortCapa || '';
        this._capa.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || '')));
        this._renderCapa();
      });
    });
    root.querySelector('#supGrid')?.addEventListener('click', (e) => {
      const card = (e.target as HTMLElement).closest('[data-sid]') as HTMLElement;
      if (card) this._showSupModal(parseInt(card.dataset.sid || '0'));
    });
    root.querySelector('#supBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('[data-sid]') as HTMLElement;
      if (row) this._showSupModal(parseInt(row.dataset.sid || '0'));
    });
    root.querySelector('#capaBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('[data-cid]') as HTMLElement;
      if (row) this._showCapaModal(parseInt(row.dataset.cid || '0'));
    });
    const closeModal = () => (root.querySelector('#supModal') as HTMLElement)?.classList.remove('open');
    root.querySelector('#mClose')?.addEventListener('click', closeModal);
    root.querySelector('#mDismiss')?.addEventListener('click', closeModal);
    root.querySelector('#supModal')?.addEventListener('click', (e) => {
      if ((e.target as HTMLElement).id === 'supModal') closeModal();
    });
    root.querySelector('#btnNewSup')?.addEventListener('click', () => window.open(`${siteUrl}/Lists/MRS_SupplierRegistry/NewForm.aspx`, '_blank'));
    root.querySelector('#btnNewCapa')?.addEventListener('click', () => window.open(`${siteUrl}/Lists/MRS_CAPALog/NewForm.aspx`, '_blank'));
  }

  private _set(id: string, html: string): void {
    const el = this.domElement.querySelector(`#${id}`) as HTMLElement;
    if (el) el.innerHTML = html;
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
}
