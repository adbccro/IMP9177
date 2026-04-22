import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IPmFormsWebPartProps {}

export default class PmFormsWebPart extends BaseClientSideWebPart<IPmFormsWebPartProps> {
  private _ms: any[] = [];
  private _oi: any[] = [];
  private _del: any[] = [];
  private _bud: any[] = [];
  private _msFilter: string = 'all';
  private _oiFilter: string = 'all';
  private _oiSearch: string = '';

  public render(): void {
    this.domElement.innerHTML = this._getShell();
    this._loadData();
    this._bindEvents();
  }

  private _getShell(): string {
    return `
<div id="pmRoot" style="font-family:'DM Sans',Segoe UI,sans-serif;background:#f4f6f9;min-height:100vh;color:#1a2332">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600;700&display=swap');
    #pmRoot *{box-sizing:border-box;margin:0;padding:0}
    #pmRoot .ph{background:linear-gradient(135deg,#0a3259 0%,#0f4c81 60%,#1a6bb5 100%);padding:28px 32px 0;overflow:hidden}
    #pmRoot .ph-top{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;position:relative;z-index:1}
    #pmRoot .ph-eye{font-family:'DM Mono',monospace;font-size:10px;letter-spacing:2px;color:rgba(255,255,255,.5);text-transform:uppercase;margin-bottom:6px}
    #pmRoot .ph-title{font-size:22px;font-weight:700;color:#fff}
    #pmRoot .ph-sub{font-size:13px;color:rgba(255,255,255,.65);margin-top:4px}
    #pmRoot .ph-ts{font-family:'DM Mono',monospace;font-size:10px;color:rgba(255,255,255,.4);text-align:right;white-space:nowrap;margin-top:4px}
    #pmRoot .cd{font-family:'DM Mono',monospace;font-size:12px;padding:4px 12px;border-radius:20px;margin-top:8px;display:inline-block}
    #pmRoot .cd-ok{background:#e8f5e9;color:#2e7d32} #pmRoot .cd-warn{background:#fff3e0;color:#e65100} #pmRoot .cd-danger{background:#fde8e8;color:#c62828}
    #pmRoot .kpi-strip{background:#0a3259;border-top:1px solid rgba(255,255,255,.08);padding:0 32px;display:flex;overflow-x:auto}
    #pmRoot .kpi-c{padding:14px 24px 14px 0;margin-right:24px;border-right:1px solid rgba(255,255,255,.08);min-width:110px;flex-shrink:0}
    #pmRoot .kpi-c:last-child{border-right:none}
    #pmRoot .kpi-v{font-family:'DM Mono',monospace;font-size:24px;font-weight:500;color:#fff;line-height:1}
    #pmRoot .kpi-v.red{color:#ef9a9a} #pmRoot .kpi-v.green{color:#a5d6a7} #pmRoot .kpi-v.amber{color:#ffcc80}
    #pmRoot .kpi-l{font-size:10px;color:rgba(255,255,255,.45);letter-spacing:.8px;text-transform:uppercase;margin-top:4px}
    #pmRoot .kpi-load{animation:pmPulse 1.2s ease-in-out infinite;color:rgba(255,255,255,.2)}
    @keyframes pmPulse{0%,100%{opacity:.3}50%{opacity:1}}
    #pmRoot .prog-hdr{padding:0 32px 20px;position:relative;z-index:1}
    #pmRoot .prog-lbl{font-size:11px;color:rgba(255,255,255,.55);margin-bottom:8px;display:flex;justify-content:space-between}
    #pmRoot .prog-track{height:8px;background:rgba(255,255,255,.15);border-radius:4px;overflow:hidden}
    #pmRoot .prog-fill{height:100%;background:linear-gradient(90deg,#a5d6a7,#66bb6a);border-radius:4px;transition:width .8s ease}
    #pmRoot .prog-meta{display:flex;gap:12px;margin-top:8px;flex-wrap:wrap}
    #pmRoot .chip{font-family:'DM Mono',monospace;font-size:10px;color:rgba(255,255,255,.5);padding:3px 8px;background:rgba(255,255,255,.08);border-radius:12px}
    #pmRoot .alert{margin:16px 32px 0;padding:12px 16px;border-radius:8px;border-left:4px solid #c62828;background:#fde8e8;color:#c62828;display:none;align-items:center;gap:10px;font-size:13px;font-weight:500}
    #pmRoot .alert.show{display:flex}
    #pmRoot .content{padding:20px 32px 32px}
    #pmRoot .sec-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;margin-top:24px}
    #pmRoot .sec-title{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#0f4c81;display:flex;align-items:center;gap:8px}
    #pmRoot .sec-title::before{content:'';width:3px;height:14px;background:#0f4c81;border-radius:2px;display:block}
    #pmRoot .sec-cnt{font-family:'DM Mono',monospace;font-size:11px;color:#5a6a7e;background:#e0e6ed;padding:2px 8px;border-radius:10px}
    #pmRoot .bud-row{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:12px;margin-bottom:20px}
    #pmRoot .bc{background:#fff;border-radius:10px;border:1px solid #e0e6ed;padding:16px;box-shadow:0 1px 3px rgba(0,0,0,.04)}
    #pmRoot .bc-lbl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;margin-bottom:4px}
    #pmRoot .bc-val{font-family:'DM Mono',monospace;font-size:22px;font-weight:500;color:#0f4c81}
    #pmRoot .bc-sub{font-size:11px;color:#5a6a7e;margin-top:2px}
    #pmRoot .fbar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center}
    #pmRoot .fbtn{font-family:'DM Sans',sans-serif;font-size:11px;font-weight:600;padding:5px 12px;border-radius:20px;border:1px solid #e0e6ed;background:#fff;color:#5a6a7e;cursor:pointer;transition:all .15s}
    #pmRoot .fbtn:hover,#pmRoot .fbtn.active{background:#0f4c81;color:#fff;border-color:#0f4c81}
    #pmRoot .sbox{font-family:'DM Sans',sans-serif;font-size:12px;padding:5px 12px;border-radius:20px;border:1px solid #e0e6ed;background:#fff;outline:none;min-width:180px;margin-left:auto}
    #pmRoot .sbox:focus{border-color:#0f4c81}
    #pmRoot .tcard{background:#fff;border-radius:10px;border:1px solid #e0e6ed;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.05)}
    #pmRoot table{width:100%;border-collapse:collapse;font-size:12.5px}
    #pmRoot thead th{background:#f8fafc;padding:10px 14px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;border-bottom:1px solid #e0e6ed;white-space:nowrap;cursor:pointer}
    #pmRoot thead th:hover{color:#0f4c81}
    #pmRoot tbody tr{border-bottom:1px solid #f0f4f8;transition:background .1s;cursor:pointer}
    #pmRoot tbody tr:last-child{border-bottom:none}
    #pmRoot tbody tr:hover{background:#f8fafc}
    #pmRoot td{padding:10px 14px;vertical-align:middle}
    #pmRoot .cid{font-family:'DM Mono',monospace;font-size:11px;color:#0f4c81;font-weight:500}
    #pmRoot .cmut{font-size:11px;color:#5a6a7e}
    #pmRoot .cdate{font-family:'DM Mono',monospace;font-size:11px;color:#5a6a7e;white-space:nowrap}
    #pmRoot .pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:10.5px;font-weight:600;white-space:nowrap}
    #pmRoot .pr{background:#fde8e8;color:#c62828} #pmRoot .pg{background:#e8f5e9;color:#2e7d32}
    #pmRoot .pa{background:#fff3e0;color:#e65100} #pmRoot .pb{background:#e3f2fd;color:#0d47a1} #pmRoot .pz{background:#f0f0f0;color:#616161}
    #pmRoot .pri{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700}
    #pmRoot .pc{background:#fde8e8;color:#c62828} #pmRoot .ph2{background:#fff3e0;color:#e65100} #pmRoot .pm{background:#e3f2fd;color:#0d47a1}
    #pmRoot .phb{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700}
    #pmRoot .ph0{background:#f3e5f5;color:#4a148c} #pmRoot .ph1{background:#e3f2fd;color:#0d47a1}
    #pmRoot .ph2a{background:#e8f5e9;color:#2e7d32} #pmRoot .ph2b{background:#fff8e1;color:#f57f17}
    #pmRoot .ph2c{background:#fce4ec;color:#880e4f} #pmRoot .ph2d{background:#fff3e0;color:#e65100}
    #pmRoot .pbar{display:flex;align-items:center;gap:8px}
    #pmRoot .ptrack{width:80px;height:6px;background:#e0e0e0;border-radius:3px;overflow:hidden;flex-shrink:0}
    #pmRoot .pfill{height:100%;border-radius:3px}
    #pmRoot .pfill-done{background:#2e7d32} #pmRoot .pfill-act{background:#0f4c81} #pmRoot .pfill-none{background:#e0e0e0}
    #pmRoot .ppct{font-family:'DM Mono',monospace;font-size:10.5px;color:#5a6a7e}
    #pmRoot .btn-new{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:7px 16px;border-radius:6px;background:#0f4c81;color:#fff;border:none;cursor:pointer;display:inline-flex;align-items:center;gap:6px;transition:background .15s}
    #pmRoot .btn-new:hover{background:#1a6bb5}
    #pmRoot .load-row td{padding:24px;text-align:center;color:#5a6a7e;font-size:12px;font-family:'DM Mono',monospace}
    #pmRoot .modal-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:9999;align-items:center;justify-content:center;padding:20px}
    #pmRoot .modal-ov.open{display:flex}
    #pmRoot .modal{background:#fff;border-radius:12px;width:100%;max-width:620px;max-height:85vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
    #pmRoot .modal-hdr{padding:20px 24px 16px;border-bottom:1px solid #e0e6ed;display:flex;justify-content:space-between;align-items:flex-start;position:sticky;top:0;background:#fff;z-index:1}
    #pmRoot .modal-title{font-size:15px;font-weight:700;color:#0f4c81}
    #pmRoot .modal-sub{font-family:'DM Mono',monospace;font-size:10px;color:#5a6a7e;margin-top:2px}
    #pmRoot .modal-x{background:none;border:none;font-size:22px;cursor:pointer;color:#5a6a7e;padding:0 4px;line-height:1}
    #pmRoot .modal-body{padding:20px 24px}
    #pmRoot .fg{margin-bottom:14px}
    #pmRoot .fl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;margin-bottom:4px}
    #pmRoot .fv{font-size:13px;color:#1a2332;line-height:1.5;padding:8px 10px;background:#f8fafc;border-radius:6px;border:1px solid #e0e6ed}
    #pmRoot .modal-ft{padding:16px 24px;border-top:1px solid #e0e6ed;display:flex;justify-content:flex-end;gap:8px}
    #pmRoot .btn-sec{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:8px 18px;border-radius:6px;background:#f0f4f8;color:#1a2332;border:1px solid #e0e6ed;cursor:pointer}
    #pmRoot .btn-pri{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:8px 18px;border-radius:6px;background:#0f4c81;color:#fff;border:none;cursor:pointer;text-decoration:none;display:inline-block}
  </style>

  <div class="ph">
    <div class="ph-top">
      <div>
        <div class="ph-eye">IMP9177 · 3H Pharmaceuticals LLC</div>
        <div class="ph-title">Project Management</div>
        <div class="ph-sub" id="pmSub">Loading live data...</div>
      </div>
      <div>
        <div class="ph-ts" id="pmTs">Fetching...</div>
        <div id="cdDisplay"></div>
      </div>
    </div>
    <div class="kpi-strip">
      <div class="kpi-c"><div class="kpi-v" id="kPH"><span class="kpi-load">—</span></div><div class="kpi-l">Phase</div></div>
      <div class="kpi-c"><div class="kpi-v amber" id="kPCT"><span class="kpi-load">—</span></div><div class="kpi-l">Complete</div></div>
      <div class="kpi-c"><div class="kpi-v" id="kDAYS"><span class="kpi-load">—</span></div><div class="kpi-l">Days to M-1</div></div>
      <div class="kpi-c"><div class="kpi-v red" id="kOI"><span class="kpi-load">—</span></div><div class="kpi-l">Open Items</div></div>
      <div class="kpi-c"><div class="kpi-v green" id="kREM"><span class="kpi-load">—</span></div><div class="kpi-l">Budget Left</div></div>
      <div class="kpi-c"><div class="kpi-v" id="kSPT"><span class="kpi-load">—</span></div><div class="kpi-l">Spent</div></div>
    </div>
    <div class="prog-hdr">
      <div class="prog-lbl"><span id="progLabel">Phase 2A — Loading...</span><span id="progPct">—%</span></div>
      <div class="prog-track"><div class="prog-fill" id="progFill" style="width:0%"></div></div>
      <div class="prog-meta" id="progMeta"></div>
    </div>
  </div>

  <div class="alert" id="pmAlert"><span>⚠️</span><span id="pmAlertTxt"></span></div>

  <div class="content">
    <div class="sec-hdr" style="margin-top:8px">
      <div class="sec-title">Budget</div>
      <button class="btn-new" id="btnNewBud">+ Entry</button>
    </div>
    <div class="bud-row" id="budCards"></div>

    <div class="sec-hdr">
      <div class="sec-title">Milestone Tracker</div>
      <div style="display:flex;align-items:center;gap:8px"><span class="sec-cnt" id="msCnt">—</span><button class="btn-new" id="btnNewMS">+ Milestone</button></div>
    </div>
    <div class="fbar" id="msFbar">
      <button class="fbtn active" data-mf="all">All</button>
      <button class="fbtn" data-mf="active">Active</button>
      <button class="fbtn" data-mf="complete">Complete</button>
      <button class="fbtn" data-mf="at-risk">⚠ At Risk</button>
      <button class="fbtn" data-mf="Phase 2A">Phase 2A</button>
      <button class="fbtn" data-mf="Phase 2B">Phase 2B</button>
    </div>
    <div class="tcard">
      <table><thead><tr>
        <th data-sms="Title">ID</th><th>Phase</th><th>Milestone</th>
        <th data-sms="TargetEnd">Target</th><th>Owner</th>
        <th data-sms="MilestoneStatus">Status</th><th data-sms="PercentDone">Progress</th>
      </tr></thead>
      <tbody id="msBody"><tr class="load-row"><td colspan="7">Loading milestones...</td></tr></tbody></table>
    </div>

    <div class="sec-hdr">
      <div class="sec-title">Open Items</div>
      <div style="display:flex;align-items:center;gap:8px"><span class="sec-cnt" id="oiCnt">—</span><button class="btn-new" id="btnNewOI">+ Item</button></div>
    </div>
    <div class="fbar" id="oiFbar">
      <button class="fbtn active" data-of="all">All</button>
      <button class="fbtn" data-of="critical">🔴 Critical</button>
      <button class="fbtn" data-of="adb">ADB Owned</button>
      <button class="fbtn" data-of="3h">3H Owned</button>
      <input class="sbox" id="oiSearch" placeholder="Search items...">
    </div>
    <div class="tcard">
      <table><thead><tr>
        <th>Ref</th><th>Description</th><th>Owner</th>
        <th data-soi="OIPriority">Priority</th><th data-soi="OITargetDate">Target</th><th data-soi="OIStatus">Status</th>
      </tr></thead>
      <tbody id="oiBody"><tr class="load-row"><td colspan="6">Loading open items...</td></tr></tbody></table>
    </div>

    <div class="sec-hdr">
      <div class="sec-title">Deliverables</div>
      <div style="display:flex;align-items:center;gap:8px"><span class="sec-cnt" id="delCnt">—</span><button class="btn-new" id="btnNewDel">+ Deliverable</button></div>
    </div>
    <div class="tcard">
      <table><thead><tr>
        <th>Deliverable</th><th data-sdel="DeliveredDate">Delivered</th>
        <th data-sdel="FeedbackDue">Feedback Due</th><th data-sdel="SignOffDate">Signed Off</th>
        <th data-sdel="DelStatus">Status</th><th>Notes</th>
      </tr></thead>
      <tbody id="delBody"><tr class="load-row"><td colspan="6">Loading deliverables...</td></tr></tbody></table>
    </div>
  </div>

  <div class="modal-ov" id="pmModal">
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
    const site = this.context.pageContext.web.absoluteUrl;
    try {
      const [msR, oiR, delR, budR]: SPHttpClientResponse[] = await Promise.all([
        this.context.spHttpClient.get(`${site}/_api/web/lists/getbytitle('PM_Milestones')/items?$select=Id,Title,Phase,MilestoneDesc,StartDate,TargetEnd,MilestoneOwner,MilestoneStatus,PercentDone,MilestoneNotes&$top=500&$orderby=Title`, SPHttpClient.configurations.v1),
        this.context.spHttpClient.get(`${site}/_api/web/lists/getbytitle('PM_OpenItems')/items?$select=Id,Title,OIDescription,OIOwner,OIPriority,OITargetDate,OIStatus&$top=500&$orderby=Title`, SPHttpClient.configurations.v1),
        this.context.spHttpClient.get(`${site}/_api/web/lists/getbytitle('PM_DocumentDeliverables')/items?$select=Id,Title,DeliveredDate,FeedbackDue,SignOffDate,DelStatus,DelNotes&$top=500&$orderby=Title`, SPHttpClient.configurations.v1),
        this.context.spHttpClient.get(`${site}/_api/web/lists/getbytitle('PM_Budget')/items?$select=Id,Title,EntryType,Reference,EntryDate,Amount,BudgetNotes&$top=500&$orderby=Title`, SPHttpClient.configurations.v1)
      ]);
      const [msD, oiD, delD, budD] = await Promise.all([msR.json(), oiR.json(), delR.json(), budR.json()]);
      this._ms = msD.value || []; this._oi = oiD.value || []; this._del = delD.value || []; this._bud = budD.value || [];
    } catch (e) { console.error('PM Forms: load failed', e); }
    this._updateKPIs(); this._renderBudget(); this._renderMS(); this._renderOI(); this._renderDel();
    this._set('pmTs', 'Refreshed ' + new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' }));
  }

  private _fmt(s: string): string {
    if (!s) return '—';
    try { return new Date(s).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' }); } catch { return s; }
  }
  private _money(n: number): string { if (!n && n !== 0) return '—'; return '$' + Number(n).toLocaleString(); }

  private _pill(s: string): string {
    if (!s) return '<span class="pill pz">—</span>';
    const l = s.toLowerCase();
    if (l.includes('past due') || l.includes('critical') || l.includes('overdue') || l.includes('blocking')) return `<span class="pill pr">${s}</span>`;
    if (l.includes('complete') || l.includes('closed') || l.includes('issued apr')) return `<span class="pill pg">${s}</span>`;
    if (l.includes('progress') || l.includes('partial') || l.includes('pending feedback')) return `<span class="pill pa">${s}</span>`;
    if (l.includes('pending') || l.includes('open') || l.includes('not started') || l.includes('tbd')) return `<span class="pill pb">${s}</span>`;
    return `<span class="pill pz">${s}</span>`;
  }

  private _priPill(s: string): string {
    if (!s) return '';
    const l = s.toLowerCase();
    if (l === 'critical') return `<span class="pri pc">${s}</span>`;
    if (l === 'high') return `<span class="pri ph2">${s}</span>`;
    return `<span class="pri pm">${s}</span>`;
  }

  private _phaseBadge(p: string): string {
    const m: { [k: string]: string } = { 'Phase 0': 'ph0', 'Phase 1': 'ph1', 'Phase 2A': 'ph2a', 'Phase 2B': 'ph2b', 'Phase 2C': 'ph2c', 'Phase 2D': 'ph2d' };
    return `<span class="phb ${m[p] || ''}">${p}</span>`;
  }

  private _progBar(pct: number): string {
    const cls = pct >= 100 ? 'pfill-done' : pct > 0 ? 'pfill-act' : 'pfill-none';
    return `<div class="pbar"><div class="ptrack"><div class="pfill ${cls}" style="width:${pct}%"></div></div><span class="ppct">${pct}%</span></div>`;
  }

  private _daysUntil(s: string): number | null {
    if (!s) return null;
    try { return Math.round((new Date(s).getTime() - Date.now()) / (1000 * 60 * 60 * 24)); } catch { return null; }
  }

  private _updateKPIs(): void {
    const m15 = this._ms.find(m => m.Title === 'M-15');
    const pct = m15 ? parseInt(m15.PercentDone) || 0 : 0;
    const phase = m15 ? m15.Phase || 'Phase 2A' : 'Phase 2A';
    const days = m15 ? this._daysUntil(m15.TargetEnd) : null;
    const openOI = this._oi.filter(o => !(o.OIStatus || '').toLowerCase().includes('complete')).length;
    const total = this._bud.find(b => b.Title === 'Budget-Total');
    const remaining = this._bud.find(b => b.Title === 'Budget-Remaining');
    const inv = this._bud.find(b => b.Title === 'INV-IMP9177I002');

    this._set('kPH', phase); this._set('kPCT', pct + '%');
    if (days !== null) {
      this._set('kDAYS', String(days));
      const dEl = this.domElement.querySelector('#kDAYS') as HTMLElement;
      if (dEl) dEl.className = `kpi-v ${days > 7 ? 'green' : days > 3 ? 'amber' : 'red'}`;
      const cdCls = days > 7 ? 'cd-ok' : days > 3 ? 'cd-warn' : 'cd-danger';
      this._set('cdDisplay', `<span class="cd ${cdCls}">⏱ ${days} days to M-1</span>`);
    }
    this._set('kOI', String(openOI));
    this._set('kREM', remaining ? this._money(remaining.Amount) : '—');
    this._set('kSPT', inv ? this._money(inv.Amount) : '—');
    this._set('pmSub', `${phase} · ${pct}% complete · ${openOI} open items${days !== null ? ` · ${days} days to Milestone 1` : ''}`);
    (this.domElement.querySelector('#progFill') as HTMLElement)?.setAttribute('style', `width:${pct}%;height:100%;background:linear-gradient(90deg,#a5d6a7,#66bb6a);border-radius:4px;transition:width .8s ease`);
    this._set('progLabel', `${phase} — GMP Implementation`);
    this._set('progPct', pct + '%');
    this._set('progMeta', `<span class="chip">M-1 Target: ${m15 ? this._fmt(m15.TargetEnd) : '—'}</span><span class="chip">Budget: ${total ? this._money(total.Amount) : '—'}</span><span class="chip">Remaining: ${remaining ? this._money(remaining.Amount) : '—'}</span>`);
    const pastDueMS = this._ms.filter(m => (m.MilestoneStatus || '').toLowerCase().includes('past due')).length;
    if (pastDueMS > 0) {
      this._set('pmAlertTxt', `${pastDueMS} milestone(s) past due — immediate attention required`);
      (this.domElement.querySelector('#pmAlert') as HTMLElement)?.classList.add('show');
    }
  }

  private _renderBudget(): void {
    const total = this._bud.find(b => b.Title === 'Budget-Total');
    const remaining = this._bud.find(b => b.Title === 'Budget-Remaining');
    const inv = this._bud.find(b => b.Title === 'INV-IMP9177I002');
    const co = this._bud.find(b => b.Title === 'PSO-CO-001');
    const tot = total ? total.Amount : 25000;
    const spent = inv ? inv.Amount : 0;
    const spentPct = Math.round((spent / tot) * 100);
    this._set('budCards', `
      <div class="bc"><div class="bc-lbl">Total Budget</div><div class="bc-val">${this._money(tot)}</div><div class="bc-sub">PSO V7 authorized</div></div>
      <div class="bc"><div class="bc-lbl">Spent</div><div class="bc-val" style="color:#e65100">${this._money(spent)}</div><div class="bc-sub">${spentPct}% of budget</div></div>
      <div class="bc"><div class="bc-lbl">Remaining</div><div class="bc-val" style="color:#2e7d32">${remaining ? this._money(remaining.Amount) : '—'}</div><div class="bc-sub">Phase 2 intact</div></div>
      <div class="bc"><div class="bc-lbl">Change Order</div><div class="bc-val" style="font-size:14px;color:#e65100">PSO-CO-001</div><div class="bc-sub">${co ? 'Pending execution' : 'Not on file'}</div></div>`);
  }

  private _filteredMS(): any[] {
    let l = this._ms;
    if (this._msFilter === 'active') l = l.filter(m => !['complete', 'not started'].some(x => (m.MilestoneStatus || '').toLowerCase().includes(x)));
    else if (this._msFilter === 'complete') l = l.filter(m => (m.MilestoneStatus || '').toLowerCase().includes('complete'));
    else if (this._msFilter === 'at-risk') l = l.filter(m => (m.MilestoneStatus || '').toLowerCase().includes('past due'));
    else if (this._msFilter.startsWith('Phase')) l = l.filter(m => m.Phase === this._msFilter);
    return l;
  }

  private _renderMS(): void {
    const l = this._filteredMS();
    this._set('msCnt', String(l.length));
    const tb = this.domElement.querySelector('#msBody') as HTMLElement;
    if (!tb) return;
    tb.innerHTML = l.length ? l.map(m => `<tr data-msid="${m.Id}">
      <td><span class="cid">${m.Title}</span></td>
      <td>${this._phaseBadge(m.Phase)}</td>
      <td><span style="font-size:12px;display:block;max-width:220px">${m.MilestoneDesc || '—'}</span></td>
      <td><span class="cdate">${this._fmt(m.TargetEnd)}</span></td>
      <td><span class="cmut">${m.MilestoneOwner || '—'}</span></td>
      <td>${this._pill(m.MilestoneStatus)}</td>
      <td>${this._progBar(parseInt(m.PercentDone) || 0)}</td>
    </tr>`).join('') : '<tr><td colspan="7" style="text-align:center;padding:32px;color:#aaa">No milestones match filter</td></tr>';
  }

  private _filteredOI(): any[] {
    let l = this._oi;
    if (this._oiFilter === 'critical') l = l.filter(o => (o.OIPriority || '').toLowerCase() === 'critical');
    else if (this._oiFilter === 'adb') l = l.filter(o => (o.OIOwner || '').toLowerCase().includes('adb') || (o.OIOwner || '').toLowerCase().includes('andre'));
    else if (this._oiFilter === '3h') l = l.filter(o => ['3h', 'tina', 'cindy', 'kj'].some(x => (o.OIOwner || '').toLowerCase().includes(x)));
    if (this._oiSearch) { const q = this._oiSearch.toLowerCase(); l = l.filter(o => (o.Title || '').toLowerCase().includes(q) || (o.OIDescription || '').toLowerCase().includes(q)); }
    return l;
  }

  private _renderOI(): void {
    const l = this._filteredOI();
    this._set('oiCnt', String(l.length));
    const tb = this.domElement.querySelector('#oiBody') as HTMLElement;
    if (!tb) return;
    tb.innerHTML = l.length ? l.map(o => `<tr data-oiid="${o.Id}">
      <td><span class="cid">${o.Title}</span></td>
      <td><span style="font-size:12px;display:block;max-width:260px">${(o.OIDescription || '—').substring(0, 100)}${(o.OIDescription || '').length > 100 ? '…' : ''}</span></td>
      <td><span class="cmut">${o.OIOwner || '—'}</span></td>
      <td>${this._priPill(o.OIPriority)}</td>
      <td><span class="cdate">${this._fmt(o.OITargetDate)}</span></td>
      <td>${this._pill(o.OIStatus)}</td>
    </tr>`).join('') : '<tr><td colspan="6" style="text-align:center;padding:32px;color:#aaa">No open items match filter</td></tr>';
  }

  private _renderDel(): void {
    this._set('delCnt', String(this._del.length));
    const tb = this.domElement.querySelector('#delBody') as HTMLElement;
    if (!tb) return;
    tb.innerHTML = this._del.length ? this._del.map(d => `<tr data-delid="${d.Id}">
      <td><span style="font-size:12px">${d.Title || '—'}</span></td>
      <td><span class="cdate">${this._fmt(d.DeliveredDate)}</span></td>
      <td><span class="cdate">${this._fmt(d.FeedbackDue)}</span></td>
      <td><span class="cdate">${this._fmt(d.SignOffDate)}</span></td>
      <td>${this._pill(d.DelStatus)}</td>
      <td><span class="cmut" style="font-size:11.5px">${(d.DelNotes || '—').substring(0, 80)}${(d.DelNotes || '').length > 80 ? '…' : ''}</span></td>
    </tr>`).join('') : '<tr><td colspan="6" style="text-align:center;padding:32px;color:#aaa">No deliverables on file</td></tr>';
  }

  private _openModal(title: string, sub: string, body: string, editUrl: string): void {
    this._set('mTitle', title); this._set('mSub', sub); this._set('mBody', body);
    const el = this.domElement.querySelector('#mEdit') as HTMLAnchorElement;
    if (el) el.href = editUrl;
    (this.domElement.querySelector('#pmModal') as HTMLElement)?.classList.add('open');
  }

  private _bindEvents(): void {
    const root = this.domElement;
    const site = this.context.pageContext.web.absoluteUrl;

    root.querySelector('#msFbar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-mf]') as HTMLElement;
      if (!btn) return;
      root.querySelectorAll('#msFbar .fbtn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active'); this._msFilter = btn.dataset.mf || 'all'; this._renderMS();
    });
    root.querySelector('#oiFbar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-of]') as HTMLElement;
      if (!btn) return;
      root.querySelectorAll('#oiFbar .fbtn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active'); this._oiFilter = btn.dataset.of || 'all'; this._renderOI();
    });
    root.querySelector('#oiSearch')?.addEventListener('input', (e) => { this._oiSearch = (e.target as HTMLInputElement).value; this._renderOI(); });

    root.querySelectorAll('[data-sms]').forEach(th => th.addEventListener('click', () => { const f = (th as HTMLElement).dataset.sms || ''; this._ms.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || ''))); this._renderMS(); }));
    root.querySelectorAll('[data-soi]').forEach(th => th.addEventListener('click', () => { const f = (th as HTMLElement).dataset.soi || ''; this._oi.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || ''))); this._renderOI(); }));
    root.querySelectorAll('[data-sdel]').forEach(th => th.addEventListener('click', () => { const f = (th as HTMLElement).dataset.sdel || ''; this._del.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || ''))); this._renderDel(); }));

    root.querySelector('#msBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('[data-msid]') as HTMLElement;
      if (!row) return;
      const m = this._ms.find(x => x.Id === parseInt(row.dataset.msid || '0'));
      if (!m) return;
      this._openModal(m.Title, `Milestone · ${m.Phase}`, `
        <div class="fg"><div class="fl">Phase</div><div class="fv">${this._phaseBadge(m.Phase)}</div></div>
        <div class="fg"><div class="fl">Description</div><div class="fv">${m.MilestoneDesc || '—'}</div></div>
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(m.MilestoneStatus)}</div></div>
        <div class="fg"><div class="fl">Progress</div><div class="fv">${this._progBar(parseInt(m.PercentDone) || 0)}</div></div>
        <div class="fg"><div class="fl">Target End</div><div class="fv">${this._fmt(m.TargetEnd)}</div></div>
        <div class="fg"><div class="fl">Owner</div><div class="fv">${m.MilestoneOwner || '—'}</div></div>
        ${m.MilestoneNotes ? `<div class="fg"><div class="fl">Notes</div><div class="fv" style="font-size:12px">${m.MilestoneNotes}</div></div>` : ''}`,
        `${site}/Lists/PM_Milestones/EditForm.aspx?ID=${m.Id}`);
    });
    root.querySelector('#oiBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('[data-oiid]') as HTMLElement;
      if (!row) return;
      const o = this._oi.find(x => x.Id === parseInt(row.dataset.oiid || '0'));
      if (!o) return;
      this._openModal(o.Title, 'Open Items', `
        <div class="fg"><div class="fl">Description</div><div class="fv">${o.OIDescription || '—'}</div></div>
        <div class="fg"><div class="fl">Owner</div><div class="fv">${o.OIOwner || '—'}</div></div>
        <div class="fg"><div class="fl">Priority</div><div class="fv">${this._priPill(o.OIPriority)}</div></div>
        <div class="fg"><div class="fl">Target Date</div><div class="fv">${this._fmt(o.OITargetDate)}</div></div>
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(o.OIStatus)}</div></div>`,
        `${site}/Lists/PM_OpenItems/EditForm.aspx?ID=${o.Id}`);
    });
    root.querySelector('#delBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('[data-delid]') as HTMLElement;
      if (!row) return;
      const d = this._del.find(x => x.Id === parseInt(row.dataset.delid || '0'));
      if (!d) return;
      this._openModal(d.Title, 'Deliverables', `
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(d.DelStatus)}</div></div>
        <div class="fg"><div class="fl">Delivered</div><div class="fv">${this._fmt(d.DeliveredDate)}</div></div>
        <div class="fg"><div class="fl">Feedback Due</div><div class="fv">${this._fmt(d.FeedbackDue)}</div></div>
        <div class="fg"><div class="fl">Sign-Off Date</div><div class="fv">${this._fmt(d.SignOffDate)}</div></div>
        <div class="fg"><div class="fl">Notes</div><div class="fv" style="font-size:12px">${d.DelNotes || '—'}</div></div>`,
        `${site}/Lists/PM_DocumentDeliverables/EditForm.aspx?ID=${d.Id}`);
    });

    const closeModal = () => (root.querySelector('#pmModal') as HTMLElement)?.classList.remove('open');
    root.querySelector('#mClose')?.addEventListener('click', closeModal);
    root.querySelector('#mDismiss')?.addEventListener('click', closeModal);
    root.querySelector('#pmModal')?.addEventListener('click', (e) => { if ((e.target as HTMLElement).id === 'pmModal') closeModal(); });

    root.querySelector('#btnNewMS')?.addEventListener('click', () => window.open(`${site}/Lists/PM_Milestones/NewForm.aspx`, '_blank'));
    root.querySelector('#btnNewOI')?.addEventListener('click', () => window.open(`${site}/Lists/PM_OpenItems/NewForm.aspx`, '_blank'));
    root.querySelector('#btnNewDel')?.addEventListener('click', () => window.open(`${site}/Lists/PM_DocumentDeliverables/NewForm.aspx`, '_blank'));
    root.querySelector('#btnNewBud')?.addEventListener('click', () => window.open(`${site}/Lists/PM_Budget/NewForm.aspx`, '_blank'));
  }

  private _set(id: string, html: string): void {
    const el = this.domElement.querySelector(`#${id}`) as HTMLElement;
    if (el) el.innerHTML = html;
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
}
