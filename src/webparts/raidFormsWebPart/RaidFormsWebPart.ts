import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IRaidFormsWebPartProps {}

export default class RaidFormsWebPart extends BaseClientSideWebPart<IRaidFormsWebPartProps> {
  private _ac: any[] = [];
  private _iss: any[] = [];
  private _dec: any[] = [];
  private _mtg: any[] = [];
  private _acFilter: string = 'all';
  private _issFilter: string = 'all';
  private _decFilter: string = 'all';
  private _acSearch: string = '';
  private _issSearch: string = '';
  private _decSearch: string = '';
  private _activeTab: string = 'actions';

  public render(): void {
    this.domElement.innerHTML = this._getShell();
    this._loadData();
    this._bindEvents();
  }

  private _getShell(): string {
    return `
<div id="raidRoot" style="font-family:'DM Sans',Segoe UI,sans-serif;background:#f4f6f9;min-height:100vh;color:#1a2332">
  <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600;700&display=swap');
    #raidRoot *{box-sizing:border-box;margin:0;padding:0}
    #raidRoot .ph{background:linear-gradient(135deg,#0a3259 0%,#0f4c81 60%,#1a6bb5 100%);padding:28px 32px 0;overflow:hidden}
    #raidRoot .ph-top{display:flex;align-items:flex-start;justify-content:space-between;gap:16px;position:relative;z-index:1}
    #raidRoot .ph-eye{font-family:'DM Mono',monospace;font-size:10px;letter-spacing:2px;color:rgba(255,255,255,.5);text-transform:uppercase;margin-bottom:6px}
    #raidRoot .ph-title{font-size:22px;font-weight:700;color:#fff}
    #raidRoot .ph-sub{font-size:13px;color:rgba(255,255,255,.65);margin-top:4px}
    #raidRoot .ph-ts{font-family:'DM Mono',monospace;font-size:10px;color:rgba(255,255,255,.4);text-align:right;white-space:nowrap;margin-top:4px}
    #raidRoot .kpi-strip{background:#0a3259;border-top:1px solid rgba(255,255,255,.08);padding:0 32px;display:flex;overflow-x:auto}
    #raidRoot .kpi-c{padding:14px 24px 14px 0;margin-right:24px;border-right:1px solid rgba(255,255,255,.08);min-width:110px;flex-shrink:0}
    #raidRoot .kpi-c:last-child{border-right:none}
    #raidRoot .kpi-v{font-family:'DM Mono',monospace;font-size:24px;font-weight:500;color:#fff;line-height:1}
    #raidRoot .kpi-v.red{color:#ef9a9a} #raidRoot .kpi-v.green{color:#a5d6a7}
    #raidRoot .kpi-l{font-size:10px;color:rgba(255,255,255,.45);letter-spacing:.8px;text-transform:uppercase;margin-top:4px}
    #raidRoot .kpi-load{animation:raidPulse 1.2s ease-in-out infinite;color:rgba(255,255,255,.2)}
    @keyframes raidPulse{0%,100%{opacity:.3}50%{opacity:1}}
    #raidRoot .tab-bar{display:flex;background:rgba(255,255,255,.06);margin-top:16px;border-top:1px solid rgba(255,255,255,.08)}
    #raidRoot .tab-btn{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:12px 20px;color:rgba(255,255,255,.55);background:transparent;border:none;cursor:pointer;border-bottom:3px solid transparent;transition:all .15s;white-space:nowrap}
    #raidRoot .tab-btn:hover{color:rgba(255,255,255,.85)}
    #raidRoot .tab-btn.active{color:#fff;border-bottom-color:#fff;background:rgba(255,255,255,.06)}
    #raidRoot .alert{margin:16px 32px 0;padding:12px 16px;border-radius:8px;border-left:4px solid #c62828;background:#fde8e8;color:#c62828;display:none;align-items:center;gap:10px;font-size:13px;font-weight:500}
    #raidRoot .alert.show{display:flex}
    #raidRoot .content{padding:20px 32px 32px}
    #raidRoot .tab-panel{display:none} #raidRoot .tab-panel.active{display:block}
    #raidRoot .sec-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;margin-top:24px}
    #raidRoot .sec-title{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#0f4c81;display:flex;align-items:center;gap:8px}
    #raidRoot .sec-title::before{content:'';width:3px;height:14px;background:#0f4c81;border-radius:2px;display:block}
    #raidRoot .sec-cnt{font-family:'DM Mono',monospace;font-size:11px;color:#5a6a7e;background:#e0e6ed;padding:2px 8px;border-radius:10px}
    #raidRoot .fbar{display:flex;gap:8px;margin-bottom:12px;flex-wrap:wrap;align-items:center}
    #raidRoot .fbtn{font-family:'DM Sans',sans-serif;font-size:11px;font-weight:600;padding:5px 12px;border-radius:20px;border:1px solid #e0e6ed;background:#fff;color:#5a6a7e;cursor:pointer;transition:all .15s}
    #raidRoot .fbtn:hover,#raidRoot .fbtn.active{background:#0f4c81;color:#fff;border-color:#0f4c81}
    #raidRoot .sbox{font-family:'DM Sans',sans-serif;font-size:12px;padding:5px 12px;border-radius:20px;border:1px solid #e0e6ed;background:#fff;outline:none;min-width:180px;margin-left:auto}
    #raidRoot .sbox:focus{border-color:#0f4c81}
    #raidRoot .tcard{background:#fff;border-radius:10px;border:1px solid #e0e6ed;overflow:hidden;box-shadow:0 1px 4px rgba(0,0,0,.05)}
    #raidRoot table{width:100%;border-collapse:collapse;font-size:12.5px}
    #raidRoot thead th{background:#f8fafc;padding:10px 14px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;border-bottom:1px solid #e0e6ed;white-space:nowrap;cursor:pointer}
    #raidRoot thead th:hover{color:#0f4c81}
    #raidRoot tbody tr{border-bottom:1px solid #f0f4f8;transition:background .1s;cursor:pointer}
    #raidRoot tbody tr:last-child{border-bottom:none}
    #raidRoot tbody tr:hover{background:#f8fafc}
    #raidRoot td{padding:10px 14px;vertical-align:middle}
    #raidRoot .cid{font-family:'DM Mono',monospace;font-size:11px;color:#0f4c81;font-weight:500}
    #raidRoot .cmut{font-size:11px;color:#5a6a7e}
    #raidRoot .cdate{font-family:'DM Mono',monospace;font-size:11px;color:#5a6a7e;white-space:nowrap}
    #raidRoot .pill{display:inline-block;padding:3px 10px;border-radius:20px;font-size:10.5px;font-weight:600;white-space:nowrap}
    #raidRoot .pr{background:#fde8e8;color:#c62828} #raidRoot .pg{background:#e8f5e9;color:#2e7d32}
    #raidRoot .pa{background:#fff3e0;color:#e65100} #raidRoot .pb{background:#e3f2fd;color:#0d47a1} #raidRoot .pz{background:#f0f0f0;color:#616161}
    #raidRoot .pri{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700}
    #raidRoot .pc{background:#fde8e8;color:#c62828} #raidRoot .ph2{background:#fff3e0;color:#e65100} #raidRoot .pm{background:#e3f2fd;color:#0d47a1}
    #raidRoot .sev-c{background:#fde8e8;color:#c62828} #raidRoot .sev-h{background:#fff3e0;color:#e65100}
    #raidRoot .sev-m{background:#fff8e1;color:#f57f17} #raidRoot .sev-l{background:#e8f5e9;color:#2e7d32}
    #raidRoot .mtg-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(220px,1fr));gap:12px;margin-bottom:16px}
    #raidRoot .mc{background:#fff;border-radius:10px;border:1px solid #e0e6ed;padding:16px;cursor:pointer;transition:all .15s;box-shadow:0 1px 3px rgba(0,0,0,.04)}
    #raidRoot .mc:hover{border-color:#0f4c81;box-shadow:0 4px 12px rgba(15,76,129,.1);transform:translateY(-1px)}
    #raidRoot .mc-id{font-family:'DM Mono',monospace;font-size:11px;color:#0f4c81;font-weight:500;margin-bottom:4px}
    #raidRoot .mc-date{font-size:11px;color:#5a6a7e;margin-bottom:6px}
    #raidRoot .mc-sum{font-size:11.5px;color:#1a2332;line-height:1.4}
    #raidRoot .dec-list{display:flex;flex-direction:column}
    #raidRoot .dec-item{display:flex;gap:0;cursor:pointer}
    #raidRoot .dec-item:hover .dec-body{background:#f8fafc}
    #raidRoot .dec-line{display:flex;flex-direction:column;align-items:center;width:32px;flex-shrink:0}
    #raidRoot .dec-dot{width:10px;height:10px;border-radius:50%;background:#0f4c81;flex-shrink:0;margin-top:4px}
    #raidRoot .dec-con{flex:1;width:2px;background:#e0e6ed}
    #raidRoot .dec-item:last-child .dec-con{display:none}
    #raidRoot .dec-body{flex:1;padding:10px 14px 14px;border-radius:8px;margin-bottom:4px;transition:background .1s}
    #raidRoot .dec-ref{font-family:'DM Mono',monospace;font-size:10px;color:#0f4c81;font-weight:500}
    #raidRoot .dec-mtg{font-size:10px;color:#5a6a7e;margin-left:8px}
    #raidRoot .dec-text{font-size:12px;color:#1a2332;margin-top:4px;line-height:1.4}
    #raidRoot .dec-by{font-size:10.5px;color:#5a6a7e;margin-top:6px}
    #raidRoot .btn-new{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:7px 16px;border-radius:6px;background:#0f4c81;color:#fff;border:none;cursor:pointer;display:inline-flex;align-items:center;gap:6px;transition:background .15s}
    #raidRoot .btn-new:hover{background:#1a6bb5}
    #raidRoot .load-row td{padding:24px;text-align:center;color:#5a6a7e;font-size:12px;font-family:'DM Mono',monospace}
    #raidRoot .modal-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.4);z-index:9999;align-items:center;justify-content:center;padding:20px}
    #raidRoot .modal-ov.open{display:flex}
    #raidRoot .modal{background:#fff;border-radius:12px;width:100%;max-width:620px;max-height:85vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
    #raidRoot .modal-hdr{padding:20px 24px 16px;border-bottom:1px solid #e0e6ed;display:flex;justify-content:space-between;align-items:flex-start;position:sticky;top:0;background:#fff;z-index:1}
    #raidRoot .modal-title{font-size:15px;font-weight:700;color:#0f4c81}
    #raidRoot .modal-sub{font-family:'DM Mono',monospace;font-size:10px;color:#5a6a7e;margin-top:2px}
    #raidRoot .modal-x{background:none;border:none;font-size:22px;cursor:pointer;color:#5a6a7e;padding:0 4px;line-height:1}
    #raidRoot .modal-body{padding:20px 24px}
    #raidRoot .fg{margin-bottom:14px}
    #raidRoot .fl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#5a6a7e;margin-bottom:4px}
    #raidRoot .fv{font-size:13px;color:#1a2332;line-height:1.5;padding:8px 10px;background:#f8fafc;border-radius:6px;border:1px solid #e0e6ed}
    #raidRoot .fv.blk{color:#c62828}
    #raidRoot .modal-ft{padding:16px 24px;border-top:1px solid #e0e6ed;display:flex;justify-content:flex-end;gap:8px}
    #raidRoot .btn-sec{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:8px 18px;border-radius:6px;background:#f0f4f8;color:#1a2332;border:1px solid #e0e6ed;cursor:pointer}
    #raidRoot .btn-pri{font-family:'DM Sans',sans-serif;font-size:12px;font-weight:600;padding:8px 18px;border-radius:6px;background:#0f4c81;color:#fff;border:none;cursor:pointer;text-decoration:none;display:inline-block}
  </style>

  <div class="ph">
    <div class="ph-top">
      <div>
        <div class="ph-eye">IMP9177 · 3H Pharmaceuticals LLC</div>
        <div class="ph-title">RAID Register</div>
        <div class="ph-sub" id="raidSub">Loading live data...</div>
      </div>
      <div><div class="ph-ts" id="raidTs">Fetching...</div></div>
    </div>
    <div class="kpi-strip">
      <div class="kpi-c"><div class="kpi-v" id="kAT"><span class="kpi-load">—</span></div><div class="kpi-l">Actions</div></div>
      <div class="kpi-c"><div class="kpi-v red" id="kAO"><span class="kpi-load">—</span></div><div class="kpi-l">Open</div></div>
      <div class="kpi-c"><div class="kpi-v green" id="kAC"><span class="kpi-load">—</span></div><div class="kpi-l">Complete</div></div>
      <div class="kpi-c"><div class="kpi-v red" id="kIS"><span class="kpi-load">—</span></div><div class="kpi-l">Issues</div></div>
      <div class="kpi-c"><div class="kpi-v" id="kDE"><span class="kpi-load">—</span></div><div class="kpi-l">Decisions</div></div>
      <div class="kpi-c"><div class="kpi-v" id="kMT"><span class="kpi-load">—</span></div><div class="kpi-l">Meetings</div></div>
    </div>
    <div class="tab-bar">
      <button class="tab-btn active" data-tab="actions">Actions (AC)</button>
      <button class="tab-btn" data-tab="issues">Issues (ISS)</button>
      <button class="tab-btn" data-tab="decisions">Decisions (DEC)</button>
      <button class="tab-btn" data-tab="meetings">Meetings (MM)</button>
    </div>
  </div>

  <div class="alert" id="raidAlert"><span>⚠️</span><span id="raidAlertTxt"></span></div>

  <div class="content">
    <div class="tab-panel active" id="tp-actions">
      <div class="sec-hdr">
        <div class="sec-title">Action Register</div>
        <div style="display:flex;align-items:center;gap:8px"><span class="sec-cnt" id="acCnt">—</span><button class="btn-new" id="btnNewAC">+ New Action</button></div>
      </div>
      <div class="fbar" id="acFbar">
        <button class="fbtn active" data-af="all">All</button>
        <button class="fbtn" data-af="open">Open</button>
        <button class="fbtn" data-af="critical">🔴 Critical</button>
        <button class="fbtn" data-af="adb">ADB Owned</button>
        <button class="fbtn" data-af="3h">3H Owned</button>
        <button class="fbtn" data-af="complete">Complete</button>
        <input class="sbox" id="acSearch" placeholder="Search actions...">
      </div>
      <div class="tcard">
        <table><thead><tr>
          <th data-sa="Title">AC#</th><th>Category</th><th>Description</th><th>Owner</th>
          <th data-sa="ActionStatus">Status</th><th data-sa="ActionPriority">Priority</th>
          <th data-sa="DueDate">Due</th><th>Blocked By</th>
        </tr></thead>
        <tbody id="acBody"><tr class="load-row"><td colspan="8">Loading actions...</td></tr></tbody></table>
      </div>
    </div>

    <div class="tab-panel" id="tp-issues">
      <div class="sec-hdr">
        <div class="sec-title">Issue Register</div>
        <div style="display:flex;align-items:center;gap:8px"><span class="sec-cnt" id="issCnt">—</span><button class="btn-new" id="btnNewISS">+ New Issue</button></div>
      </div>
      <div class="fbar" id="issFbar">
        <button class="fbtn active" data-if="all">All</button>
        <button class="fbtn" data-if="open">Open</button>
        <button class="fbtn" data-if="critical">Critical</button>
        <button class="fbtn" data-if="closed">Closed</button>
        <input class="sbox" id="issSearch" placeholder="Search issues...">
      </div>
      <div class="tcard">
        <table><thead><tr>
          <th data-si="Title">ISS#</th><th>Description</th><th>Regulation</th>
          <th data-si="IssueSeverity">Severity</th><th data-si="IssueStatus">Status</th><th>Owner</th><th>CAPA</th>
        </tr></thead>
        <tbody id="issBody"><tr class="load-row"><td colspan="7">Loading issues...</td></tr></tbody></table>
      </div>
    </div>

    <div class="tab-panel" id="tp-decisions">
      <div class="sec-hdr">
        <div class="sec-title">Decision Log</div>
        <div style="display:flex;align-items:center;gap:8px"><span class="sec-cnt" id="decCnt">—</span><button class="btn-new" id="btnNewDEC">+ New Decision</button></div>
      </div>
      <div class="fbar" id="decFbar">
        <button class="fbtn active" data-df="all">All</button>
        <button class="fbtn" data-df="MM-005">MM-005</button>
        <button class="fbtn" data-df="MM-004">MM-004</button>
        <button class="fbtn" data-df="MM-003">MM-003</button>
        <input class="sbox" id="decSearch" placeholder="Search decisions...">
      </div>
      <div class="dec-list" id="decList"></div>
    </div>

    <div class="tab-panel" id="tp-meetings">
      <div class="sec-hdr">
        <div class="sec-title">Meeting Log</div>
        <div style="display:flex;align-items:center;gap:8px"><span class="sec-cnt" id="mtgCnt">—</span><button class="btn-new" id="btnNewMTG">+ New Meeting</button></div>
      </div>
      <div class="mtg-grid" id="mtgGrid"></div>
      <div class="tcard">
        <table><thead><tr>
          <th>MM#</th><th data-sm="MeetingDate">Date</th><th>Type</th><th>Platform</th>
          <th>Duration</th><th>Agenda</th><th>Actions</th><th>Decisions</th>
        </tr></thead>
        <tbody id="mtgBody"><tr class="load-row"><td colspan="8">Loading meetings...</td></tr></tbody></table>
      </div>
    </div>
  </div>

  <div class="modal-ov" id="raidModal">
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
      const [aRes, iRes, dRes, mRes]: SPHttpClientResponse[] = await Promise.all([
        this.context.spHttpClient.get(`${site}/_api/web/lists/getbytitle('RAID_Actions')/items?$select=Id,Title,ActionCategory,ActionDesc,ActionOwner,ActionStatus,ActionPriority,DueDate,RaisedInMeeting,UpdateNotes,CompletionDate,BlockedBy&$top=500&$orderby=Title`, SPHttpClient.configurations.v1),
        this.context.spHttpClient.get(`${site}/_api/web/lists/getbytitle('RAID_Issues')/items?$select=Id,Title,IssueDescription,RegulationRef,IssueSeverity,IssueStatus,IssueOwner,MitigationDescription,ResolutionDescription,ClosureDate,LinkedCAPARef&$top=500&$orderby=Title`, SPHttpClient.configurations.v1),
        this.context.spHttpClient.get(`${site}/_api/web/lists/getbytitle('RAID_Decisions')/items?$select=Id,Title,DecisionDate,MeetingRef,DecisionDesc,MadeBy,LinkedActionRefs&$top=500&$orderby=Title`, SPHttpClient.configurations.v1),
        this.context.spHttpClient.get(`${site}/_api/web/lists/getbytitle('RAID_Meetings')/items?$select=Id,Title,MeetingDate,MeetingType,Platform,Duration,AgendaSummary,ActionsRaised,DecisionsMade,NextMeetingTarget,MinutesURL,MeetingNotes&$top=500&$orderby=Title`, SPHttpClient.configurations.v1)
      ]);
      const [aD, iD, dD, mD] = await Promise.all([aRes.json(), iRes.json(), dRes.json(), mRes.json()]);
      this._ac = aD.value || []; this._iss = iD.value || []; this._dec = dD.value || []; this._mtg = mD.value || [];
    } catch (e) { console.error('RAID Forms: load failed', e); }
    this._updateKPIs(); this._renderAC(); this._renderISS(); this._renderDEC(); this._renderMTG();
    this._set('raidTs', 'Refreshed ' + new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' }));
  }

  private _fmt(s: string): string {
    if (!s) return '—';
    try { return new Date(s).toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' }); } catch { return s; }
  }

  private _pill(s: string): string {
    if (!s) return '<span class="pill pz">—</span>';
    const l = s.toLowerCase();
    if (l.includes('past due') || l.includes('critical') || l.includes('overdue') || l.includes('blocking')) return `<span class="pill pr">${s}</span>`;
    if (l.includes('complete') || l.includes('closed')) return `<span class="pill pg">${s}</span>`;
    if (l.includes('partial') || l.includes('progress') || l.includes('drafting') || l.includes('researching')) return `<span class="pill pa">${s}</span>`;
    if (l.includes('open') || l.includes('pending') || l.includes('new') || l.includes('issued')) return `<span class="pill pb">${s}</span>`;
    return `<span class="pill pz">${s}</span>`;
  }

  private _priPill(s: string): string {
    if (!s) return '';
    const l = s.toLowerCase();
    if (l === 'critical') return `<span class="pri pc">${s}</span>`;
    if (l === 'high') return `<span class="pri ph2">${s}</span>`;
    return `<span class="pri pm">${s}</span>`;
  }

  private _sevPill(s: string): string {
    if (!s) return '<span class="pill pz">—</span>';
    const l = s.toLowerCase();
    if (l === 'critical') return `<span class="pill sev-c">${s}</span>`;
    if (l === 'high') return `<span class="pill sev-h">${s}</span>`;
    if (l === 'major') return `<span class="pill sev-m">${s}</span>`;
    return `<span class="pill sev-l">${s}</span>`;
  }

  private _updateKPIs(): void {
    const open = this._ac.filter(a => !['complete', 'closed'].some(x => (a.ActionStatus || '').toLowerCase().includes(x))).length;
    const critIss = this._iss.filter(i => (i.IssueSeverity || '').toLowerCase() === 'critical' && !(i.IssueStatus || '').toLowerCase().includes('closed')).length;
    this._set('kAT', String(this._ac.length)); this._set('kAO', String(open));
    this._set('kAC', String(this._ac.length - open)); this._set('kIS', String(this._iss.length));
    this._set('kDE', String(this._dec.length)); this._set('kMT', String(this._mtg.length));
    this._set('raidSub', `${this._ac.length} actions (${open} open) · ${this._iss.length} issues · ${this._dec.length} decisions`);
    if (critIss > 0) {
      this._set('raidAlertTxt', `${critIss} CRITICAL open issue(s) requiring immediate attention`);
      (this.domElement.querySelector('#raidAlert') as HTMLElement)?.classList.add('show');
    }
  }

  private _filteredAC(): any[] {
    let l = this._ac;
    if (this._acFilter === 'open') l = l.filter(a => !['complete', 'closed'].some(x => (a.ActionStatus || '').toLowerCase().includes(x)));
    else if (this._acFilter === 'critical') l = l.filter(a => (a.ActionPriority || '').toLowerCase() === 'critical');
    else if (this._acFilter === 'adb') l = l.filter(a => (a.ActionOwner || '').toLowerCase().includes('andre') || (a.ActionOwner || '').toLowerCase().includes('adb'));
    else if (this._acFilter === '3h') l = l.filter(a => ['3h', 'tina', 'cindy', 'kj', 'baolong'].some(x => (a.ActionOwner || '').toLowerCase().includes(x)));
    else if (this._acFilter === 'complete') l = l.filter(a => ['complete', 'closed'].some(x => (a.ActionStatus || '').toLowerCase().includes(x)));
    if (this._acSearch) { const q = this._acSearch.toLowerCase(); l = l.filter(a => (a.Title || '').toLowerCase().includes(q) || (a.ActionDesc || '').toLowerCase().includes(q) || (a.ActionOwner || '').toLowerCase().includes(q)); }
    return l;
  }

  private _renderAC(): void {
    const l = this._filteredAC();
    this._set('acCnt', String(l.length));
    const tb = this.domElement.querySelector('#acBody') as HTMLElement;
    if (!tb) return;
    tb.innerHTML = l.length ? l.map(a => `<tr data-aid="${a.Id}">
      <td><span class="cid">${a.Title}</span></td>
      <td><span class="cmut">${a.ActionCategory || '—'}</span></td>
      <td><span style="font-size:12px;display:block;max-width:240px">${(a.ActionDesc || '—').substring(0, 80)}${(a.ActionDesc || '').length > 80 ? '…' : ''}</span></td>
      <td><span class="cmut">${a.ActionOwner || '—'}</span></td>
      <td>${this._pill(a.ActionStatus)}</td>
      <td>${this._priPill(a.ActionPriority)}</td>
      <td><span class="cdate">${this._fmt(a.DueDate)}</span></td>
      <td><span class="cmut" style="font-size:10.5px">${a.BlockedBy || '—'}</span></td>
    </tr>`).join('') : '<tr><td colspan="8" style="text-align:center;padding:32px;color:#aaa">No actions match filter</td></tr>';
  }

  private _filteredISS(): any[] {
    let l = this._iss;
    if (this._issFilter === 'open') l = l.filter(i => !(i.IssueStatus || '').toLowerCase().includes('closed'));
    else if (this._issFilter === 'critical') l = l.filter(i => (i.IssueSeverity || '').toLowerCase() === 'critical');
    else if (this._issFilter === 'closed') l = l.filter(i => (i.IssueStatus || '').toLowerCase().includes('closed'));
    if (this._issSearch) { const q = this._issSearch.toLowerCase(); l = l.filter(i => (i.Title || '').toLowerCase().includes(q) || (i.IssueDescription || '').toLowerCase().includes(q)); }
    return l;
  }

  private _renderISS(): void {
    const l = this._filteredISS();
    this._set('issCnt', String(l.length));
    const tb = this.domElement.querySelector('#issBody') as HTMLElement;
    if (!tb) return;
    tb.innerHTML = l.length ? l.map(i => `<tr data-iid="${i.Id}">
      <td><span class="cid">${i.Title}</span></td>
      <td><span style="font-size:12px;display:block;max-width:240px">${(i.IssueDescription || '—').substring(0, 90)}${(i.IssueDescription || '').length > 90 ? '…' : ''}</span></td>
      <td><span class="cmut" style="font-family:'DM Mono',monospace;font-size:10.5px">${i.RegulationRef || '—'}</span></td>
      <td>${this._sevPill(i.IssueSeverity)}</td>
      <td>${this._pill(i.IssueStatus)}</td>
      <td><span class="cmut">${i.IssueOwner || '—'}</span></td>
      <td><span class="cmut">${i.LinkedCAPARef || '—'}</span></td>
    </tr>`).join('') : '<tr><td colspan="7" style="text-align:center;padding:32px;color:#aaa">No issues match filter</td></tr>';
  }

  private _filteredDEC(): any[] {
    let l = this._dec;
    if (this._decFilter !== 'all') l = l.filter(d => d.MeetingRef === this._decFilter);
    if (this._decSearch) { const q = this._decSearch.toLowerCase(); l = l.filter(d => (d.Title || '').toLowerCase().includes(q) || (d.DecisionDesc || '').toLowerCase().includes(q)); }
    return l;
  }

  private _renderDEC(): void {
    const l = this._filteredDEC();
    this._set('decCnt', String(l.length));
    const el = this.domElement.querySelector('#decList') as HTMLElement;
    if (!el) return;
    el.innerHTML = l.length ? l.map(d => `<div class="dec-item" data-did="${d.Id}">
      <div class="dec-line"><div class="dec-dot"></div><div class="dec-con"></div></div>
      <div class="dec-body">
        <div><span class="dec-ref">${d.Title}</span><span class="dec-mtg">${d.MeetingRef || ''} · ${this._fmt(d.DecisionDate)}</span></div>
        <div class="dec-text">${d.DecisionDesc || '—'}</div>
        <div class="dec-by">By: ${d.MadeBy || '—'}${d.LinkedActionRefs ? ` · Links: ${d.LinkedActionRefs}` : ''}</div>
      </div>
    </div>`).join('') : '<div style="text-align:center;padding:40px;color:#aaa">No decisions match filter</div>';
  }

  private _renderMTG(): void {
    this._set('mtgCnt', String(this._mtg.length));
    const grid = this.domElement.querySelector('#mtgGrid') as HTMLElement;
    if (grid) grid.innerHTML = this._mtg.map(m => `<div class="mc" data-mid="${m.Id}">
      <div class="mc-id">${m.Title}</div>
      <div class="mc-date">📅 ${this._fmt(m.MeetingDate)} · ${m.Duration || '?'} min · ${m.Platform || ''}</div>
      <div style="margin-bottom:8px"><span class="pill pb">${m.MeetingType || 'Meeting'}</span></div>
      <div class="mc-sum">${(m.AgendaSummary || '').substring(0, 100)}${(m.AgendaSummary || '').length > 100 ? '…' : ''}</div>
    </div>`).join('');
    const tb = this.domElement.querySelector('#mtgBody') as HTMLElement;
    if (tb) tb.innerHTML = this._mtg.map(m => `<tr data-mid="${m.Id}">
      <td><span class="cid">${m.Title}</span></td>
      <td><span class="cdate">${this._fmt(m.MeetingDate)}</span></td>
      <td><span class="cmut">${m.MeetingType || '—'}</span></td>
      <td><span class="cmut">${m.Platform || '—'}</span></td>
      <td><span class="cmut">${m.Duration || '—'} min</span></td>
      <td><span style="font-size:11.5px;display:block;max-width:200px">${(m.AgendaSummary || '—').substring(0, 80)}${(m.AgendaSummary || '').length > 80 ? '…' : ''}</span></td>
      <td><span class="cmut">${m.ActionsRaised || 0}</span></td>
      <td><span class="cmut">${m.DecisionsMade || 0}</span></td>
    </tr>`).join('');
  }

  private _openModal(title: string, sub: string, body: string, editUrl: string): void {
    this._set('mTitle', title); this._set('mSub', sub); this._set('mBody', body);
    const el = this.domElement.querySelector('#mEdit') as HTMLAnchorElement;
    if (el) el.href = editUrl;
    (this.domElement.querySelector('#raidModal') as HTMLElement)?.classList.add('open');
  }

  private _bindEvents(): void {
    const root = this.domElement;
    const site = this.context.pageContext.web.absoluteUrl;

    // Tab switching
    root.querySelector('.tab-bar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-tab]') as HTMLElement;
      if (!btn) return;
      const tab = btn.dataset.tab || 'actions';
      root.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
      root.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      (root.querySelector(`#tp-${tab}`) as HTMLElement)?.classList.add('active');
    });

    // AC filter
    root.querySelector('#acFbar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-af]') as HTMLElement;
      if (!btn) return;
      root.querySelectorAll('#acFbar .fbtn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active'); this._acFilter = btn.dataset.af || 'all'; this._renderAC();
    });
    root.querySelector('#acSearch')?.addEventListener('input', (e) => { this._acSearch = (e.target as HTMLInputElement).value; this._renderAC(); });

    // ISS filter
    root.querySelector('#issFbar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-if]') as HTMLElement;
      if (!btn) return;
      root.querySelectorAll('#issFbar .fbtn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active'); this._issFilter = btn.dataset.if || 'all'; this._renderISS();
    });
    root.querySelector('#issSearch')?.addEventListener('input', (e) => { this._issSearch = (e.target as HTMLInputElement).value; this._renderISS(); });

    // DEC filter
    root.querySelector('#decFbar')?.addEventListener('click', (e) => {
      const btn = (e.target as HTMLElement).closest('[data-df]') as HTMLElement;
      if (!btn) return;
      root.querySelectorAll('#decFbar .fbtn').forEach(b => b.classList.remove('active'));
      btn.classList.add('active'); this._decFilter = btn.dataset.df || 'all'; this._renderDEC();
    });
    root.querySelector('#decSearch')?.addEventListener('input', (e) => { this._decSearch = (e.target as HTMLInputElement).value; this._renderDEC(); });

    // Sort
    root.querySelectorAll('[data-sa]').forEach(th => th.addEventListener('click', () => { const f = (th as HTMLElement).dataset.sa || ''; this._ac.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || ''))); this._renderAC(); }));
    root.querySelectorAll('[data-si]').forEach(th => th.addEventListener('click', () => { const f = (th as HTMLElement).dataset.si || ''; this._iss.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || ''))); this._renderISS(); }));
    root.querySelectorAll('[data-sm]').forEach(th => th.addEventListener('click', () => { const f = (th as HTMLElement).dataset.sm || ''; this._mtg.sort((a, b) => String(a[f] || '').localeCompare(String(b[f] || ''))); this._renderMTG(); }));

    // Row clicks
    root.querySelector('#acBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('[data-aid]') as HTMLElement;
      if (!row) return;
      const a = this._ac.find(x => x.Id === parseInt(row.dataset.aid || '0'));
      if (!a) return;
      this._openModal(a.Title, `Action Register · ${a.ActionCategory}`, `
        <div class="fg"><div class="fl">Description</div><div class="fv">${a.ActionDesc || '—'}</div></div>
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(a.ActionStatus)}</div></div>
        <div class="fg"><div class="fl">Priority</div><div class="fv">${this._priPill(a.ActionPriority)}</div></div>
        <div class="fg"><div class="fl">Owner</div><div class="fv">${a.ActionOwner || '—'}</div></div>
        <div class="fg"><div class="fl">Due Date</div><div class="fv">${this._fmt(a.DueDate)}</div></div>
        <div class="fg"><div class="fl">Raised In</div><div class="fv">${a.RaisedInMeeting || '—'}</div></div>
        ${a.BlockedBy ? `<div class="fg"><div class="fl">Blocked By</div><div class="fv blk">${a.BlockedBy}</div></div>` : ''}
        <div class="fg"><div class="fl">Update Notes</div><div class="fv" style="font-size:12px">${a.UpdateNotes || '—'}</div></div>`,
        `${site}/Lists/RAID_Actions/EditForm.aspx?ID=${a.Id}`);
    });

    root.querySelector('#issBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('[data-iid]') as HTMLElement;
      if (!row) return;
      const i = this._iss.find(x => x.Id === parseInt(row.dataset.iid || '0'));
      if (!i) return;
      this._openModal(i.Title, 'Issue Register', `
        <div class="fg"><div class="fl">Description</div><div class="fv">${i.IssueDescription || '—'}</div></div>
        <div class="fg"><div class="fl">Regulation</div><div class="fv">${i.RegulationRef || '—'}</div></div>
        <div class="fg"><div class="fl">Severity</div><div class="fv">${this._sevPill(i.IssueSeverity)}</div></div>
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(i.IssueStatus)}</div></div>
        <div class="fg"><div class="fl">Owner</div><div class="fv">${i.IssueOwner || '—'}</div></div>
        <div class="fg"><div class="fl">Mitigation</div><div class="fv" style="font-size:12px">${i.MitigationDescription || '—'}</div></div>`,
        `${site}/Lists/RAID_Issues/EditForm.aspx?ID=${i.Id}`);
    });

    root.querySelector('#decList')?.addEventListener('click', (e) => {
      const item = (e.target as HTMLElement).closest('[data-did]') as HTMLElement;
      if (!item) return;
      const d = this._dec.find(x => x.Id === parseInt(item.dataset.did || '0'));
      if (!d) return;
      this._openModal(d.Title, `Decision Log · ${d.MeetingRef}`, `
        <div class="fg"><div class="fl">Meeting</div><div class="fv">${d.MeetingRef || '—'}</div></div>
        <div class="fg"><div class="fl">Date</div><div class="fv">${this._fmt(d.DecisionDate)}</div></div>
        <div class="fg"><div class="fl">Decision</div><div class="fv">${d.DecisionDesc || '—'}</div></div>
        <div class="fg"><div class="fl">Made By</div><div class="fv">${d.MadeBy || '—'}</div></div>
        ${d.LinkedActionRefs ? `<div class="fg"><div class="fl">Linked Actions</div><div class="fv">${d.LinkedActionRefs}</div></div>` : ''}`,
        `${site}/Lists/RAID_Decisions/EditForm.aspx?ID=${d.Id}`);
    });

    root.querySelector('#mtgGrid')?.addEventListener('click', (e) => {
      const card = (e.target as HTMLElement).closest('[data-mid]') as HTMLElement;
      if (card) this._showMTGModal(parseInt(card.dataset.mid || '0'), site);
    });
    root.querySelector('#mtgBody')?.addEventListener('click', (e) => {
      const row = (e.target as HTMLElement).closest('[data-mid]') as HTMLElement;
      if (row) this._showMTGModal(parseInt(row.dataset.mid || '0'), site);
    });

    // Modal close
    const closeModal = () => (root.querySelector('#raidModal') as HTMLElement)?.classList.remove('open');
    root.querySelector('#mClose')?.addEventListener('click', closeModal);
    root.querySelector('#mDismiss')?.addEventListener('click', closeModal);
    root.querySelector('#raidModal')?.addEventListener('click', (e) => { if ((e.target as HTMLElement).id === 'raidModal') closeModal(); });

    // New buttons
    root.querySelector('#btnNewAC')?.addEventListener('click', () => window.open(`${site}/Lists/RAID_Actions/NewForm.aspx`, '_blank'));
    root.querySelector('#btnNewISS')?.addEventListener('click', () => window.open(`${site}/Lists/RAID_Issues/NewForm.aspx`, '_blank'));
    root.querySelector('#btnNewDEC')?.addEventListener('click', () => window.open(`${site}/Lists/RAID_Decisions/NewForm.aspx`, '_blank'));
    root.querySelector('#btnNewMTG')?.addEventListener('click', () => window.open(`${site}/Lists/RAID_Meetings/NewForm.aspx`, '_blank'));
  }

  private _showMTGModal(id: number, site: string): void {
    const m = this._mtg.find(x => x.Id === id);
    if (!m) return;
    this._openModal(m.Title, `Meeting Log · ${m.MeetingType}`, `
      <div class="fg"><div class="fl">Date</div><div class="fv">${this._fmt(m.MeetingDate)}</div></div>
      <div class="fg"><div class="fl">Type / Platform</div><div class="fv">${m.MeetingType || '—'} · ${m.Platform || '—'} · ${m.Duration || '?'} min</div></div>
      <div class="fg"><div class="fl">Agenda</div><div class="fv" style="font-size:12px">${m.AgendaSummary || '—'}</div></div>
      <div class="fg"><div class="fl">Notes</div><div class="fv" style="font-size:12px">${m.MeetingNotes || '—'}</div></div>
      <div class="fg"><div class="fl">Actions Raised</div><div class="fv">${m.ActionsRaised || 0}</div></div>
      <div class="fg"><div class="fl">Next Meeting</div><div class="fv">${this._fmt(m.NextMeetingTarget)}</div></div>`,
      `${site}/Lists/RAID_Meetings/EditForm.aspx?ID=${id}`);
  }

  private _set(id: string, html: string): void {
    const el = this.domElement.querySelector(`#${id}`) as HTMLElement;
    if (el) el.innerHTML = html;
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }
}
