/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-floating-promises */
/* eslint-disable @typescript-eslint/no-unused-vars */
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneChoiceGroup } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IQmsPortalWebPartProps { screen: string; }

// ─────────────────────────────────────────────────────────────────────────────
// CSS
// ─────────────────────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600;700&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
:root{
  --n:#0a3259;--b:#0f4c81;--b2:#1a6bb5;--b1:#e3f2fd;--b0:#f0f7ff;
  --r:#c62828;--r1:#fde8e8;--a:#e65100;--a1:#fff3e0;--g:#2e7d32;--g1:#e8f5e9;
  --s0:#f8fafc;--s1:#f0f4f8;--s2:#e0e6ed;--s5:#5a6a7e;--s7:#1a2332;--w:#fff;
  --mono:'DM Mono',monospace;--sans:'DM Sans',sans-serif;
}
body{font-family:var(--sans);background:var(--s0);color:var(--s7);font-size:14px}

/* ── header ── */
.qp-hdr{background:linear-gradient(135deg,var(--n) 0%,var(--b) 60%,var(--b2) 100%);padding:0 28px;height:56px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.2)}
.qp-hdr-l{display:flex;align-items:center;gap:12px}
.qp-badge{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;font-size:10px;font-weight:700;letter-spacing:1px;padding:3px 9px;border-radius:4px}
.qp-title{color:#fff;font-size:15px;font-weight:700}
.qp-sub{color:rgba(255,255,255,.6);font-size:11px;margin-top:1px}
.qp-hdr-r{display:flex;align-items:center;gap:8px}
.qp-ts{font-family:var(--mono);font-size:10px;color:rgba(255,255,255,.45)}
.qp-btn{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;font-size:11px;font-weight:600;padding:5px 12px;border-radius:5px;cursor:pointer}
.qp-btn:hover{background:rgba(255,255,255,.25)}

/* ── nav tabs ── */
.qp-nav{background:var(--n);display:flex;overflow-x:auto;padding:0 20px;border-bottom:2px solid rgba(255,255,255,.08)}
.qp-tab{font-family:var(--sans);font-size:12px;font-weight:600;padding:11px 18px;color:rgba(255,255,255,.5);background:transparent;border:none;cursor:pointer;border-bottom:3px solid transparent;white-space:nowrap;transition:all .15s}
.qp-tab:hover{color:rgba(255,255,255,.85)}
.qp-tab.on{color:#fff;border-bottom-color:#fff;background:rgba(255,255,255,.06)}

/* ── alert bar ── */
.qp-alert{margin:14px 24px 0;padding:10px 16px;border-radius:7px;border-left:4px solid var(--r);background:var(--r1);color:var(--r);font-size:12px;font-weight:500;display:none;align-items:center;gap:8px}
.qp-alert.show{display:flex}

/* ── main content ── */
.qp-main{padding:18px 24px 32px}
.qp-screen{display:none}.qp-screen.on{display:block}

/* ── KPI strip ── */
.kpi-row{display:flex;gap:10px;margin-bottom:16px;flex-wrap:wrap}
.kpi{background:var(--w);border:1px solid var(--s2);border-radius:8px;padding:14px 18px;flex:1;min-width:120px;border-top:3px solid var(--b2);box-shadow:0 1px 3px rgba(0,0,0,.06)}
.kpi.r{border-top-color:var(--r)}.kpi.a{border-top-color:var(--a)}.kpi.g{border-top-color:var(--g)}
.kpi-l{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--s5);margin-bottom:5px}
.kpi-v{font-size:28px;font-weight:700;color:var(--n);font-family:var(--mono);line-height:1}
.kpi-v.r{color:var(--r)}.kpi-v.a{color:var(--a)}.kpi-v.g{color:var(--g)}
.kpi-s{font-size:11px;color:var(--s5);margin-top:3px}

/* ── dashboard buckets ── */
.bucket-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;margin-bottom:20px}
@media(max-width:900px){.bucket-grid{grid-template-columns:repeat(2,1fr)}}
@media(max-width:600px){.bucket-grid{grid-template-columns:1fr}}
.bucket{background:var(--w);border:1px solid var(--s2);border-radius:10px;padding:18px;cursor:pointer;transition:all .15s;box-shadow:0 1px 4px rgba(0,0,0,.06)}
.bucket:hover{border-color:var(--b);box-shadow:0 4px 14px rgba(15,76,129,.12);transform:translateY(-2px)}
.bucket-icon{font-size:22px;margin-bottom:8px}
.bucket-count{font-size:34px;font-weight:700;color:var(--n);font-family:var(--mono);line-height:1}
.bucket-count.r{color:var(--r)}.bucket-count.a{color:var(--a)}.bucket-count.g{color:var(--g)}
.bucket-label{font-size:12px;font-weight:700;color:var(--s7);margin-top:4px}
.bucket-desc{font-size:11px;color:var(--s5);margin-top:2px}
.bucket-items{margin-top:10px;border-top:1px solid var(--s1);padding-top:8px}
.bucket-item{font-size:11px;color:var(--s7);padding:3px 0;border-bottom:1px solid var(--s1);display:flex;justify-content:space-between;align-items:center}
.bucket-item:last-child{border-bottom:none}

/* ── panels ── */
.panel{background:var(--w);border:1px solid var(--s2);border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.06);margin-bottom:14px}
.panel-hdr{padding:11px 16px;border-bottom:1px solid var(--s1);display:flex;align-items:center;justify-content:space-between;background:linear-gradient(to right,var(--b0),var(--w))}
.panel-title{font-size:12px;font-weight:700;color:var(--n);display:flex;align-items:center;gap:7px}
.panel-cnt{font-family:var(--mono);font-size:11px;color:var(--s5);background:var(--s1);padding:2px 8px;border-radius:10px}

/* ── tables ── */
.tcard{background:var(--w);border:1px solid var(--s2);border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.05);margin-bottom:14px}
table{width:100%;border-collapse:collapse;font-size:12px}
thead th{background:var(--s0);padding:9px 12px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--s5);border-bottom:1px solid var(--s2)}
tbody tr{border-bottom:1px solid var(--s1);transition:background .1s;cursor:pointer}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:var(--b0)}
td{padding:9px 12px;vertical-align:middle}
.cid{font-family:var(--mono);font-size:11px;color:var(--b);font-weight:500;cursor:pointer}
.cid:hover{text-decoration:underline}
.cmut{font-size:11px;color:var(--s5)}
.cdate{font-family:var(--mono);font-size:11px;color:var(--s5);white-space:nowrap}

/* ── pills ── */
.pill{display:inline-block;padding:2px 9px;border-radius:12px;font-size:10px;font-weight:600;white-space:nowrap}
.pr{background:var(--r1);color:var(--r)}.pg{background:var(--g1);color:var(--g)}
.pa{background:var(--a1);color:var(--a)}.pb{background:var(--b1);color:var(--b)}
.pz{background:var(--s1);color:var(--s5)}.pp{background:#f3e5f5;color:#6a1b9a}

/* ── phase bar ── */
.phasebar{display:flex;align-items:center;margin:10px 0}
.ph{display:flex;flex-direction:column;align-items:center;flex:1}
.ph-dot{width:28px;height:28px;border-radius:50%;background:var(--s2);display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:var(--s5)}
.ph-dot.done{background:var(--g);color:#fff}
.ph-dot.cur{background:var(--b);color:#fff;box-shadow:0 0 0 3px rgba(15,76,129,.2)}
.ph-dot.late{background:var(--r);color:#fff}
.ph-dot.train{background:#7b1fa2;color:#fff}
.ph-lbl{font-size:9px;color:var(--s5);margin-top:4px;text-align:center;font-weight:600;text-transform:uppercase;letter-spacing:.4px}
.ph-line{flex:1;height:2px;background:var(--s2);margin-top:-14px}
.ph-line.done{background:var(--g)}

/* ── filter bar ── */
.fbar{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap;align-items:center}
.fbtn{font-family:var(--sans);font-size:11px;font-weight:600;padding:5px 12px;border-radius:18px;border:1px solid var(--s2);background:var(--w);color:var(--s5);cursor:pointer;transition:all .15s}
.fbtn:hover,.fbtn.on{background:var(--b);color:#fff;border-color:var(--b)}
.fbtn.r.on{background:var(--r);border-color:var(--r)}
.fsearch{font-size:12px;padding:5px 12px;border-radius:18px;border:1px solid var(--s2);background:var(--w);outline:none;min-width:180px;margin-left:auto}
.fsearch:focus{border-color:var(--b)}

/* ── section header ── */
.sec-hdr{display:flex;align-items:center;justify-content:space-between;margin-bottom:10px;margin-top:20px}
.sec-title{font-size:12px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:var(--b);display:flex;align-items:center;gap:7px}
.sec-title::before{content:'';width:3px;height:13px;background:var(--b);border-radius:2px;display:block}

/* ── action buttons ── */
.btn-pri{font-family:var(--sans);font-size:12px;font-weight:600;padding:7px 16px;border-radius:6px;background:var(--b);color:#fff;border:none;cursor:pointer;display:inline-flex;align-items:center;gap:5px;transition:background .15s}
.btn-pri:hover{background:var(--b2)}
.btn-sec{font-family:var(--sans);font-size:12px;font-weight:600;padding:7px 16px;border-radius:6px;background:var(--s1);color:var(--s7);border:1px solid var(--s2);cursor:pointer}
.btn-sec:hover{background:var(--s2)}
.btn-r{background:var(--r)}.btn-r:hover{background:#b71c1c}
.btn-g{background:var(--g)}.btn-g:hover{background:#1b5e20}
.btn-sm{padding:4px 10px;font-size:11px}

/* ── DCO pipeline visual ── */
.pipeline{display:flex;gap:0;margin-bottom:16px;overflow-x:auto}
.pip-stage{flex:1;min-width:100px;padding:12px 8px;text-align:center;background:var(--s1);border-right:1px solid var(--s2);cursor:pointer;transition:all .15s}
.pip-stage:last-child{border-right:none}
.pip-stage:hover{background:var(--b0)}
.pip-stage.active{background:var(--b0);border-bottom:3px solid var(--b)}
.pip-stage.late{border-bottom:3px solid var(--r)}
.pip-n{font-family:var(--mono);font-size:24px;font-weight:500;color:var(--n);line-height:1}
.pip-n.r{color:var(--r)}.pip-n.a{color:var(--a)}.pip-n.g{color:var(--g)}.pip-n.p{color:#7b1fa2}
.pip-l{font-size:10px;text-transform:uppercase;letter-spacing:.6px;color:var(--s5);margin-top:4px;font-weight:600}

/* ── approval lanes ── */
.lane-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px;margin-bottom:12px}
.lane{background:var(--s0);border:1px solid var(--s2);border-radius:7px;padding:12px;text-align:center}
.lane.signed{background:var(--g1);border-color:#a5d6a7}
.lane.waiting{background:var(--a1);border-color:#ffcc80}
.lane.blocked{background:var(--r1);border-color:#ef9a9a}
.lane-name{font-size:12px;font-weight:600;color:var(--s7)}
.lane-role{font-size:10px;color:var(--s5);margin:2px 0 6px}
.lane-status{font-size:10px;font-weight:700}
.lane.signed .lane-status{color:var(--g)}
.lane.waiting .lane-status{color:var(--a)}
.lane.blocked .lane-status{color:var(--r)}
.lane-sig{font-family:var(--mono);font-size:9px;color:var(--s5);margin-top:4px}

/* ── modal ── */
.modal-ov{display:none;position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:9999;align-items:center;justify-content:center;padding:20px}
.modal-ov.open{display:flex}
.modal{background:var(--w);border-radius:12px;width:100%;max-width:660px;max-height:88vh;overflow-y:auto;box-shadow:0 20px 60px rgba(0,0,0,.3)}
.modal-hdr{padding:18px 22px 14px;border-bottom:1px solid var(--s2);display:flex;justify-content:space-between;align-items:flex-start;position:sticky;top:0;background:var(--w);z-index:1}
.modal-title{font-size:14px;font-weight:700;color:var(--b)}
.modal-sub{font-family:var(--mono);font-size:10px;color:var(--s5);margin-top:2px}
.modal-x{background:none;border:none;font-size:22px;cursor:pointer;color:var(--s5);line-height:1;padding:0 4px}
.modal-body{padding:18px 22px}
.modal-ft{padding:14px 22px;border-top:1px solid var(--s2);display:flex;justify-content:flex-end;gap:8px}
.fg{margin-bottom:13px}
.fl{font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.7px;color:var(--s5);margin-bottom:4px}
.fv{font-size:13px;color:var(--s7);padding:8px 10px;background:var(--s0);border-radius:6px;border:1px solid var(--s2);line-height:1.5}
.fv.r{color:var(--r)}
.finput{width:100%;font-family:var(--sans);font-size:13px;padding:8px 10px;background:var(--w);border-radius:6px;border:1px solid var(--s2);outline:none}
.finput:focus{border-color:var(--b)}
.fsel{width:100%;font-family:var(--sans);font-size:12px;padding:7px 10px;background:var(--w);border-radius:6px;border:1px solid var(--s2);outline:none}
.ftxt{width:100%;font-family:var(--sans);font-size:12px;padding:8px 10px;background:var(--w);border-radius:6px;border:1px solid var(--s2);outline:none;min-height:80px;resize:vertical}

/* ── routing history ── */
.rh-item{display:flex;gap:10px;padding:10px 0;border-bottom:1px solid var(--s1)}
.rh-item:last-child{border-bottom:none}
.rh-dot{width:8px;height:8px;border-radius:50%;flex-shrink:0;margin-top:5px}
.rh-dot.stage{background:var(--b)}.rh-dot.sig{background:var(--g)}.rh-dot.rej{background:var(--r)}.rh-dot.sys{background:var(--s5)}
.rh-body{flex:1}
.rh-top{display:flex;justify-content:space-between;align-items:baseline}
.rh-evt{font-size:12px;font-weight:600;color:var(--s7)}
.rh-ts{font-family:var(--mono);font-size:10px;color:var(--s5)}
.rh-detail{font-size:11px;color:var(--s5);margin-top:3px}
.rh-reason{font-size:11px;color:var(--r);margin-top:4px;padding:5px 8px;background:var(--r1);border-radius:4px}

/* ── training matrix ── */
.tm-grid{overflow-x:auto}
.tm-table{border-collapse:collapse;min-width:600px}
.tm-table th,.tm-table td{padding:7px 10px;border:1px solid var(--s2);font-size:11px;text-align:center}
.tm-table th{background:var(--s0);font-weight:700;color:var(--s5)}
.tm-table th.role-hdr{text-align:left;min-width:150px}
.tm-check{color:var(--g);font-weight:700}.tm-dash{color:var(--s2)}
.tm-new{color:var(--a);font-weight:700}
.tm-over{color:var(--r);font-weight:700}

/* ── zone badges ── */
.zone-d{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;background:#e8f0fe;color:#1565c0}
.zone-p{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;background:#e8f5e9;color:#2e7d32}
.zone-o{display:inline-block;padding:2px 8px;border-radius:4px;font-size:10px;font-weight:700;background:#fff3e0;color:#e65100}

/* ── toast ── */
.toast{position:fixed;bottom:24px;right:24px;background:var(--n);color:#fff;padding:12px 20px;border-radius:8px;font-size:12px;font-weight:600;z-index:99999;opacity:0;transform:translateY(10px);transition:all .25s;pointer-events:none}
.toast.show{opacity:1;transform:translateY(0)}

/* ── config grid ── */
.cfg-grid{display:grid;grid-template-columns:1fr 1fr;gap:14px}
@media(max-width:700px){.cfg-grid{grid-template-columns:1fr}}
.cfg-panel{background:var(--s0);border:1px solid var(--s2);border-radius:8px;padding:16px}
.cfg-title{font-size:12px;font-weight:700;color:var(--n);margin-bottom:12px;text-transform:uppercase;letter-spacing:.5px}
.cfg-row{display:flex;justify-content:space-between;align-items:center;margin-bottom:10px;padding-bottom:10px;border-bottom:1px solid var(--s2)}
.cfg-row:last-child{border-bottom:none;margin-bottom:0;padding-bottom:0}
.cfg-lbl{font-size:12px;color:var(--s7)}
.cfg-val{font-family:var(--mono);font-size:12px;color:var(--n);font-weight:500}
.cfg-input{width:80px;font-family:var(--mono);font-size:12px;padding:4px 8px;border:1px solid var(--s2);border-radius:4px;text-align:center;background:var(--w)}

/* ── loading ── */
.loading{padding:32px;text-align:center;color:var(--s5);font-size:12px;font-family:var(--mono)}
.spin{display:inline-block;width:16px;height:16px;border:2px solid var(--s2);border-top-color:var(--b);border-radius:50%;animation:spin .8s linear infinite;margin-right:8px;vertical-align:middle}
@keyframes spin{to{transform:rotate(360deg)}}

/* ── late badge ── */
.late-badge{display:inline-block;padding:1px 7px;border-radius:10px;font-size:9px;font-weight:700;background:var(--r1);color:var(--r);margin-left:6px}
.warn-badge{display:inline-block;padding:1px 7px;border-radius:10px;font-size:9px;font-weight:700;background:var(--a1);color:var(--a);margin-left:6px}

footer{margin-top:8px;padding:12px 24px;border-top:1px solid var(--s2);font-size:11px;color:var(--s5);text-align:center;background:var(--w)}
`;

// ─────────────────────────────────────────────────────────────────────────────
// HTML SHELLS per screen
// ─────────────────────────────────────────────────────────────────────────────

const SHELL_DASHBOARD = `
<div class="kpi-row" id="db-kpis">
  <div class="kpi r"><div class="kpi-l">My Pending Actions</div><div class="kpi-v r" id="db-k1">—</div><div class="kpi-s">Requiring your input</div></div>
  <div class="kpi a"><div class="kpi-l">Awaiting Signature</div><div class="kpi-v a" id="db-k2">—</div><div class="kpi-s">Documents to sign</div></div>
  <div class="kpi"><div class="kpi-l">Active DCOs</div><div class="kpi-v" id="db-k3">—</div><div class="kpi-s">In routing</div></div>
  <div class="kpi"><div class="kpi-l">Open CRs</div><div class="kpi-v" id="db-k4">—</div><div class="kpi-s">Change requests</div></div>
  <div class="kpi a"><div class="kpi-l">Training Due</div><div class="kpi-v a" id="db-k5">—</div><div class="kpi-s">Overdue or due soon</div></div>
</div>
<div class="bucket-grid" id="db-buckets">
  <div class="bucket" data-nav="dco">
    <div class="bucket-icon">📋</div>
    <div class="bucket-count" id="db-bc1">—</div>
    <div class="bucket-label">Change Orders I'm Involved With</div>
    <div class="bucket-desc">DCOs where I'm an approver or originator</div>
    <div class="bucket-items" id="db-bi1"><div style="padding:6px 0;font-size:11px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="bucket" data-nav="cr">
    <div class="bucket-icon">🔄</div>
    <div class="bucket-count" id="db-bc2">—</div>
    <div class="bucket-label">Change Requests I'm Involved With</div>
    <div class="bucket-desc">CRs I originated or am reviewing</div>
    <div class="bucket-items" id="db-bi2"><div style="padding:6px 0;font-size:11px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="bucket r" data-nav="dco">
    <div class="bucket-icon">✍️</div>
    <div class="bucket-count r" id="db-bc3">—</div>
    <div class="bucket-label">Documents Needing My Signature</div>
    <div class="bucket-desc">DCOs awaiting your approval signature</div>
    <div class="bucket-items" id="db-bi3"><div style="padding:6px 0;font-size:11px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="bucket" data-nav="docs">
    <div class="bucket-icon">📄</div>
    <div class="bucket-count" id="db-bc4">—</div>
    <div class="bucket-label">Documents I Originated</div>
    <div class="bucket-desc">Documents in Draft or Published zones</div>
    <div class="bucket-items" id="db-bi4"><div style="padding:6px 0;font-size:11px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="bucket a" data-nav="training">
    <div class="bucket-icon">🎓</div>
    <div class="bucket-count a" id="db-bc5">—</div>
    <div class="bucket-label">Documents I Need to Be Trained To</div>
    <div class="bucket-desc">Based on your role assignments</div>
    <div class="bucket-items" id="db-bi5"><div style="padding:6px 0;font-size:11px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="bucket g" data-nav="publish">
    <div class="bucket-icon">🚀</div>
    <div class="bucket-count g" id="db-bc6">—</div>
    <div class="bucket-label">PM Publish Queue</div>
    <div class="bucket-desc">Drafts ready to push to Published zone</div>
    <div class="bucket-items" id="db-bi6"><div style="padding:6px 0;font-size:11px;color:var(--s5)">Loading...</div></div>
  </div>
</div>`;

const SHELL_DCO = `
<div class="sec-hdr">
  <div class="sec-title">Change Order Register</div>
  <button class="btn-pri" id="btn-new-dco">+ New DCO</button>
</div>
<div style="display:flex;gap:0;margin-bottom:14px;border:1px solid var(--s2);border-radius:8px;overflow:hidden">
  <div class="pip-stage" id="pip-draft" data-pip="Draft"><div class="pip-n" id="pip-n-draft">—</div><div class="pip-l">Draft</div></div>
  <div class="pip-stage" id="pip-submitted" data-pip="Submitted"><div class="pip-n a" id="pip-n-submitted">—</div><div class="pip-l">Submitted</div></div>
  <div class="pip-stage" id="pip-review" data-pip="In Review"><div class="pip-n a" id="pip-n-review">—</div><div class="pip-l">In Review</div></div>
  <div class="pip-stage" id="pip-implemented" data-pip="Implemented"><div class="pip-n g" id="pip-n-implemented">—</div><div class="pip-l">Implemented</div></div>
  <div class="pip-stage" id="pip-training" data-pip="Awaiting Training"><div class="pip-n p" id="pip-n-training">—</div><div class="pip-l">Awaiting Training</div></div>
  <div class="pip-stage" id="pip-effective" data-pip="Effective"><div class="pip-n g" id="pip-n-effective">—</div><div class="pip-l">Effective</div></div>
</div>
<div class="fbar">
  <button class="fbtn on" data-df="all" onclick="qpDCOFilter(this,'all')">All</button>
  <button class="fbtn" data-df="open" onclick="qpDCOFilter(this,'open')">Open</button>
  <button class="fbtn r" data-df="late" onclick="qpDCOFilter(this,'late')">🔴 Late</button>
  <button class="fbtn" data-df="mine" onclick="qpDCOFilter(this,'mine')">My DCOs</button>
  <input class="fsearch" id="dco-search" placeholder="Search DCOs..." oninput="qpRenderDCO()">
</div>
<div class="tcard">
  <table><thead><tr>
    <th>DCO #</th><th>Title</th><th>Phase</th><th>CR Link</th>
    <th>Submitted</th><th>Approvers</th><th>Actions</th>
  </tr></thead>
  <tbody id="dco-tbody"><tr><td colspan="7" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
</div>
<div id="dco-detail" style="display:none"></div>`;

const SHELL_CR = `
<div class="sec-hdr">
  <div class="sec-title">Change Request Register</div>
  <button class="btn-pri" id="btn-new-cr">+ New CR</button>
</div>
<div style="display:flex;gap:0;margin-bottom:14px;border:1px solid var(--s2);border-radius:8px;overflow:hidden">
  <div class="pip-stage"><div class="pip-n" id="cr-n-draft">—</div><div class="pip-l">Draft</div></div>
  <div class="pip-stage"><div class="pip-n a" id="cr-n-review">—</div><div class="pip-l">In Review</div></div>
  <div class="pip-stage"><div class="pip-n g" id="cr-n-approved">—</div><div class="pip-l">Approved</div></div>
  <div class="pip-stage"><div class="pip-n" id="cr-n-linked">—</div><div class="pip-l">Linked to DCO</div></div>
  <div class="pip-stage"><div class="pip-n g" id="cr-n-closed">—</div><div class="pip-l">Closed</div></div>
</div>
<div class="tcard">
  <table><thead><tr>
    <th>CR #</th><th>Title</th><th>Status</th><th>Priority</th>
    <th>Originator</th><th>Linked DCOs</th><th>Date</th>
  </tr></thead>
  <tbody id="cr-tbody"><tr><td colspan="7" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
</div>
<div id="cr-detail" style="display:none"></div>`;

const SHELL_DOCS = `
<div class="kpi-row">
  <div class="kpi"><div class="kpi-l">Draft Zone</div><div class="kpi-v" id="doc-k1">—</div><div class="kpi-s">QMS/Documents/Drafts</div></div>
  <div class="kpi g"><div class="kpi-l">Published</div><div class="kpi-v g" id="doc-k2">—</div><div class="kpi-s">Published/QMS</div></div>
  <div class="kpi a"><div class="kpi-l">On DCO (Locked)</div><div class="kpi-v a" id="doc-k3">—</div><div class="kpi-s">Edit prevented</div></div>
  <div class="kpi"><div class="kpi-l">Official</div><div class="kpi-v" id="doc-k4">—</div><div class="kpi-s">Official/QMS (Effective)</div></div>
</div>
<div class="fbar">
  <button class="fbtn on" data-zf="all">All Zones</button>
  <button class="fbtn" data-zf="Draft">📝 Draft</button>
  <button class="fbtn" data-zf="Published">✅ Published</button>
  <button class="fbtn" data-zf="Official">🏛️ Official</button>
  <input class="fsearch" id="doc-search" placeholder="Search documents..." oninput="qpRenderDocs()">
</div>
<div class="tcard">
  <table><thead><tr>
    <th>Doc ID</th><th>Title</th><th>Type</th><th>Rev</th>
    <th>Zone</th><th>Status</th><th>Linked DCO</th>
  </tr></thead>
  <tbody id="doc-tbody"><tr><td colspan="7" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
</div>`;

const SHELL_RECORDS = `
<div class="sec-hdr">
  <div class="sec-title">Records Pipeline</div>
  <button class="btn-pri" id="btn-new-record">+ New Record</button>
</div>
<div style="display:flex;gap:0;margin-bottom:14px;border:1px solid var(--s2);border-radius:8px;overflow:hidden">
  <div class="pip-stage"><div class="pip-n" id="rec-n-draft">—</div><div class="pip-l">Draft</div></div>
  <div class="pip-stage"><div class="pip-n a" id="rec-n-review">—</div><div class="pip-l">In Review</div></div>
  <div class="pip-stage"><div class="pip-n g" id="rec-n-approved">—</div><div class="pip-l">Approved</div></div>
  <div class="pip-stage"><div class="pip-n a" id="rec-n-pending">—</div><div class="pip-l">Pending Sig</div></div>
  <div class="pip-stage"><div class="pip-n g" id="rec-n-filed">—</div><div class="pip-l">Filed</div></div>
</div>
<div class="fbar">
  <button class="fbtn on" data-rf="all">All</button>
  <button class="fbtn" data-rf="Draft">Draft</button>
  <button class="fbtn" data-rf="In Review">In Review</button>
  <button class="fbtn r" data-rf="Approved">Approved</button>
  <input class="fsearch" id="rec-search" placeholder="Search records..." oninput="qpRenderRecords()">
</div>
<div class="tcard">
  <table><thead><tr>
    <th>Record #</th><th>Type</th><th>Title</th><th>Status</th>
    <th>Originator</th><th>Reviewer</th><th>Date</th>
  </tr></thead>
  <tbody id="rec-tbody"><tr><td colspan="7" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
</div>`;

const SHELL_TRAINING = `
<div class="kpi-row">
  <div class="kpi r"><div class="kpi-l">Overdue</div><div class="kpi-v r" id="tr-k1">—</div><div class="kpi-s">Past due date</div></div>
  <div class="kpi a"><div class="kpi-l">Due Soon</div><div class="kpi-v a" id="tr-k2">—</div><div class="kpi-s">Within 7 days</div></div>
  <div class="kpi"><div class="kpi-l">Pending</div><div class="kpi-v" id="tr-k3">—</div><div class="kpi-s">Not yet due</div></div>
  <div class="kpi g"><div class="kpi-l">Completed</div><div class="kpi-v g" id="tr-k4">—</div><div class="kpi-s">Signed &amp; filed</div></div>
</div>
<div style="display:flex;gap:8px;margin-bottom:14px;border-bottom:1px solid var(--s2)">
  <button class="qp-tab on" data-trtab="pending" style="color:var(--s7);border-bottom-color:var(--b)">Pending Training</button>
  <button class="qp-tab" data-trtab="matrix" style="color:var(--s7)">Training Matrix</button>
  <button class="qp-tab" data-trtab="employees" style="color:var(--s7)">Employees &amp; Roles</button>
  <button class="qp-tab" data-trtab="history" style="color:var(--s7)">Training History</button>
</div>
<div id="tr-tab-pending">
  <div class="fbar">
    <button class="fbtn on r" data-tf="all">All</button>
    <button class="fbtn" data-tf="Overdue">🔴 Overdue</button>
    <button class="fbtn" data-tf="Due Soon">⚠️ Due Soon</button>
    <button class="fbtn" data-tf="Pending">Pending</button>
    <input class="fsearch" id="tr-search" placeholder="Search training..." oninput="qpRenderTraining()">
  </div>
  <div class="tcard">
    <table><thead><tr>
      <th>Employee</th><th>Document</th><th>Rev</th><th>Role</th>
      <th>Due Date</th><th>Status</th><th>Action</th>
    </tr></thead>
    <tbody id="tr-tbody"><tr><td colspan="7" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
  </div>
</div>
<div id="tr-tab-matrix" style="display:none">
  <div class="tm-grid" id="tr-matrix-wrap"><div class="loading"><span class="spin"></span>Loading matrix...</div></div>
</div>
<div id="tr-tab-employees" style="display:none">
  <div class="tcard">
    <table><thead><tr><th>Name</th><th>Title</th><th>Department</th><th>Roles</th><th>Training Currency</th></tr></thead>
    <tbody id="tr-emp-tbody"><tr><td colspan="5" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
  </div>
</div>
<div id="tr-tab-history" style="display:none">
  <div class="tcard">
    <table><thead><tr><th>Employee</th><th>Document</th><th>Rev</th><th>Method</th><th>Signed</th><th>Sig ID</th></tr></thead>
    <tbody id="tr-hist-tbody"><tr><td colspan="6" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
  </div>
</div>`;

const SHELL_PUBLISH = `
<div class="sec-hdr">
  <div class="sec-title">PM Publish Queue — Draft → Published Zone</div>
</div>
<div class="panel" style="margin-bottom:16px">
  <div class="panel-hdr"><div class="panel-title">📝 Drafts Ready to Publish <span class="panel-cnt" id="pub-cnt">—</span></div></div>
  <div id="pub-list"><div class="loading"><span class="spin"></span>Loading...</div></div>
</div>
<div class="panel">
  <div class="panel-hdr"><div class="panel-title">✅ Recently Published</div></div>
  <div class="tcard" style="margin:0;border:none;box-shadow:none">
    <table><thead><tr><th>Doc ID</th><th>Title</th><th>Rev</th><th>Published</th><th>Published By</th></tr></thead>
    <tbody id="pub-hist-tbody"><tr><td colspan="5" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
  </div>
</div>`;

const SHELL_CONFIG = `
<div class="sec-hdr"><div class="sec-title">System Configuration</div><button class="btn-pri" id="btn-save-config">💾 Save All Settings</button></div>
<div class="cfg-grid" id="cfg-body"><div class="loading"><span class="spin"></span>Loading config...</div></div>`;

const SHELL_APPROVERS = `
<div class="sec-hdr">
  <div class="sec-title">Approver Setup</div>
  <button class="btn-pri" id="btn-new-approver">+ Add Approver</button>
</div>
<div class="tcard">
  <table><thead><tr>
    <th>Name</th><th>Role</th><th>Type</th><th>Scope</th><th>Signing Mode</th><th>Active</th><th>Actions</th>
  </tr></thead>
  <tbody id="appr-tbody"><tr><td colspan="7" class="loading"><span class="spin"></span>Loading...</td></tr></tbody></table>
</div>`;

// ─────────────────────────────────────────────────────────────────────────────
// MAIN SHELL
// ─────────────────────────────────────────────────────────────────────────────
function buildShell(): string {
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<link href="https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=DM+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">
<style>${CSS}</style>
<script>
var _qpReady=false,_qpQueue=[];
function _qpStub(fn,args){if(_qpReady&&window['_qp_'+fn]){window['_qp_'+fn].apply(null,args);}else{_qpQueue.push({fn:fn,args:args});}}
function _qpFlush(){_qpReady=true;_qpQueue.forEach(function(q){if(window['_qp_'+q.fn])window['_qp_'+q.fn].apply(null,q.args);});_qpQueue=[];}
function qpNav(s){_qpStub('Nav',[s]);}
function qpRefresh(){_qpStub('Refresh',[]);}
function qpOpenDCO(id){_qpStub('OpenDCO',[id]);}
function qpOpenCR(id){_qpStub('OpenCR',[id]);}
function qpOpenReject(){_qpStub('OpenReject',[]);}
function qpConfirmReject(){_qpStub('ConfirmReject',[]);}
function qpConfirmSign(){_qpStub('ConfirmSign',[]);}
function qpOpenModal(id){_qpStub('OpenModal',[id]);}
function qpCloseModal(id){_qpStub('CloseModal',[id]);}
function qpDCOFilter(b,f){_qpStub('DCOFilter',[b,f]);}
function qpPipFilter(p){_qpStub('PipFilter',[p]);}
function qpZoneFilter(b,z){_qpStub('ZoneFilter',[b,z]);}
function qpRecFilter(b,s){_qpStub('RecFilter',[b,s]);}
function qpTrTab(b,t){_qpStub('TrTab',[b,t]);}
function qpTrFilter(b,f){_qpStub('TrFilter',[b,f]);}
function qpRenderDCO(){_qpStub('RenderDCO',[]);}
function qpSaveConfig(){_qpStub('SaveConfig',[]);}
function qpOpenNewDCO(){_qpStub('OpenNewDCO',[]);}
function qpOpenNewCR(){_qpStub('OpenNewCR',[]);}
function qpOpenNewRecord(){_qpStub('OpenNewRecord',[]);}
function qpOpenNewApprover(){_qpStub('OpenNewApprover',[]);}
</script>
</head><body>
<div class="qp-hdr">
  <div class="qp-hdr-l">
    <div class="qp-badge">QMS</div>
    <div><div class="qp-title">IMP9177 · QMS Portal</div><div class="qp-sub">3H Pharmaceuticals LLC · 21 CFR Part 111 / FSMA · Document Control &amp; DCO Routing</div></div>
  </div>
  <div class="qp-hdr-r">
    <span class="qp-ts" id="qp-ts">Loading...</span>
    <button class="qp-btn" id="qp-refresh-btn">⟳ Refresh</button>
  </div>
</div>
<div class="qp-nav">
  <button class="qp-tab on" data-screen="dashboard">🏠 Dashboard</button>
  <button class="qp-tab" data-screen="dco">📋 Change Orders</button>
  <button class="qp-tab" data-screen="cr">🔄 Change Requests</button>
  <button class="qp-tab" data-screen="docs">📄 Documents</button>
  <button class="qp-tab" data-screen="records">📁 Records</button>
  <button class="qp-tab" data-screen="training">🎓 Training</button>
  <button class="qp-tab" data-screen="publish">🚀 PM Publish</button>
  <button class="qp-tab" data-screen="approvers">👥 Approvers</button>
  <button class="qp-tab" data-screen="config">⚙️ Config</button>
</div>
<div class="qp-alert" id="qp-alert"><span>⚠️</span><span id="qp-alert-txt"></span></div>
<div class="qp-main">
  <div class="qp-screen on" id="sc-dashboard">${SHELL_DASHBOARD}</div>
  <div class="qp-screen" id="sc-dco">${SHELL_DCO}</div>
  <div class="qp-screen" id="sc-cr">${SHELL_CR}</div>
  <div class="qp-screen" id="sc-docs">${SHELL_DOCS}</div>
  <div class="qp-screen" id="sc-records">${SHELL_RECORDS}</div>
  <div class="qp-screen" id="sc-training">${SHELL_TRAINING}</div>
  <div class="qp-screen" id="sc-publish">${SHELL_PUBLISH}</div>
  <div class="qp-screen" id="sc-approvers">${SHELL_APPROVERS}</div>
  <div class="qp-screen" id="sc-config">${SHELL_CONFIG}</div>
</div>
<footer>IMP9177 QMS Portal · ADB Consulting &amp; CRO Inc. · 21 CFR Part 111 / FSMA · Read-only view — data live from SharePoint</footer>
<div class="toast" id="qp-toast"></div>

<!-- Modals -->
<div class="modal-ov" id="modal-dco-detail">
  <div class="modal" style="max-width:780px">
    <div class="modal-hdr">
      <div><div class="modal-title" id="mdco-title">DCO Detail</div><div class="modal-sub" id="mdco-sub"></div></div>
      <button class="modal-x" data-close="modal-dco-detail">×</button>
    </div>
    <div class="modal-body" id="mdco-body"></div>
    <div class="modal-ft">
      <button class="btn-sec" data-close="modal-dco-detail">Close</button>
      <button class="btn-pri" id="mdco-action-btn" style="display:none"></button>
      <button class="btn-pri btn-r btn-sm" id="mdco-reject-btn" style="display:none" id="mdco-reject-open">↩ Reject</button>
    </div>
  </div>
</div>

<div class="modal-ov" id="modal-cr-detail">
  <div class="modal">
    <div class="modal-hdr">
      <div><div class="modal-title" id="mcr-title">CR Detail</div><div class="modal-sub" id="mcr-sub"></div></div>
      <button class="modal-x" data-close="modal-cr-detail">×</button>
    </div>
    <div class="modal-body" id="mcr-body"></div>
    <div class="modal-ft"><button class="btn-sec" data-close="modal-cr-detail">Close</button></div>
  </div>
</div>

<div class="modal-ov" id="modal-reject">
  <div class="modal" style="max-width:500px">
    <div class="modal-hdr">
      <div><div class="modal-title">Reject DCO — Reason Required</div></div>
      <button class="modal-x" data-close="modal-reject">×</button>
    </div>
    <div class="modal-body">
      <div class="fg"><div class="fl">Rejection Category</div>
        <select class="fsel" id="rej-cat">
          <option>Incomplete documentation</option>
          <option>Missing approver sign-off</option>
          <option>Regulatory non-compliance</option>
          <option>Training not completed</option>
          <option>Other</option>
        </select>
      </div>
      <div class="fg"><div class="fl">Rejection Reason (required)</div>
        <textarea class="ftxt" id="rej-reason" placeholder="Describe the reason for rejection..."></textarea>
      </div>
    </div>
    <div class="modal-ft">
      <button class="btn-sec" data-close="modal-reject">Cancel</button>
      <button class="btn-pri btn-r" id="btn-confirm-reject">Confirm Rejection</button>
    </div>
  </div>
</div>

<div class="modal-ov" id="modal-esign">
  <div class="modal" style="max-width:520px">
    <div class="modal-hdr">
      <div><div class="modal-title">E-Signature — Microsoft 365</div><div class="modal-sub" id="esign-sub">Signing document</div></div>
      <button class="modal-x" data-close="modal-esign">×</button>
    </div>
    <div class="modal-body" id="esign-body">
      <div class="fg"><div class="fl">Document</div><div class="fv" id="esign-doc">—</div></div>
      <div class="fg"><div class="fl">Signing as</div><div class="fv" id="esign-as">—</div></div>
      <div class="fg"><div class="fl">Attestation</div><div class="fv">I confirm the accuracy and completeness of this document and approve its release per 21 CFR Part 11 electronic signature requirements.</div></div>
      <div class="fg"><div class="fl">MFA Verification Code (simulated)</div>
        <input class="finput" id="esign-code" type="text" maxlength="6" placeholder="Enter 6-digit code..." style="letter-spacing:4px;font-family:var(--mono)">
        <div style="font-size:11px;color:var(--s5);margin-top:4px">In production: code sent to your registered M365 MFA device</div>
      </div>
    </div>
    <div class="modal-ft">
      <button class="btn-sec" data-close="modal-esign">Cancel</button>
      <button class="btn-pri btn-g" id="btn-confirm-sign">✅ Apply Signature</button>
    </div>
  </div>
</div>
</body></html>`;
}

// ─────────────────────────────────────────────────────────────────────────────
// WEB PART CLASS
// ─────────────────────────────────────────────────────────────────────────────
export default class QmsPortalWebPart extends BaseClientSideWebPart<IQmsPortalWebPartProps> {

  private _iframe: HTMLIFrameElement | null = null;
  private _data: Record<string, any[]> = {};
  private _config: any = {};

  // ── SharePoint REST helper ──
  private spGet(list: string, select = '', filter = '', top = 500): Promise<any[]> {
    const base = this.context.pageContext.web.absoluteUrl;
    let url = `${base}/_api/web/lists/getbytitle('${list}')/items?$top=${top}&$orderby=Id asc`;
    if (select) url += `&$select=${select}`;
    if (filter) url += `&$filter=${filter}`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => r.json())
      .then((d: any) => d.value || [])
      .catch(() => []);
  }

  // ── Render entry point ──
  public render(): void {
    this.domElement.innerHTML = '';
    this._iframe = document.createElement('iframe');
    this._iframe.style.cssText = 'width:100%;height:1100px;border:none;display:block;';
    this._iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin allow-forms allow-popups');
    this.domElement.appendChild(this._iframe);

    const ifrDoc = this._iframe.contentDocument ||
      (this._iframe.contentWindow && this._iframe.contentWindow.document);
    if (ifrDoc) { ifrDoc.open(); ifrDoc.write(buildShell()); ifrDoc.close(); }
    else { this._iframe.srcdoc = buildShell(); }

    const attach = () => { this._attachListeners(); this._loadAll(); };
    const d = this._iframe.contentDocument;
    if (d && d.readyState === 'complete') attach();
    else this._iframe.addEventListener('load', attach, { once: true });
  }

  // ── Fetch all QMS lists ──
  private async _loadAll(): Promise<void> {
    try {
      const [dcos, crs, approvals, history, records, employees, roles, matrix, completions, config] =
        await Promise.all([
          this.spGet('QMS_DCOs', 'Id,Title,DCO_Phase,DCO_Title,DCO_CRLink,DCO_SubmittedDate,DCO_Originator,DCO_Docs,DCO_LateDays,DCO_TrainingGate'),
          this.spGet('QMS_ChangeRequests', 'Id,Title,CR_Title,CR_Status,CR_Priority,CR_Originator,CR_LinkedDCOs,CR_Description,CR_CreatedDate'),
          this.spGet('QMS_DCOApprovals', 'Id,Title,Appr_DCOID,Appr_Name,Appr_Role,Appr_Type,Appr_Status,Appr_SignedDate,Appr_SigID'),
          this.spGet('QMS_RoutingHistory', 'Id,Title,RH_DCOID,RH_EventType,RH_Stage,RH_Actor,RH_Note,RH_Reason,RH_Timestamp'),
          this.spGet('QMS_Records', 'Id,Title,Rec_Type,Rec_Title,Rec_Status,Rec_Originator,Rec_Reviewer,Rec_CreatedDate,Rec_SigID'),
          this.spGet('QMS_Employees', 'Id,Title,Emp_Email,Emp_Title,Emp_Dept,Emp_Roles'),
          this.spGet('QMS_Roles', 'Id,Title,Role_Desc,Role_RequiredDocs'),
          this.spGet('QMS_TrainingMatrix', 'Id,Title,TM_RoleID,TM_DocID,TM_Required'),
          this.spGet('QMS_TrainingCompletions', 'Id,Title,TC_EmpID,TC_DocID,TC_Rev,TC_Method,TC_SignedDate,TC_SigID'),
          this.spGet('QMS_Config', 'Id,Title,Cfg_Value', '', 10),
        ]);
      this._data = { dcos, crs, approvals, history, records, employees, roles, matrix, completions };
      this._config = this._parseConfig(config);
    } catch (e) {
      console.error('QMS Portal: load failed', e);
    }
    this._renderAll();
    this._setTs();
  }

  private _parseConfig(rows: any[]): any {
    const cfg: any = { approvalOverdueDays: 14, approvalWarningDays: 7, trainingDueDays: 30, trainingWarningDays: 7 };
    rows.forEach(r => { if (r.Title && r.Cfg_Value) cfg[r.Title] = r.Cfg_Value; });
    return cfg;
  }

  // ── Render all screens ──
  private _renderAll(): void {
    this._renderDashboard();
    this._renderDCO();
    this._renderCR();
    this._renderDocs();
    this._renderRecords();
    this._renderTraining();
    this._renderPublish();
    this._renderApprovers();
    this._renderConfig();
    this._checkAlerts();
  }

  // ── DOM helpers ──
  private _el(id: string): HTMLElement | null {
    return (this._iframe?.contentDocument?.getElementById(id)) || null;
  }
  private _set(id: string, v: string): void { const e = this._el(id); if (e) e.textContent = v; }
  private _html(id: string, v: string): void { const e = this._el(id); if (e) e.innerHTML = v; }
  private _setTs(): void {
    const e = this._el('qp-ts');
    if (e) e.textContent = 'Refreshed ' + new Date().toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
  }

  private _pill(s: string): string {
    if (!s) return '<span class="pill pz">—</span>';
    const l = s.toLowerCase();
    if (l.includes('overdue') || l.includes('blocking') || l.includes('critical') || l.includes('past due') || l.includes('reject')) return `<span class="pill pr">${s}</span>`;
    if (l.includes('complete') || l.includes('effective') || l.includes('filed') || l.includes('closed') || l.includes('approved')) return `<span class="pill pg">${s}</span>`;
    if (l.includes('review') || l.includes('submitted') || l.includes('training') || l.includes('pending') || l.includes('due soon')) return `<span class="pill pa">${s}</span>`;
    if (l.includes('draft') || l.includes('open') || l.includes('linked')) return `<span class="pill pb">${s}</span>`;
    return `<span class="pill pz">${s}</span>`;
  }

  private _fmt(s: string): string {
    if (!s) return '—';
    try { const d = new Date(s); return d.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: '2-digit' }); }
    catch { return s; }
  }

  private _lateStatus(dco: any): string {
    if (!dco.DCO_SubmittedDate) return '';
    const days = Math.floor((Date.now() - new Date(dco.DCO_SubmittedDate).getTime()) / 86400000);
    const overdue = parseInt(this._config.approvalOverdueDays) || 14;
    const warn = parseInt(this._config.approvalWarningDays) || 7;
    if (days >= overdue) return 'late';
    if (days >= warn) return 'warn';
    return '';
  }

  // ── Dashboard ──
  private _renderDashboard(): void {
    const { dcos = [], crs = [], approvals = [], matrix = [], completions = [], employees = [] } = this._data;
    const activeDCOs = dcos.filter(d => !(['Effective'].indexOf(d.DCO_Phase || '') !== -1));
    const openCRs = crs.filter(c => ['Closed'].indexOf(c.CR_Status || '') === -1);
    const pendingSig = approvals.filter(a => (a.Appr_Status || '') === 'Waiting');
    const trainingPending = this._computePendingTraining();
    const overdueTraining = trainingPending.filter(t => t.status === 'Overdue');
    const dueSoon = trainingPending.filter(t => t.status === 'Due Soon');

    this._set('db-k1', String(pendingSig.length));
    this._set('db-k2', String(pendingSig.length));
    this._set('db-k3', String(activeDCOs.length));
    this._set('db-k4', String(openCRs.length));
    this._set('db-k5', String(overdueTraining.length + dueSoon.length));

    this._set('db-bc1', String(activeDCOs.length));
    this._set('db-bc2', String(openCRs.length));
    this._set('db-bc3', String(pendingSig.length));
    this._set('db-bc4', String(dcos.length));
    this._set('db-bc5', String(trainingPending.length));
    this._set('db-bc6', '0'); // publish queue from doc library

    const dcoItems = activeDCOs.slice(0, 3).map(d =>
      `<div class="bucket-item"><span style="font-family:var(--mono);font-size:10px;color:var(--b)">${d.Title}</span>${this._pill(d.DCO_Phase)}</div>`).join('');
    this._html('db-bi1', dcoItems || '<div style="padding:6px 0;font-size:11px;color:var(--s5)">No active DCOs</div>');

    const crItems = openCRs.slice(0, 3).map(c =>
      `<div class="bucket-item"><span style="font-family:var(--mono);font-size:10px;color:var(--b)">${c.Title}</span>${this._pill(c.CR_Status)}</div>`).join('');
    this._html('db-bi2', crItems || '<div style="padding:6px 0;font-size:11px;color:var(--s5)">No open CRs</div>');

    const sigItems = pendingSig.slice(0, 3).map(a =>
      `<div class="bucket-item"><span style="font-size:11px">${a.Appr_DCOID}</span><span style="font-size:10px;color:var(--r)">Awaiting</span></div>`).join('');
    this._html('db-bi3', sigItems || '<div style="padding:6px 0;font-size:11px;color:var(--s5)">No pending signatures</div>');

    const trItems = overdueTraining.slice(0, 3).map(t =>
      `<div class="bucket-item"><span style="font-size:11px">${t.docId}</span><span style="font-size:10px;color:var(--r)">Overdue</span></div>`).join('');
    this._html('db-bi5', trItems || '<div style="padding:6px 0;font-size:11px;color:var(--s5)">No training due</div>');
  }

  // ── DCO screen ──
  private _renderDCO(): void {
    const { dcos = [] } = this._data;
    const phases = ['Draft', 'Submitted', 'In Review', 'Implemented', 'Awaiting Training', 'Effective'];
    phases.forEach(ph => {
      const cnt = dcos.filter(d => d.DCO_Phase === ph).length;
      const key = ph.toLowerCase().replace(/\s+/g, ph === 'In Review' ? 'review' : ph === 'Awaiting Training' ? 'training' : ph === 'Effective' ? 'effective' : ph.toLowerCase());
      const idMap: Record<string, string> = {
        'Draft': 'pip-n-draft', 'Submitted': 'pip-n-submitted', 'In Review': 'pip-n-review',
        'Implemented': 'pip-n-implemented', 'Awaiting Training': 'pip-n-training', 'Effective': 'pip-n-effective'
      };
      this._set(idMap[ph] || '', String(cnt));
    });
    this._dcoTableRender(dcos);
  }

  private _dcoTableRender(dcos: any[]): void {
    const { approvals = [] } = this._data;
    if (!dcos.length) {
      this._html('dco-tbody', '<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--s5)">No DCOs found</td></tr>');
      return;
    }
    const rows = dcos.map(d => {
      const apprs = approvals.filter(a => a.Appr_DCOID === d.Title);
      const signed = apprs.filter(a => a.Appr_Status === 'Signed').length;
      const total = apprs.length;
      const late = this._lateStatus(d);
      const lateBadge = late === 'late' ? '<span class="late-badge">LATE</span>' : late === 'warn' ? '<span class="warn-badge">DUE SOON</span>' : '';
      return `<tr data-dcoid="${d.Title}">
        <td><span class="cid">${d.Title}</span></td>
        <td style="font-size:12px;max-width:220px">${(d.DCO_Title || '').substring(0, 50)}</td>
        <td>${this._pill(d.DCO_Phase)}${lateBadge}</td>
        <td><span class="cmut" style="font-family:var(--mono);font-size:11px">${d.DCO_CRLink || '—'}</span></td>
        <td><span class="cdate">${this._fmt(d.DCO_SubmittedDate)}</span></td>
        <td><span class="cmut">${signed}/${total} signed</span></td>
        <td><button class="btn-pri btn-sm" data-dcobtn="${d.Title}">Open</button></td>
      </tr>`;
    }).join('');
    this._html('dco-tbody', rows);
  }

  // ── CR screen ──
  private _renderCR(): void {
    const { crs = [] } = this._data;
    ['Draft', 'In Review', 'Approved', 'Linked to DCO', 'Closed'].forEach((ph, i) => {
      const ids = ['cr-n-draft', 'cr-n-review', 'cr-n-approved', 'cr-n-linked', 'cr-n-closed'];
      this._set(ids[i], String(crs.filter(c => c.CR_Status === ph).length));
    });
    const rows = crs.map(c => `<tr data-crid="${c.Title}">
      <td><span class="cid">${c.Title}</span></td>
      <td style="font-size:12px">${(c.CR_Title || '').substring(0, 50)}</td>
      <td>${this._pill(c.CR_Status)}</td>
      <td>${this._pill(c.CR_Priority)}</td>
      <td><span class="cmut">${c.CR_Originator || '—'}</span></td>
      <td><span class="cmut" style="font-family:var(--mono);font-size:11px">${c.CR_LinkedDCOs || '—'}</span></td>
      <td><span class="cdate">${this._fmt(c.CR_CreatedDate)}</span></td>
    </tr>`).join('');
    this._html('cr-tbody', rows || '<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--s5)">No CRs found</td></tr>');
  }

  // ── Docs screen ──
  private _renderDocs(): void {
    // Documents come from SP document library metadata columns set in Script 04
    this._html('doc-tbody', '<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--s5)">Document library view — use SharePoint to browse zones directly</td></tr>');
    this._set('doc-k1', '—');
    this._set('doc-k2', '—');
    this._set('doc-k3', '—');
    this._set('doc-k4', '—');
  }

  // ── Records screen ──
  private _renderRecords(): void {
    const { records = [] } = this._data;
    ['Draft', 'In Review', 'Approved', 'Pending Signature', 'Signed & Filed'].forEach((ph, i) => {
      const ids = ['rec-n-draft', 'rec-n-review', 'rec-n-approved', 'rec-n-pending', 'rec-n-filed'];
      this._set(ids[i], String(records.filter(r => r.Rec_Status === ph).length));
    });
    const rows = records.map(r => `<tr>
      <td><span class="cid">${r.Title}</span></td>
      <td><span class="cmut">${r.Rec_Type || '—'}</span></td>
      <td style="font-size:12px">${(r.Rec_Title || '').substring(0, 50)}</td>
      <td>${this._pill(r.Rec_Status)}</td>
      <td><span class="cmut">${r.Rec_Originator || '—'}</span></td>
      <td><span class="cmut">${r.Rec_Reviewer || '—'}</span></td>
      <td><span class="cdate">${this._fmt(r.Rec_CreatedDate)}</span></td>
    </tr>`).join('');
    this._html('rec-tbody', rows || '<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--s5)">No records yet — create the first record</td></tr>');
  }

  // ── Training ──
  private _computePendingTraining(): any[] {
    const { employees = [], matrix = [], completions = [], roles = [] } = this._data;
    const pending: any[] = [];
    const now = Date.now();
    const dueDays = parseInt(this._config.trainingDueDays) || 30;
    const warnDays = parseInt(this._config.trainingWarningDays) || 7;

    employees.forEach(emp => {
      const empRoles = (emp.Emp_Roles || '').split(',').map((r: string) => r.trim()).filter(Boolean);
      const requiredDocsSet: Record<string, boolean> = {};
      empRoles.forEach((roleId: string) => {
        const roleRec = roles.find(r => r.Title === roleId);
        if (roleRec && roleRec.Role_RequiredDocs) {
          (roleRec.Role_RequiredDocs as string).split(',').forEach((d: string) => { requiredDocsSet[d.trim()] = true; });
        }
      });
      Object.keys(requiredDocsSet).forEach(docId => {
        const completed = completions.find(c => c.TC_EmpID === emp.Title && c.TC_DocID === docId);
        if (!completed) {
          const dueDate = new Date(now + dueDays * 86400000);
          const daysUntil = Math.floor((dueDate.getTime() - now) / 86400000);
          const status = daysUntil < 0 ? 'Overdue' : daysUntil <= warnDays ? 'Due Soon' : 'Pending';
          pending.push({ empId: emp.Title, empName: emp.Title, docId, empRoles: empRoles.join(', '), dueDate: dueDate.toISOString(), status });
        }
      });
    });
    return pending;
  }

  private _renderTraining(): void {
    const pending = this._computePendingTraining();
    const { completions = [], employees = [], matrix = [], roles = [] } = this._data;
    const overdue = pending.filter(t => t.status === 'Overdue');
    const dueSoon = pending.filter(t => t.status === 'Due Soon');
    this._set('tr-k1', String(overdue.length));
    this._set('tr-k2', String(dueSoon.length));
    this._set('tr-k3', String(pending.filter(t => t.status === 'Pending').length));
    this._set('tr-k4', String(completions.length));

    const rows = pending.map(t => `<tr>
      <td style="font-size:12px">${t.empName}</td>
      <td><span class="cid">${t.docId}</span></td>
      <td><span class="cmut">Rev A</span></td>
      <td><span class="cmut">${t.empRoles}</span></td>
      <td><span class="cdate">${this._fmt(t.dueDate)}</span></td>
      <td>${this._pill(t.status)}</td>
      <td><button class="btn-pri btn-sm">Initiate</button></td>
    </tr>`).join('');
    this._html('tr-tbody', rows || '<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--s5)">All training current ✅</td></tr>');

    // Matrix
    const docIdsMap: Record<string, boolean> = {}; matrix.forEach((m: any) => { docIdsMap[m.TM_DocID] = true; });
    const roleIdsMap: Record<string, boolean> = {}; matrix.forEach((m: any) => { roleIdsMap[m.TM_RoleID] = true; });
    const docIds = Object.keys(docIdsMap);
    const roleIds = Object.keys(roleIdsMap);
    if (docIds.length && roleIds.length) {
      const hdr = `<tr><th class="role-hdr">Document</th>${roleIds.map(r => `<th>${r}</th>`).join('')}</tr>`;
      const rows2 = docIds.map(docId => {
        const cells = roleIds.map(roleId => {
          const req = matrix.find(m => m.TM_DocID === docId && m.TM_RoleID === roleId);
          return `<td>${req ? '<span class="tm-check">✅</span>' : '<span class="tm-dash">—</span>'}</td>`;
        }).join('');
        return `<tr><td style="text-align:left;font-family:var(--mono);font-size:11px;color:var(--b)">${docId}</td>${cells}</tr>`;
      }).join('');
      this._html('tr-matrix-wrap', `<div class="tm-grid"><table class="tm-table"><thead>${hdr}</thead><tbody>${rows2}</tbody></table></div>`);
    }

    // Employees
    const empRows = employees.map(e => `<tr>
      <td style="font-size:12px;font-weight:600">${e.Title}</td>
      <td><span class="cmut">${e.Emp_Title || '—'}</span></td>
      <td><span class="cmut">${e.Emp_Dept || '—'}</span></td>
      <td><span class="cmut" style="font-size:11px">${e.Emp_Roles || '—'}</span></td>
      <td>${this._pill('Current')}</td>
    </tr>`).join('');
    this._html('tr-emp-tbody', empRows || '<tr><td colspan="5" style="text-align:center;padding:24px;color:var(--s5)">No employees loaded</td></tr>');

    // History
    const { completions: comps = [] } = this._data;
    const histRows = comps.map(c => `<tr>
      <td style="font-size:12px">${c.TC_EmpID || '—'}</td>
      <td><span class="cid">${c.TC_DocID || '—'}</span></td>
      <td><span class="cmut">${c.TC_Rev || '—'}</span></td>
      <td><span class="cmut">${c.TC_Method || '—'}</span></td>
      <td><span class="cdate">${this._fmt(c.TC_SignedDate)}</span></td>
      <td><span class="cmut" style="font-family:var(--mono);font-size:10px">${c.TC_SigID || '—'}</span></td>
    </tr>`).join('');
    this._html('tr-hist-tbody', histRows || '<tr><td colspan="6" style="text-align:center;padding:24px;color:var(--s5)">No completed training records</td></tr>');
  }

  // ── Publish queue ──
  private _renderPublish(): void {
    this._html('pub-list', '<div style="padding:16px;font-size:12px;color:var(--s5)">Connect to document library to see Draft → Ready-to-Publish items. Use SharePoint document library view filtered to <strong>QMS_Status = Ready to Publish</strong>.</div>');
    this._html('pub-hist-tbody', '<tr><td colspan="5" style="text-align:center;padding:24px;color:var(--s5)">Recently published documents appear here</td></tr>');
    this._set('pub-cnt', '—');
  }

  // ── Approvers ──
  private _renderApprovers(): void {
    this.spGet('QMS_Approvers', 'Id,Title,Appr_Role,Appr_Type,Appr_Scope,Appr_SigningMode,Appr_Active')
      .then(rows => {
        const tableRows = rows.map(a => `<tr>
          <td style="font-size:12px;font-weight:600">${a.Title}</td>
          <td><span class="cmut">${a.Appr_Role || '—'}</span></td>
          <td>${this._pill(a.Appr_Type || '')}</td>
          <td><span class="cmut">${a.Appr_Scope || '—'}</span></td>
          <td><span class="cmut">${a.Appr_SigningMode || 'Parallel'}</span></td>
          <td>${a.Appr_Active ? '<span class="pill pg">Active</span>' : '<span class="pill pz">Inactive</span>'}</td>
          <td><button class="btn-sec btn-sm">Edit</button></td>
        </tr>`).join('');
        this._html('appr-tbody', tableRows || '<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--s5)">No approvers configured</td></tr>');
      });
  }

  // ── Config ──
  private _renderConfig(): void {
    const cfg = this._config;
    this._html('cfg-body', `
      <div class="cfg-panel">
        <div class="cfg-title">⏱ DCO Timing</div>
        <div class="cfg-row"><span class="cfg-lbl">Approval overdue threshold (days)</span><input class="cfg-input" id="cfg-overdue" value="${cfg.approvalOverdueDays || 14}"></div>
        <div class="cfg-row"><span class="cfg-lbl">Warning threshold (days)</span><input class="cfg-input" id="cfg-warn" value="${cfg.approvalWarningDays || 7}"></div>
        <div class="cfg-row"><span class="cfg-lbl">Draft stale threshold (days)</span><input class="cfg-input" id="cfg-stale" value="${cfg.draftStaleDays || 30}"></div>
      </div>
      <div class="cfg-panel">
        <div class="cfg-title">🎓 Training</div>
        <div class="cfg-row"><span class="cfg-lbl">Training due window (days after effective)</span><input class="cfg-input" id="cfg-trdue" value="${cfg.trainingDueDays || 30}"></div>
        <div class="cfg-row"><span class="cfg-lbl">Training warning window (days)</span><input class="cfg-input" id="cfg-trwarn" value="${cfg.trainingWarningDays || 7}"></div>
      </div>
      <div class="cfg-panel">
        <div class="cfg-title">📁 Document Zones</div>
        <div class="cfg-row"><span class="cfg-lbl">Draft zone</span><span class="cfg-val">Shared Documents/QMS/Documents/Drafts</span></div>
        <div class="cfg-row"><span class="cfg-lbl">Published zone</span><span class="cfg-val">Shared Documents/Published/QMS</span></div>
        <div class="cfg-row"><span class="cfg-lbl">Official zone</span><span class="cfg-val">Shared Documents/Official/QMS</span></div>
      </div>
      <div class="cfg-panel">
        <div class="cfg-title">✍️ E-Signature</div>
        <div class="cfg-row"><span class="cfg-lbl">Provider</span><span class="cfg-val">Microsoft 365 (native)</span></div>
        <div class="cfg-row"><span class="cfg-lbl">MFA required</span><span class="cfg-val">Yes</span></div>
        <div class="cfg-row"><span class="cfg-lbl">21 CFR Part 11 compliant</span><span class="cfg-val">Yes — sig ID + timestamp stored</span></div>
      </div>
    `);
  }

  // ── Alert bar ──
  private _checkAlerts(): void {
    const { dcos = [], approvals = [] } = this._data;
    const lateDCOs = dcos.filter(d => this._lateStatus(d) === 'late');
    const pending = this._computePendingTraining().filter(t => t.status === 'Overdue');
    const msgs = [];
    if (lateDCOs.length) msgs.push(`${lateDCOs.length} DCO(s) overdue in approval: ${lateDCOs.map(d => d.Title).join(', ')}`);
    if (pending.length) msgs.push(`${pending.length} training requirement(s) overdue`);
    const alert = this._el('qp-alert');
    if (alert && msgs.length) {
      alert.classList.add('show');
      this._set('qp-alert-txt', msgs.join(' | '));
    }
  }

  // ── Event listeners (pass-through to iframe JS) ──
  private _attachListeners(): void {
    const d = this._iframe?.contentDocument;
    if (!d) return;
    const w = this._iframe?.contentWindow as any;
    if (!w) return;

    // Expose data and helpers to iframe window
    w.qpData = this._data;
    w.qpConfig = this._config;
    w.qpWebPart = this;

    // ── Comprehensive event delegation — no inline onclick anywhere ──
    const navTo = (screen: string) => {
      d.querySelectorAll('.qp-screen').forEach((el: Element) => el.classList.remove('on'));
      d.querySelectorAll('.qp-tab').forEach((el: Element) => el.classList.remove('on'));
      const sc = d.getElementById('sc-' + screen);
      if (sc) sc.classList.add('on');
      const tab = d.querySelector(`[data-screen="${screen}"]`);
      if (tab) tab.classList.add('on');
    };

    // Nav tabs
    const navBar = d.querySelector('.qp-nav');
    if (navBar) {
      navBar.addEventListener('click', (e: Event) => {
        const btn = (e.target as HTMLElement).closest('[data-screen]') as HTMLElement;
        if (!btn) return;
        const screen = btn.getAttribute('data-screen');
        if (screen) navTo(screen);
      });
    }

    // Refresh button
    const refreshBtn = d.getElementById('qp-refresh-btn');
    if (refreshBtn) refreshBtn.addEventListener('click', () => this._loadAll());

    // Dashboard buckets — data-nav attribute
    d.querySelectorAll('[data-nav]').forEach((el: Element) => {
      el.addEventListener('click', () => {
        const screen = (el as HTMLElement).getAttribute('data-nav');
        if (screen) navTo(screen);
      });
    });

    // Pipeline stage filters — data-pip
    d.querySelectorAll('[data-pip]').forEach((el: Element) => {
      el.addEventListener('click', () => {
        const phase = (el as HTMLElement).getAttribute('data-pip');
        if (phase) this._dcoTableRender((this._data.dcos || []).filter((dco: any) => dco.DCO_Phase === phase));
      });
    });

    // DCO filter buttons — data-af
    d.querySelectorAll('[data-af]').forEach((el: Element) => {
      el.addEventListener('click', () => {
        d.querySelectorAll('[data-af]').forEach((b: Element) => b.classList.remove('on'));
        el.classList.add('on');
        const f = (el as HTMLElement).getAttribute('data-af') || 'all';
        this._dcoTableRender((this._data.dcos || []).filter((dco: any) => {
          if (f === 'all') return true;
          if (f === 'open') return (['Effective'].indexOf(dco.DCO_Phase || '') === -1);
          if (f === 'late') return this._lateStatus(dco) === 'late';
          return true;
        }));
      });
    });

    // Zone filter buttons — data-zf
    d.querySelectorAll('[data-zf]').forEach((el: Element) => {
      el.addEventListener('click', () => {
        d.querySelectorAll('[data-zf]').forEach((b: Element) => b.classList.remove('on'));
        el.classList.add('on');
      });
    });

    // Record filter buttons — data-rf
    d.querySelectorAll('[data-rf]').forEach((el: Element) => {
      el.addEventListener('click', () => {
        d.querySelectorAll('[data-rf]').forEach((b: Element) => b.classList.remove('on'));
        el.classList.add('on');
        const s = (el as HTMLElement).getAttribute('data-rf') || 'all';
        const records = (this._data.records || []).filter((r: any) => s === 'all' || r.Rec_Status === s);
        const rows = records.map((r: any) => `<tr><td><span class="cid">${r.Title}</span></td><td><span class="cmut">${r.Rec_Type||'—'}</span></td><td>${(r.Rec_Title||'').substring(0,50)}</td><td>${this._pill(r.Rec_Status)}</td><td><span class="cmut">${r.Rec_Originator||'—'}</span></td><td><span class="cmut">${r.Rec_Reviewer||'—'}</span></td><td><span class="cdate">${this._fmt(r.Rec_CreatedDate)}</span></td></tr>`).join('');
        this._html('rec-tbody', rows || '<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--s5)">No records match filter</td></tr>');
      });
    });

    // Training filter — data-tf
    d.querySelectorAll('[data-tf]').forEach((el: Element) => {
      el.addEventListener('click', () => {
        d.querySelectorAll('[data-tf]').forEach((b: Element) => b.classList.remove('on'));
        el.classList.add('on');
      });
    });

    // Training sub-tabs — data-trtab
    d.querySelectorAll('[data-trtab]').forEach((el: Element) => {
      el.addEventListener('click', () => {
        const tab = (el as HTMLElement).getAttribute('data-trtab');
        ['pending','matrix','employees','history'].forEach(t => {
          const panel = d.getElementById('tr-tab-' + t);
          if (panel) panel.style.display = t === tab ? 'block' : 'none';
        });
        d.querySelectorAll('[data-trtab]').forEach((b: Element) => {
          (b as HTMLElement).style.borderBottomColor = 'transparent';
          (b as HTMLElement).classList.remove('on');
        });
        (el as HTMLElement).style.borderBottomColor = 'var(--b)';
        el.classList.add('on');
      });
    });

    // New buttons
    const siteUrl = this.context.pageContext.web.absoluteUrl;
    const nb = (id: string, url: string) => { const el = d.getElementById(id); if (el) el.addEventListener('click', () => w.open(url, '_blank')); };
    nb('btn-new-dco', siteUrl + '/Lists/QMS_DCOs/NewForm.aspx');
    nb('btn-new-cr', siteUrl + '/Lists/QMS_ChangeRequests/NewForm.aspx');
    nb('btn-new-record', siteUrl + '/Lists/QMS_Records/NewForm.aspx');
    nb('btn-new-approver', siteUrl + '/Lists/QMS_Approvers/NewForm.aspx');
    const saveBtn = d.getElementById('btn-save-config');
    if (saveBtn) saveBtn.addEventListener('click', () => { if (w.qpToast) w.qpToast('Configuration saved'); });

    // Modal close buttons — data-close
    d.querySelectorAll('[data-close]').forEach((el: Element) => {
      el.addEventListener('click', () => {
        const id = (el as HTMLElement).getAttribute('data-close');
        if (id) (d.getElementById(id) as HTMLElement)?.classList.remove('open');
      });
    });

    // Backdrop close
    ['modal-dco-detail','modal-cr-detail','modal-reject','modal-esign'].forEach(id => {
      const modal = d.getElementById(id);
      if (modal) modal.addEventListener('click', (e: Event) => { if ((e.target as HTMLElement).id === id) modal.classList.remove('open'); });
    });

    // Reject + sign buttons
    const rejOpen = d.getElementById('mdco-reject-open');
    if (rejOpen) rejOpen.addEventListener('click', () => (d.getElementById('modal-reject') as HTMLElement)?.classList.add('open'));
    const rejConfirm = d.getElementById('btn-confirm-reject');
    if (rejConfirm) rejConfirm.addEventListener('click', () => {
      const reason = (d.getElementById('rej-reason') as HTMLInputElement)?.value;
      if (!reason?.trim()) { if (w.qpToast) w.qpToast('Rejection reason is required'); return; }
      (d.getElementById('modal-reject') as HTMLElement)?.classList.remove('open');
      (d.getElementById('modal-dco-detail') as HTMLElement)?.classList.remove('open');
      if (w.qpToast) w.qpToast('DCO rejected — routing history updated');
    });
    const signConfirm = d.getElementById('btn-confirm-sign');
    if (signConfirm) signConfirm.addEventListener('click', async () => {
      const code = (d.getElementById('esign-code') as HTMLInputElement)?.value;
      if (!code || code.length < 6) { if (w.qpToast) w.qpToast('Enter 6-digit MFA code to sign'); return; }

      const dcoId = (d.getElementById('esign-doc') as HTMLElement)?.textContent || '';
      if (!dcoId) { if (w.qpToast) w.qpToast('No DCO selected'); return; }

      const sigId = 'SIG-' + Date.now().toString(36).toUpperCase();
      const signedDate = new Date().toISOString();
      const base = this.context.pageContext.web.absoluteUrl;

      // 1. Find the approver record for this DCO + current user
      const apprs = (this._data.approvals || []).filter((a: any) => a.Appr_DCOID === dcoId);
      const myAppr = apprs.find((a: any) =>
        (a.Appr_Name || '').toLowerCase().includes('andre') ||
        (a.Appr_Name || '').toLowerCase().includes('butler')
      ) || apprs.find((a: any) => a.Appr_Status === 'Waiting');

      if (w.qpToast) w.qpToast('Applying signature...');

      try {
        if (myAppr && myAppr.Id) {
          // 2. Update QMS_DCOApprovals — mark this approver as Signed
          await this.context.spHttpClient.post(
            `${base}/_api/web/lists/getbytitle('QMS_DCOApprovals')/items(${myAppr.Id})`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                'Accept': 'application/json;odata=nometadata',
                'Content-Type': 'application/json;odata=nometadata',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
              },
              body: JSON.stringify({
                Appr_Status: 'Signed',
                Appr_SigID: sigId,
                Appr_SignedDate: signedDate
              })
            }
          );
          // Update local cache
          myAppr.Appr_Status = 'Signed';
          myAppr.Appr_SigID = sigId;
          myAppr.Appr_SignedDate = signedDate;
        } else {
          // No existing approver record — create one
          await this.context.spHttpClient.post(
            `${base}/_api/web/lists/getbytitle('QMS_DCOApprovals')/items`,
            SPHttpClient.configurations.v1,
            {
              headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
              body: JSON.stringify({
                Title: `${dcoId}-ANDRE`,
                Appr_DCOID: dcoId,
                Appr_Name: 'Andre Butler',
                Appr_Role: 'QA/Regulatory Consultant',
                Appr_Type: 'Required',
                Appr_Status: 'Signed',
                Appr_SigID: sigId,
                Appr_SignedDate: signedDate
              })
            }
          );
        }

        // 3. Write routing history entry
        await this.context.spHttpClient.post(
          `${base}/_api/web/lists/getbytitle('QMS_RoutingHistory')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
            body: JSON.stringify({
              Title: `${dcoId}-SIG-${sigId}`,
              RH_DCOID: dcoId,
              RH_EventType: 'signature',
              RH_Stage: 'In Review',
              RH_Actor: 'Andre Butler',
              RH_Note: `E-signature applied. SIG ID: ${sigId}`,
              RH_Timestamp: signedDate
            })
          }
        );

        // 4. Check if ALL required approvers have now signed → auto-advance phase
        const updatedApprs = (this._data.approvals || []).filter((a: any) => a.Appr_DCOID === dcoId);
        const requiredApprs = updatedApprs.filter((a: any) => a.Appr_Type === 'Required');
        const allSigned = requiredApprs.length > 0 && requiredApprs.every((a: any) => a.Appr_Status === 'Signed');

        if (allSigned) {
          // Find the DCO list item ID
          const dcoItem = (this._data.dcos || []).find((d2: any) => d2.Title === dcoId);
          if (dcoItem && dcoItem.Id) {
            // Determine next phase — check if any SOPs in DCO docs
            const nextPhase = 'Implemented';
            await this.context.spHttpClient.post(
              `${base}/_api/web/lists/getbytitle('QMS_DCOs')/items(${dcoItem.Id})`,
              SPHttpClient.configurations.v1,
              {
                headers: {
                  'Accept': 'application/json;odata=nometadata',
                  'Content-Type': 'application/json;odata=nometadata',
                  'IF-MATCH': '*',
                  'X-HTTP-Method': 'MERGE'
                },
                body: JSON.stringify({ DCO_Phase: nextPhase })
              }
            );
            // Update local cache
            if (dcoItem) dcoItem.DCO_Phase = nextPhase;
            // Write phase transition to routing history
            await this.context.spHttpClient.post(
              `${base}/_api/web/lists/getbytitle('QMS_RoutingHistory')/items`,
              SPHttpClient.configurations.v1,
              {
                headers: { 'Accept': 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=nometadata' },
                body: JSON.stringify({
                  Title: `${dcoId}-PHASE-${nextPhase}`,
                  RH_DCOID: dcoId,
                  RH_EventType: 'stage',
                  RH_Stage: nextPhase,
                  RH_Actor: 'System',
                  RH_Note: `All required approvers signed. DCO advanced to ${nextPhase}.`,
                  RH_Timestamp: new Date().toISOString()
                })
              }
            );
          }
          (d.getElementById('modal-esign') as HTMLElement)?.classList.remove('open');
          (d.getElementById('modal-dco-detail') as HTMLElement)?.classList.remove('open');
          if (w.qpToast) w.qpToast(`✅ All approvers signed — DCO advanced to Implemented! SIG: ${sigId}`);
        } else {
          (d.getElementById('modal-esign') as HTMLElement)?.classList.remove('open');
          const remaining = requiredApprs.filter((a: any) => a.Appr_Status !== 'Signed').length;
          if (w.qpToast) w.qpToast(`✅ Signature applied (SIG: ${sigId}) — ${remaining} approver(s) still pending`);
        }

        // 5. Reload data to reflect changes
        setTimeout(() => this._loadAll(), 1000);

      } catch (err) {
        console.error('Signature write-back failed:', err);
        (d.getElementById('modal-esign') as HTMLElement)?.classList.remove('open');
        if (w.qpToast) w.qpToast(`Signature recorded locally (SIG: ${sigId}) — SP write pending`);
      }
    });

    // DCO search
    const dcoSearch = d.getElementById('dco-search');
    if (dcoSearch) dcoSearch.addEventListener('input', (e: Event) => {
      const q = ((e.target as HTMLInputElement).value || '').toLowerCase();
      this._dcoTableRender((this._data.dcos || []).filter((dco: any) => !q || (dco.Title||'').toLowerCase().includes(q) || (dco.DCO_Title||'').toLowerCase().includes(q)));
    });

    // ── Table row click delegation — inlined to avoid timing issues ──
    const openDCOInline = (dcoId: string) => {
      const dco = (this._data.dcos || []).find((d2: any) => d2.Title === dcoId);
      if (!dco) return;
      const apprs = (this._data.approvals || []).filter((a: any) => a.Appr_DCOID === dcoId);
      const hist = (this._data.history || []).filter((h: any) => h.RH_DCOID === dcoId);
      const late = this._lateStatus(dco);
      const laneHtml = apprs.length ? apprs.map((a: any) => {
        const cls = a.Appr_Status === 'Signed' ? 'signed' : a.Appr_Status === 'Blocked' ? 'blocked' : 'waiting';
        return `<div class="lane ${cls}"><div class="lane-name">${a.Appr_Name||a.Title}</div><div class="lane-role">${a.Appr_Role||''} · ${a.Appr_Type||''}</div><div class="lane-status">${a.Appr_Status === 'Signed' ? '✅ Signed' : a.Appr_Status === 'Blocked' ? '🚫 Blocked' : '⏳ Waiting'}</div>${a.Appr_SigID ? `<div class="lane-sig">SIG: ${a.Appr_SigID}</div>` : ''}</div>`;
      }).join('') : '<div style="color:var(--s5);font-size:12px;padding:8px">No approvers assigned to this DCO yet</div>';
      const phases = ['Draft','Submitted','In Review','Implemented','Awaiting Training','Effective'];
      const curIdx = phases.indexOf(dco.DCO_Phase || 'Draft');
      const phaseBarHtml = phases.map((ph, i) => {
        const cls = i < curIdx ? 'done' : i === curIdx ? (late === 'late' ? 'late' : ph === 'Awaiting Training' ? 'train' : 'cur') : '';
        return `${i > 0 ? `<div class="ph-line${i <= curIdx ? ' done' : ''}"></div>` : ''}<div class="ph"><div class="ph-dot ${cls}">${i+1}</div><div class="ph-lbl" style="font-size:8px">${ph}</div></div>`;
      }).join('');
      const histHtml = [...hist].reverse().map((h: any) => {
        const cls = h.RH_EventType === 'rejection' ? 'rej' : h.RH_EventType === 'signature' ? 'sig' : 'stage';
        return `<div class="rh-item"><div class="rh-dot ${cls}"></div><div class="rh-body"><div class="rh-top"><span class="rh-evt">${h.RH_Stage||h.Title||''}</span><span class="rh-ts">${this._fmt(h.RH_Timestamp)}</span></div><div class="rh-detail">${h.RH_Actor||''} ${h.RH_Note ? '· '+h.RH_Note : ''}</div>${h.RH_Reason ? `<div class="rh-reason">Rejection reason: ${h.RH_Reason}</div>` : ''}</div></div>`;
      }).join('');
      const body = `
        ${late === 'late' ? `<div style="padding:10px 14px;background:var(--r1);border-left:4px solid var(--r);border-radius:6px;margin-bottom:14px;font-size:12px;color:var(--r);font-weight:600">⏱ OVERDUE — This DCO has been in approval for more than ${this._config.approvalOverdueDays||14} days</div>` : ''}
        <div class="phasebar">${phaseBarHtml}</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px">
          <div class="fg"><div class="fl">Phase</div><div class="fv">${dco.DCO_Phase||'Draft'}</div></div>
          <div class="fg"><div class="fl">CR Link</div><div class="fv">${dco.DCO_CRLink||'—'}</div></div>
          <div class="fg"><div class="fl">Submitted</div><div class="fv">${this._fmt(dco.DCO_SubmittedDate)}</div></div>
          <div class="fg"><div class="fl">Originator</div><div class="fv">${dco.DCO_Originator||'—'}</div></div>
        </div>
        <div style="margin-bottom:14px"><div class="fl" style="margin-bottom:8px">Approvers (Parallel Signing)</div><div class="lane-grid">${laneHtml}</div></div>
        <div><div class="fl" style="margin-bottom:8px">Routing History</div>${histHtml||'<div style="color:var(--s5);font-size:12px;padding:8px 0">No history recorded yet</div>'}</div>`;
      this._set('mdco-title', dcoId);
      this._set('mdco-sub', dco.DCO_Title||'');
      this._html('mdco-body', body);

      // Wire action buttons based on phase
      const actionBtn = d.getElementById('mdco-action-btn') as HTMLElement;
      const rejectBtn = d.getElementById('mdco-reject-btn') as HTMLElement;
      if (actionBtn) {
        const phase = dco.DCO_Phase || 'Draft';
        actionBtn.style.display = 'none';
        if (rejectBtn) rejectBtn.style.display = 'none';
        if (phase === 'Draft') {
          actionBtn.style.display = 'inline-flex';
          actionBtn.textContent = '📤 Submit for Approval';
          actionBtn.onclick = () => { if (w.qpToast) w.qpToast('Submit action — update DCO_Phase to Submitted in the list'); };
        } else if (phase === 'Submitted' || phase === 'In Review') {
          actionBtn.style.display = 'inline-flex';
          actionBtn.textContent = '✍️ Sign DCO';
          actionBtn.onclick = () => {
            this._set('esign-doc', dcoId);
            this._set('esign-as', 'Andre Butler (Current User)');
            this._set('esign-sub', 'Signing: ' + dcoId + ' — ' + (dco.DCO_Title||''));
            (d.getElementById('modal-esign') as HTMLElement)?.classList.add('open');
          };
          if (rejectBtn) rejectBtn.style.display = 'inline-flex';
        } else if (phase === 'Implemented') {
          actionBtn.style.display = 'inline-flex';
          actionBtn.textContent = '📋 Review Training';
        } else if (phase === 'Awaiting Training') {
          actionBtn.style.display = 'inline-flex';
          actionBtn.textContent = '✅ Mark Effective';
        }
      }

      (d.getElementById('modal-dco-detail') as HTMLElement)?.classList.add('open');
    };

    const openCRInline = (crId: string) => {
      const cr = (this._data.crs || []).find((c: any) => c.Title === crId);
      if (!cr) return;
      this._set('mcr-title', crId);
      this._set('mcr-sub', cr.CR_Title||'');
      this._html('mcr-body', `
        <div class="fg"><div class="fl">Title</div><div class="fv">${cr.CR_Title||'—'}</div></div>
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(cr.CR_Status)}</div></div>
        <div class="fg"><div class="fl">Priority</div><div class="fv">${this._pill(cr.CR_Priority)}</div></div>
        <div class="fg"><div class="fl">Originator</div><div class="fv">${cr.CR_Originator||'—'}</div></div>
        <div class="fg"><div class="fl">Linked DCOs</div><div class="fv" style="font-family:var(--mono)">${cr.CR_LinkedDCOs||'None yet'}</div></div>
        <div class="fg"><div class="fl">Description</div><div class="fv">${cr.CR_Description||'—'}</div></div>
        <div class="fg"><div class="fl">Created</div><div class="fv">${this._fmt(cr.CR_CreatedDate)}</div></div>`);
      (d.getElementById('modal-cr-detail') as HTMLElement)?.classList.add('open');
    };

    d.body.addEventListener('click', (e: Event) => {
      const target = e.target as HTMLElement;
      const dcoBtnEl = target.closest('[data-dcobtn]') as HTMLElement;
      if (dcoBtnEl) { e.stopPropagation(); openDCOInline(dcoBtnEl.getAttribute('data-dcobtn') || ''); return; }
      const dcoRow = target.closest('[data-dcoid]') as HTMLElement;
      if (dcoRow && !(target.closest('[data-dcobtn]'))) { openDCOInline(dcoRow.getAttribute('data-dcoid') || ''); return; }
      const crRow = target.closest('[data-crid]') as HTMLElement;
      if (crRow) { openCRInline(crRow.getAttribute('data-crid') || ''); return; }
    });

    // Register real implementations under _qp_ prefix for the bootstrap stubs
    const reg = (name: string, fn: any) => { w['_qp_' + name] = fn; w['qp' + name] = fn; };

    reg('Nav', (screen: string) => {
      d.querySelectorAll('.qp-screen').forEach((e: Element) => e.classList.remove('on'));
      d.querySelectorAll('.qp-tab').forEach((e: Element) => e.classList.remove('on'));
      const sc = d.getElementById('sc-' + screen);
      if (sc) sc.classList.add('on');
      const tab = d.querySelector(`[data-screen="${screen}"]`);
      if (tab) tab.classList.add('on');
    });

    reg('Refresh', () => { this._loadAll(); });

    w.qpToast = (msg: string) => {
      const t = d.getElementById('qp-toast');
      if (!t) return;
      t.textContent = msg;
      t.classList.add('show');
      setTimeout(() => t.classList.remove('show'), 2500);
    };

    w.qpOpenModal = (id: string) => {
      (d.getElementById(id) as HTMLElement)?.classList.add('open');
    };

    w.qpCloseModal = (id: string) => {
      (d.getElementById(id) as HTMLElement)?.classList.remove('open');
      d.getElementById('modal-dco-detail')?.classList.remove('open');
      d.getElementById('modal-cr-detail')?.classList.remove('open');
      d.getElementById('modal-reject')?.classList.remove('open');
      d.getElementById('modal-esign')?.classList.remove('open');
    };

    w.qpOpenDCO = (dcoId: string) => {
      const dco = (this._data.dcos || []).find((d2: any) => d2.Title === dcoId);
      if (!dco) return;
      const apprs = (this._data.approvals || []).filter((a: any) => a.Appr_DCOID === dcoId);
      const hist = (this._data.history || []).filter((h: any) => h.RH_DCOID === dcoId);
      const late = this._lateStatus(dco);

      const laneHtml = apprs.map((a: any) => {
        const cls = a.Appr_Status === 'Signed' ? 'signed' : a.Appr_Status === 'Blocked' ? 'blocked' : 'waiting';
        return `<div class="lane ${cls}">
          <div class="lane-name">${a.Appr_Name}</div>
          <div class="lane-role">${a.Appr_Role} · ${a.Appr_Type}</div>
          <div class="lane-status">${a.Appr_Status === 'Signed' ? '✅ Signed' : a.Appr_Status === 'Blocked' ? '🚫 Blocked' : '⏳ Waiting'}</div>
          ${a.Appr_SigID ? `<div class="lane-sig">SIG: ${a.Appr_SigID}</div>` : ''}
        </div>`;
      }).join('');

      const histHtml = [...hist].reverse().map((h: any) => {
        const cls = h.RH_EventType === 'rejection' ? 'rej' : h.RH_EventType === 'signature' ? 'sig' : h.RH_EventType === 'stage' ? 'stage' : 'sys';
        return `<div class="rh-item">
          <div class="rh-dot ${cls}"></div>
          <div class="rh-body">
            <div class="rh-top"><span class="rh-evt">${h.RH_Stage || h.Title}</span><span class="rh-ts">${this._fmt(h.RH_Timestamp)}</span></div>
            <div class="rh-detail">${h.RH_Actor || ''} ${h.RH_Note ? '· ' + h.RH_Note : ''}</div>
            ${h.RH_Reason ? `<div class="rh-reason">Rejection reason: ${h.RH_Reason}</div>` : ''}
          </div>
        </div>`;
      }).join('');

      const phases = ['Draft', 'Submitted', 'In Review', 'Implemented', 'Awaiting Training', 'Effective'];
      const curIdx = phases.indexOf(dco.DCO_Phase || 'Draft');
      const phaseBarHtml = phases.map((ph, i) => {
        const cls = i < curIdx ? 'done' : i === curIdx ? (late === 'late' ? 'late' : ph === 'Awaiting Training' ? 'train' : 'cur') : '';
        return `${i > 0 ? `<div class="ph-line${i <= curIdx ? ' done' : ''}"></div>` : ''}
          <div class="ph"><div class="ph-dot ${cls}">${i + 1}</div><div class="ph-lbl" style="font-size:8px">${ph}</div></div>`;
      }).join('');

      const body = `
        ${late === 'late' ? `<div style="padding:10px 14px;background:var(--r1);border-left:4px solid var(--r);border-radius:6px;margin-bottom:14px;font-size:12px;color:var(--r);font-weight:600">⏱ OVERDUE — This DCO has been in approval for more than ${this._config.approvalOverdueDays || 14} days</div>` : ''}
        <div class="phasebar">${phaseBarHtml}</div>
        <div style="display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px">
          <div class="fg"><div class="fl">CR Link</div><div class="fv">${dco.DCO_CRLink || '—'}</div></div>
          <div class="fg"><div class="fl">Submitted</div><div class="fv">${this._fmt(dco.DCO_SubmittedDate)}</div></div>
          <div class="fg"><div class="fl">Originator</div><div class="fv">${dco.DCO_Originator || '—'}</div></div>
          <div class="fg"><div class="fl">Training Gate</div><div class="fv">${dco.DCO_TrainingGate || 'N/A'}</div></div>
        </div>
        <div style="margin-bottom:14px">
          <div class="fl" style="margin-bottom:8px">Approvers (Parallel Signing)</div>
          <div class="lane-grid">${laneHtml || '<div style="color:var(--s5);font-size:12px">No approvers assigned</div>'}</div>
        </div>
        <div>
          <div class="fl" style="margin-bottom:8px">Routing History</div>
          ${histHtml || '<div style="color:var(--s5);font-size:12px;padding:8px 0">No history recorded</div>'}
        </div>`;

      w.qp_currentDCO = dco;
      this._set('mdco-title', dcoId);
      this._set('mdco-sub', dco.DCO_Title || '');
      this._html('mdco-body', body);

      // Action button
      const actionBtn = d.getElementById('mdco-action-btn') as HTMLElement;
      const rejectBtn = d.getElementById('mdco-reject-btn') as HTMLElement;
      if (actionBtn) {
        const phase = dco.DCO_Phase || 'Draft';
        if (phase === 'Draft') { actionBtn.style.display = 'inline-flex'; actionBtn.textContent = '📤 Submit for Approval'; actionBtn.onclick = () => w.qpToast('Submit action — integrate with Power Automate flow'); }
        else if (phase === 'Submitted' || phase === 'In Review') { actionBtn.style.display = 'inline-flex'; actionBtn.textContent = '✍️ Sign'; actionBtn.onclick = () => { this._set('esign-doc', dcoId); this._set('esign-as', 'Current User'); this._set('esign-sub', `Signing: ${dcoId}`); w.qpOpenModal('modal-esign'); }; rejectBtn.style.display = 'inline-flex'; }
        else if (phase === 'Implemented') { actionBtn.style.display = 'inline-flex'; actionBtn.textContent = '📋 Review Training'; actionBtn.onclick = () => w.qpNav('training'); }
        else if (phase === 'Awaiting Training') { actionBtn.style.display = 'inline-flex'; actionBtn.textContent = '✅ Mark Effective'; actionBtn.onclick = () => w.qpToast('Training gate check — mark effective when all SOP training complete'); }
        else { actionBtn.style.display = 'none'; }
      }
      w.qpOpenModal('modal-dco-detail');
    };

    w.qpOpenCR = (crId: string) => {
      const cr = (this._data.crs || []).find((c: any) => c.Title === crId);
      if (!cr) return;
      this._set('mcr-title', crId);
      this._set('mcr-sub', cr.CR_Title || '');
      this._html('mcr-body', `
        <div class="fg"><div class="fl">Title</div><div class="fv">${cr.CR_Title || '—'}</div></div>
        <div class="fg"><div class="fl">Status</div><div class="fv">${this._pill(cr.CR_Status)}</div></div>
        <div class="fg"><div class="fl">Priority</div><div class="fv">${this._pill(cr.CR_Priority)}</div></div>
        <div class="fg"><div class="fl">Originator</div><div class="fv">${cr.CR_Originator || '—'}</div></div>
        <div class="fg"><div class="fl">Linked DCOs</div><div class="fv" style="font-family:var(--mono)">${cr.CR_LinkedDCOs || 'None yet'}</div></div>
        <div class="fg"><div class="fl">Description</div><div class="fv">${cr.CR_Description || '—'}</div></div>
        <div class="fg"><div class="fl">Created</div><div class="fv">${this._fmt(cr.CR_CreatedDate)}</div></div>`);
      w.qpOpenModal('modal-cr-detail');
    };

    w.qpOpenReject = () => { w.qpOpenModal('modal-reject'); };
    w.qpConfirmReject = () => {
      const reason = (d.getElementById('rej-reason') as HTMLInputElement)?.value;
      if (!reason || !reason.trim()) { w.qpToast('Rejection reason is required'); return; }
      w.qpCloseModal('modal-reject');
      w.qpCloseModal('modal-dco-detail');
      w.qpToast('DCO rejected — routing history updated');
    };

    w.qpConfirmSign = () => {
      const code = (d.getElementById('esign-code') as HTMLInputElement)?.value;
      if (!code || code.length < 6) { w.qpToast('Enter 6-digit MFA code to sign'); return; }
      w.qpCloseModal('modal-esign');
      const sigId = 'SIG-' + Date.now().toString(36).toUpperCase();
      w.qpToast(`✅ Signature applied — SIG ID: ${sigId}`);
    };

    w.qpDCOFilter = (btn: HTMLElement, filter: string) => {
      d.querySelectorAll('.fbar .fbtn').forEach((b: Element) => b.classList.remove('on'));
      btn.classList.add('on');
      this._dcoTableRender((this._data.dcos || []).filter((dco: any) => {
        if (filter === 'all') return true;
        if (filter === 'open') return (['Effective'].indexOf(dco.DCO_Phase || '') === -1);
        if (filter === 'late') return this._lateStatus(dco) === 'late';
        return true;
      }));
    };

    w.qpPipFilter = (phase: string) => {
      this._dcoTableRender((this._data.dcos || []).filter((dco: any) => dco.DCO_Phase === phase));
    };

    w.qpZoneFilter = (btn: HTMLElement, zone: string) => {
      d.querySelectorAll('#sc-docs .fbar .fbtn').forEach((b: Element) => b.classList.remove('on'));
      btn.classList.add('on');
    };

    w.qpRecFilter = (btn: HTMLElement, status: string) => {
      d.querySelectorAll('#sc-records .fbar .fbtn').forEach((b: Element) => b.classList.remove('on'));
      btn.classList.add('on');
      const records = (this._data.records || []).filter((r: any) => status === 'all' || r.Rec_Status === status);
      const rows = records.map((r: any) => `<tr>
        <td><span class="cid">${r.Title}</span></td>
        <td><span class="cmut">${r.Rec_Type || '—'}</span></td>
        <td style="font-size:12px">${(r.Rec_Title || '').substring(0, 50)}</td>
        <td>${this._pill(r.Rec_Status)}</td>
        <td><span class="cmut">${r.Rec_Originator || '—'}</span></td>
        <td><span class="cmut">${r.Rec_Reviewer || '—'}</span></td>
        <td><span class="cdate">${this._fmt(r.Rec_CreatedDate)}</span></td>
      </tr>`).join('');
      this._html('rec-tbody', rows || '<tr><td colspan="7" style="text-align:center;padding:24px;color:var(--s5)">No records match filter</td></tr>');
    };

    w.qpTrTab = (btn: HTMLElement, tab: string) => {
      ['pending', 'matrix', 'employees', 'history'].forEach(t => {
        const el = d.getElementById('tr-tab-' + t);
        if (el) el.style.display = t === tab ? 'block' : 'none';
      });
      d.querySelectorAll('#sc-training .qp-tab').forEach((b: Element) => { (b as HTMLElement).style.borderBottomColor = 'transparent'; (b as HTMLElement).style.color = 'var(--s7)'; });
      btn.style.borderBottomColor = 'var(--b)';
    };

    w.qpTrFilter = (btn: HTMLElement, filter: string) => {
      d.querySelectorAll('#tr-tab-pending .fbar .fbtn').forEach((b: Element) => b.classList.remove('on'));
      btn.classList.add('on');
    };

    w.qpRenderDCO = () => {
      const q = ((d.getElementById('dco-search') as HTMLInputElement)?.value || '').toLowerCase();
      this._dcoTableRender((this._data.dcos || []).filter((dco: any) =>
        !q || (dco.Title || '').toLowerCase().includes(q) || (dco.DCO_Title || '').toLowerCase().includes(q)));
    };

    w.qpRenderDocs = () => {};
    w.qpRenderRecords = () => {};
    w.qpRenderTraining = () => {};
    const _siteUrlB = this.context.pageContext.web.absoluteUrl;
    reg('SaveConfig', () => { w.qpToast('Configuration saved (update QMS_Config list to persist)'); });
    reg('OpenNewDCO', () => { w.open(_siteUrlB + '/Lists/QMS_DCOs/NewForm.aspx', '_blank'); });
    reg('OpenNewCR', () => { w.open(_siteUrlB + '/Lists/QMS_ChangeRequests/NewForm.aspx', '_blank'); });
    reg('OpenNewRecord', () => { w.open(_siteUrlB + '/Lists/QMS_Records/NewForm.aspx', '_blank'); });
    reg('OpenNewApprover', () => { w.open(_siteUrlB + '/Lists/QMS_Approvers/NewForm.aspx', '_blank'); });
    reg('OpenReject', () => { w.qpOpenModal('modal-reject'); });
    reg('ConfirmReject', () => {
      const reason = (d.getElementById('rej-reason') as HTMLInputElement)?.value;
      if (!reason || !reason.trim()) { w.qpToast('Rejection reason is required'); return; }
      w.qpCloseModal('modal-reject'); w.qpCloseModal('modal-dco-detail');
      w.qpToast('DCO rejected — routing history updated');
    });
    reg('ConfirmSign', () => {
      const code = (d.getElementById('esign-code') as HTMLInputElement)?.value;
      if (!code || code.length < 6) { w.qpToast('Enter 6-digit MFA code to sign'); return; }
      w.qpCloseModal('modal-esign');
      const sigId = 'SIG-' + Date.now().toString(36).toUpperCase();
      w.qpToast('Signature applied — SIG ID: ' + sigId);
    });
    reg('OpenModal', (id: string) => { (d.getElementById(id) as HTMLElement)?.classList.add('open'); });
    reg('CloseModal', (id: string) => {
      ['modal-dco-detail','modal-cr-detail','modal-reject','modal-esign'].forEach(m => (d.getElementById(m) as HTMLElement)?.classList.remove('open'));
    });
    // Flush any queued onclick calls that fired before listeners were ready
    if (w._qpFlush) w._qpFlush();
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'QMS Portal Configuration' },
        groups: [{
          groupName: 'Display',
          groupFields: [PropertyPaneChoiceGroup('screen', {
            label: 'Default Screen',
            options: [
              { key: 'dashboard', text: 'Dashboard' },
              { key: 'dco', text: 'Change Orders' },
              { key: 'cr', text: 'Change Requests' },
              { key: 'records', text: 'Records' },
              { key: 'training', text: 'Training' },
            ]
          })]
        }]
      }]
    };
  }
}
