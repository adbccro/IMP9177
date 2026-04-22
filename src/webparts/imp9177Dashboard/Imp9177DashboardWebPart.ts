/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable dot-notation */
/* eslint-disable no-void */
/* eslint-disable @typescript-eslint/no-floating-promises */
import { Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneChoiceGroup } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IIMP9177DashboardWebPartProps { dashboard: string; }

// ---------------------------------------------------------------------------
// PURE HTML+CSS SHELLS -- zero <script> tags, zero inline JS
// All render logic lives in TypeScript below (CSP-safe)
// ---------------------------------------------------------------------------

const CSS = `
:root{--n:#0C2D5E;--b:#1E56A0;--b5:#3B82D4;--b1:#DBEAFE;--b0:#EFF6FF;--s0:#F8FAFC;--s1:#F1F5F9;--s2:#E2E8F0;--s5:#64748B;--s7:#334155;--r:#DC2626;--a:#D97706;--g:#059669;--w:#FFFFFF;}
*{box-sizing:border-box;margin:0;padding:0;}
body{background:var(--s0);color:var(--s7);font-family:'Source Sans 3',system-ui,sans-serif;font-size:14px;}
header{background:linear-gradient(135deg,var(--n),var(--b));padding:0 28px;display:flex;align-items:center;justify-content:space-between;height:60px;box-shadow:0 2px 8px rgba(12,45,94,.25);}
.hl{display:flex;align-items:center;gap:12px;}
.hbadge{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.25);color:#fff;font-weight:700;font-size:11px;letter-spacing:1px;padding:4px 10px;border-radius:5px;}
.ht{color:#fff;font-family:'Libre Baskerville',Georgia,serif;font-size:16px;font-weight:700;}
.hs{color:rgba(255,255,255,.6);font-size:11px;margin-top:1px;}
.hr{display:flex;align-items:center;gap:8px;}
.hm{color:rgba(255,255,255,.55);font-size:11px;}
.brf{background:rgba(255,255,255,.15);border:1px solid rgba(255,255,255,.3);color:#fff;padding:6px 13px;border-radius:5px;font-size:11px;font-weight:600;cursor:pointer;}
.brf:hover{background:rgba(255,255,255,.25);}
.tbn{padding:6px 13px;border-radius:5px;border:1px solid rgba(255,255,255,.25);background:transparent;color:rgba(255,255,255,.7);font-size:11px;cursor:pointer;}
.tbn:hover,.tbn.on{background:rgba(255,255,255,.2);color:#fff;border-color:rgba(255,255,255,.5);}
main{padding:20px 28px;max-width:1600px;margin:0 auto;}
.krow{display:grid;gap:12px;margin-bottom:20px;}
.krow.k6{grid-template-columns:repeat(6,1fr);}
.krow.k5{grid-template-columns:repeat(5,1fr);}
.krow.k7{grid-template-columns:repeat(7,1fr);}
.kpi{background:var(--w);border:1px solid var(--s2);border-radius:8px;padding:16px 18px;box-shadow:0 1px 3px rgba(0,0,0,.07);border-top:3px solid var(--b5);}
.kpi.r{border-top-color:var(--r);}.kpi.a{border-top-color:var(--a);}.kpi.g{border-top-color:var(--g);}
.kl{font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:.7px;color:var(--s5);margin-bottom:6px;}
.kv{font-size:30px;font-weight:700;color:var(--n);line-height:1;font-family:'Libre Baskerville',Georgia,serif;}
.kv.r{color:var(--r);}.kv.a{color:var(--a);}.kv.g{color:var(--g);}
.ks{font-size:11px;color:var(--s5);margin-top:3px;}
.g21{display:grid;grid-template-columns:2fr 1fr;gap:16px;margin-bottom:16px;}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:16px;margin-bottom:16px;}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:16px;margin-bottom:16px;}
.full{margin-bottom:16px;}
.panel{background:var(--w);border:1px solid var(--s2);border-radius:8px;overflow:hidden;box-shadow:0 1px 3px rgba(0,0,0,.07);}
.ph{padding:12px 18px;border-bottom:1px solid var(--s1);display:flex;align-items:center;justify-content:space-between;background:linear-gradient(to right,var(--b0),var(--w));}
.pt{font-size:13px;font-weight:700;color:var(--n);font-family:'Libre Baskerville',Georgia,serif;display:flex;align-items:center;gap:7px;}
.bx{font-size:10px;font-weight:600;padding:2px 7px;border-radius:12px;background:var(--b1);color:#1E56A0;}
.bx.r{background:#FEE2E2;color:#B91C1C;}.bx.a{background:#FEF3C7;color:#92400E;}.bx.g{background:#D1FAE5;color:#065F46;}
table{width:100%;border-collapse:collapse;font-size:12px;}
th{padding:8px 12px;text-align:left;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--s5);border-bottom:2px solid var(--s2);background:var(--s0);}
td{padding:8px 12px;border-bottom:1px solid var(--s1);color:var(--s7);vertical-align:middle;}
tr:last-child td{border-bottom:none;}tr:hover td{background:var(--b0);}
.pill{display:inline-flex;padding:2px 9px;border-radius:12px;font-size:10px;font-weight:600;white-space:nowrap;}
.pr{background:#FEE2E2;color:#B91C1C;}.pa{background:#FEF3C7;color:#92400E;}.pg{background:#D1FAE5;color:#065F46;}.pb{background:var(--b1);color:#1E56A0;}.py{background:var(--s1);color:var(--s5);}
.did{font-family:monospace;font-size:11px;color:#1E56A0;font-weight:600;}
.dlink{font-family:monospace;font-size:11px;color:#1E56A0;font-weight:600;text-decoration:none;}
.dlink:hover{text-decoration:underline;}
.gh{padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;display:flex;align-items:center;justify-content:space-between;border-bottom:1px solid var(--s1);}
.gh.red{background:#FFF5F5;color:#B91C1C;}.gh.amb{background:#FFFBEB;color:#92400E;}.gh.grn{background:#F0FDF4;color:#065F46;}.gh.blu{background:var(--b0);color:#1E56A0;}.gh.gry{background:var(--s0);color:var(--s5);}
.donwrap{display:flex;align-items:center;gap:20px;padding:16px 18px;}
.dleg{display:flex;flex-direction:column;gap:9px;flex:1;}
.dli{display:flex;align-items:center;justify-content:space-between;gap:8px;}
.dld{width:9px;height:9px;border-radius:50%;flex-shrink:0;}
.dll{font-size:12px;color:var(--s7);flex:1;}
.dlv{font-size:13px;font-weight:700;color:var(--n);}
.dw{padding:16px 18px;display:flex;flex-direction:column;gap:10px;}
.dr{display:flex;align-items:center;gap:8px;}
.dlb{font-size:11px;color:var(--s5);width:120px;text-align:right;flex-shrink:0;}
.dt{flex:1;background:var(--s1);border-radius:3px;height:16px;overflow:hidden;}
.df{height:100%;border-radius:3px;display:flex;align-items:center;padding:0 7px;}
.dn{font-size:10px;font-weight:700;color:#fff;}
.dnum{font-size:11px;color:var(--s5);width:28px;flex-shrink:0;}
.sup2{display:grid;grid-template-columns:1fr 1fr;border-bottom:1px solid var(--s1);}
.sbox{padding:14px 18px;text-align:center;}
.sbox:first-child{border-right:1px solid var(--s1);}
.sbig{font-size:28px;font-weight:700;font-family:'Libre Baskerville',Georgia,serif;}
.ssub{font-size:10px;text-transform:uppercase;letter-spacing:.6px;color:var(--s5);margin-top:3px;}
.scats{padding:12px 18px;}
.scat{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;}
.scatp{height:4px;background:var(--s1);border-radius:2px;margin-bottom:2px;overflow:hidden;}
.scatb{height:100%;background:var(--b5);border-radius:2px;}
.hgrid{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;padding:16px 18px;}
.hcell{border-radius:7px;padding:14px;text-align:center;}
.hn{font-size:28px;font-weight:700;line-height:1;font-family:'Libre Baskerville',Georgia,serif;}
.hl2{font-size:9px;text-transform:uppercase;letter-spacing:.7px;margin-top:5px;font-weight:600;}
.tl{padding:14px 18px;display:flex;flex-direction:column;gap:0;}
.ti{display:flex;gap:12px;padding-bottom:14px;}
.tline{display:flex;flex-direction:column;align-items:center;}
.tdot{width:9px;height:9px;border-radius:50%;flex-shrink:0;margin-top:4px;}
.tcon{width:2px;flex:1;background:var(--s2);margin:3px 0;}
.tmt{font-size:10px;color:var(--s5);margin-bottom:2px;font-weight:600;text-transform:uppercase;letter-spacing:.4px;}
.ttx{font-size:13px;font-weight:700;color:var(--n);}
.tdt{font-size:11px;color:var(--s5);margin-top:3px;line-height:1.45;}
.tp{display:none;}.tp.on{display:block;}
.pbar{height:8px;background:var(--s1);border-radius:4px;overflow:hidden;margin-top:4px;}
.pbf{height:100%;border-radius:4px;background:var(--b5);}
.bcard{padding:16px 18px;display:flex;flex-direction:column;gap:0;}
.brow{display:flex;justify-content:space-between;align-items:baseline;margin-bottom:4px;}
.bl{font-size:12px;color:var(--s7);}
.bv{font-size:22px;font-weight:700;color:var(--n);font-family:'Libre Baskerville',Georgia,serif;}
.btrack{height:8px;background:var(--s1);border-radius:4px;overflow:hidden;margin-bottom:12px;}
.bfill{height:100%;background:var(--b5);border-radius:4px;}
.bpso{margin-top:12px;border-top:1px solid var(--s1);padding-top:12px;}
.bptag{display:inline-block;padding:3px 9px;border-radius:4px;font-size:10px;font-weight:600;background:#FEF3C7;color:#92400E;margin-bottom:4px;}
.hchip{background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.2);color:rgba(255,255,255,.85);padding:4px 12px;border-radius:14px;font-size:11px;font-weight:600;}
.hchip.warn{background:rgba(217,119,6,.3);border-color:rgba(217,119,6,.5);color:#FCD34D;}
footer{margin-top:8px;padding:14px 28px;border-top:1px solid var(--s2);font-size:11px;color:var(--s5);text-align:center;background:var(--w);}
@media(max-width:1200px){.krow.k6,.krow.k7{grid-template-columns:repeat(3,1fr);}.g21,.g3{grid-template-columns:1fr;}}
@media(max-width:768px){.krow.k6,.krow.k5,.krow.k7{grid-template-columns:repeat(2,1fr);}.g2{grid-template-columns:1fr;}}
`;

const GFONTS = '<link href="https://fonts.googleapis.com/css2?family=Libre+Baskerville:wght@400;700&family=Source+Sans+3:wght@300;400;500;600;700&display=swap" rel="stylesheet">';

function makeHtml(badge: string, title: string, subtitle: string, bodyHtml: string, footerText: string): string {
  return `<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">${GFONTS}<style>${CSS}</style></head><body>
<header>
  <div class="hl">
    <div class="hbadge">${badge}</div>
    <div><div class="ht">${title}</div><div class="hs">${subtitle}</div></div>
  </div>
  <div class="hr"><span class="hm" id="rt"></span><button class="brf" id="refreshBtn">&#8635; Refresh</button></div>
</header>
<main id="mc">${bodyHtml}</main>
<footer>${footerText}</footer>
</body></html>`;
}

const MRS_BODY = `
<div class="krow k6">
  <div class="kpi"><div class="kl">Total Documents</div><div class="kv" id="kD">0</div><div class="ks">W1 + W2 + Existing</div></div>
  <div class="kpi r"><div class="kl">Sign-Off Past Due</div><div class="kv r" id="kP">0</div><div class="ks">Requires immediate action</div></div>
  <div class="kpi a"><div class="kl">Pending Feedback</div><div class="kv a" id="kF">0</div><div class="ks">Awaiting 3H review</div></div>
  <div class="kpi g"><div class="kl">ADB Drafting Done</div><div class="kv g" id="kA">0</div><div class="ks">Docs ready</div></div>
  <div class="kpi"><div class="kl">Total Gaps</div><div class="kv" id="kG">77</div><div class="ks">67 original + 10 GAP-REC</div></div>
  <div class="kpi a"><div class="kl">Open CAPAs</div><div class="kv a" id="kC">0</div><div class="ks">Active corrective actions</div></div>
</div>
<div class="g21">
  <div class="panel">
    <div class="ph"><div class="pt">&#128196; Document Registry <span class="bx" id="dbx">0 docs</span></div><span style="font-size:11px;color:var(--s5)">Click doc # to open in SharePoint</span></div>
    <div id="dg"><div style="padding:16px 18px;font-size:12px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="panel">
    <div class="ph"><div class="pt">&#128269; Gap Registry <span class="bx">77 Gaps</span></div></div>
    <div class="donwrap">
      <svg width="106" height="106" viewBox="0 0 106 106">
        <circle cx="53" cy="53" r="38" fill="none" stroke="#E2E8F0" stroke-width="15"/>
        <g id="ga"></g>
        <text x="53" y="48" text-anchor="middle" font-family="Libre Baskerville,serif" font-size="19" font-weight="700" fill="#0C2D5E" id="gt">77</text>
        <text x="53" y="63" text-anchor="middle" font-family="Source Sans 3,sans-serif" font-size="9" fill="#64748B">GAPS</text>
      </svg>
      <div class="dleg">
        <div class="dli"><div class="dld" style="background:#DC2626"></div><span class="dll">Critical</span><span class="dlv" id="gc">30</span></div>
        <div class="dli"><div class="dld" style="background:#D97706"></div><span class="dll">Major</span><span class="dlv" id="gm">31</span></div>
        <div class="dli"><div class="dld" style="background:#059669"></div><span class="dll">Minor</span><span class="dlv" id="gn">6</span></div>
        <div class="dli"><div class="dld" style="background:#3B82D4"></div><span class="dll">GAP-REC (3H)</span><span class="dlv" id="gr">10</span></div>
      </div>
    </div>
    <div style="border-top:1px solid var(--s1);display:grid;grid-template-columns:1fr 1fr;">
      <div style="padding:12px 18px;border-right:1px solid var(--s1);text-align:center;">
        <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--s5);margin-bottom:4px">Open</div>
        <div style="font-size:24px;font-weight:700;color:var(--r);font-family:'Libre Baskerville',serif" id="go">67</div>
      </div>
      <div style="padding:12px 18px;text-align:center;">
        <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--s5);margin-bottom:4px">Closed</div>
        <div style="font-size:24px;font-weight:700;color:var(--g);font-family:'Libre Baskerville',serif" id="gcl">0</div>
      </div>
    </div>
  </div>
</div>
<div class="g3">
  <div class="panel">
    <div class="ph"><div class="pt">&#127981; Supplier Registry <span class="bx" id="sbx">18</span></div></div>
    <div id="ss"><div style="padding:16px 18px;font-size:12px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="panel">
    <div class="ph"><div class="pt">&#129514; Ingredient / CoA Status <span class="bx" id="ibx">101</span></div></div>
    <div class="dw" id="ib"><div style="font-size:12px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="panel">
    <div class="ph"><div class="pt">&#9888; CAPA Log <span class="bx" id="cbx">0 records</span></div></div>
    <div id="cc"><div style="padding:16px 18px;font-size:12px;color:var(--s5)">Loading...</div></div>
  </div>
</div>`;

const RAID_BODY = `
<div class="krow k5">
  <div class="kpi"><div class="kl">Total Actions</div><div class="kv" id="kAt">0</div><div class="ks" id="kAo">0 open</div></div>
  <div class="kpi r"><div class="kl">Critical / Blocking</div><div class="kv r" id="kAc">0</div><div class="ks">Immediate attention</div></div>
  <div class="kpi a"><div class="kl">Open Issues</div><div class="kv a" id="kIo">0</div><div class="ks">ISS log</div></div>
  <div class="kpi g"><div class="kl">Decisions Logged</div><div class="kv g" id="kDl">0</div><div class="ks">DEC series</div></div>
  <div class="kpi"><div class="kl">Meetings Held</div><div class="kv" id="kMh">0</div><div class="ks">Next: week of Apr 21</div></div>
</div>
<div id="tp-ov" class="tp on">
  <div class="panel full">
    <div class="ph"><div class="pt">&#128202; Project Health Summary</div></div>
    <div class="hgrid">
      <div class="hcell" style="background:#FEE2E2"><div class="hn" id="hc" style="color:#B91C1C">0</div><div class="hl2" style="color:#B91C1C">Critical Actions</div></div>
      <div class="hcell" style="background:#FEF3C7"><div class="hn" id="hp" style="color:#92400E">0</div><div class="hl2" style="color:#92400E">Past Due</div></div>
      <div class="hcell" style="background:#FEE2E2"><div class="hn" id="hi" style="color:#B91C1C">0</div><div class="hl2" style="color:#B91C1C">Open Issues</div></div>
      <div class="hcell" style="background:#D1FAE5"><div class="hn" id="hd" style="color:#065F46">0</div><div class="hl2" style="color:#065F46">Closed Actions</div></div>
      <div class="hcell" style="background:var(--b1)"><div class="hn" style="color:#1E56A0">77</div><div class="hl2" style="color:#1E56A0">Total Gaps</div></div>
    </div>
  </div>
  <div class="g2">
    <div class="panel"><div class="ph"><div class="pt">&#128308; Critical / Blocking <span class="bx r" id="cbx">0 items</span></div></div><div id="ct"><div style="padding:16px;text-align:center;color:var(--s5)">Loading...</div></div></div>
    <div class="panel"><div class="ph"><div class="pt">&#128197; Meeting Timeline</div></div>
      <div class="tl">
        <div class="ti"><div class="tline"><div class="tdot" style="background:var(--g)"></div><div class="tcon"></div></div><div><div class="tmt">MM-005 &middot; Apr 8, 2026 &middot; 35 min</div><div class="ttx">W1/W2 Review + Supplier Controls</div><div class="tdt">Amazon excluded (DEC-035). Tina Qin interim QA Approver (DEC-040). 9 new actions (AC-029&ndash;037).</div></div></div>
        <div class="ti"><div class="tline"><div class="tdot" style="background:var(--g)"></div><div class="tcon"></div></div><div><div class="tmt">MM-004 &middot; Apr 6, 2026 &middot; 29 min</div><div class="ttx">W1 Deliverables Orientation</div><div class="tdt">SharePoint access resent. CR/DCO workflow explained. Training matrix (AC-028).</div></div></div>
        <div class="ti"><div class="tline"><div class="tdot" style="background:var(--a)"></div></div><div><div class="tmt">NEXT &middot; Week of Apr 21, 2026</div><div class="ttx">Tina travel Apr 13&ndash;17 &mdash; no meetings</div><div class="tdt">W2 feedback review. QD Yang invite. W3 kickoff.</div></div></div>
      </div>
    </div>
  </div>
</div>
<div id="tp-ac" class="tp">
  <div class="panel full"><div class="ph"><div class="pt">&#9889; Action Log <span class="bx" id="abx">0 actions</span></div></div><div id="ag"><div style="padding:16px;text-align:center;color:var(--s5)">Loading...</div></div></div>
</div>
<div id="tp-is" class="tp">
  <div class="panel full"><div class="ph"><div class="pt">&#128680; Issue Log <span class="bx" id="ibx">0 issues</span></div></div><div id="ig"><div style="padding:16px;text-align:center;color:var(--s5)">Loading...</div></div></div>
</div>
<div id="tp-de" class="tp">
  <div class="panel full"><div class="ph"><div class="pt">&#128202; Decision Log <span class="bx" id="dbx">0 decisions</span></div></div><div id="dg"><div style="padding:16px;text-align:center;color:var(--s5)">Loading...</div></div></div>
</div>`;

const PM_BODY = `
<div class="krow k7">
  <div class="kpi"><div class="kl">Phase</div><div class="kv" id="kPh">2A</div><div class="ks">of 4 phases</div></div>
  <div class="kpi a"><div class="kl">% Complete</div><div class="kv a" id="kPct">44%</div><div class="ks">Phase 2A</div></div>
  <div class="kpi r"><div class="kl">Days to M-1</div><div class="kv r" id="kDays">0</div><div class="ks">Target: Apr 22</div></div>
  <div class="kpi r"><div class="kl">Past Due</div><div class="kv r" id="kPD">0</div><div class="ks">Milestones overdue</div></div>
  <div class="kpi g"><div class="kl">Budget Spent</div><div class="kv g" id="kBud">$4.6K</div><div class="ks">of $25K ceiling</div></div>
  <div class="kpi a"><div class="kl">Open Items</div><div class="kv a" id="kOI">0</div><div class="ks">Requires action</div></div>
  <div class="kpi"><div class="kl">Docs Delivered</div><div class="kv" id="kDD">0/0</div><div class="ks">of total deliverables</div></div>
</div>
<div class="g21">
  <div class="panel">
    <div class="ph"><div class="pt">&#128197; Phase 2A Milestone Schedule <span class="bx" id="mbx">0 milestones</span></div></div>
    <div id="mt"><div style="padding:16px 18px;font-size:12px;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="panel">
    <div class="ph"><div class="pt">&#128176; Budget</div></div>
    <div id="bd"><div style="padding:14px;text-align:center;color:var(--s5)">Loading...</div></div>
  </div>
</div>
<div class="g2">
  <div class="panel">
    <div class="ph"><div class="pt">&#128680; Open Items <span class="bx r" id="oibx">0 items</span></div></div>
    <div id="oig"><div style="padding:14px;text-align:center;color:var(--s5)">Loading...</div></div>
  </div>
  <div class="panel">
    <div class="ph"><div class="pt">&#128196; Document Deliverables <span class="bx" id="ddbx">0 docs</span></div></div>
    <div id="ddg"><div style="padding:14px;text-align:center;color:var(--s5)">Loading...</div></div>
  </div>
</div>`;

// ---------------------------------------------------------------------------
// WEB PART CLASS
// ---------------------------------------------------------------------------

export default class IMP9177DashboardWebPart extends BaseClientSideWebPart<IIMP9177DashboardWebPartProps> {

  private iframe: HTMLIFrameElement | null = null;
  private currentDashboard = 'MRS';

  // ---------------------------------------------------------------------------
  // SharePoint REST helper
  // ---------------------------------------------------------------------------
  private spGet(list: string, select: string = '', top: number = 200): Promise<any[]> {
    const base = this.context.pageContext.web.absoluteUrl;
    const url = `${base}/_api/web/lists/getbytitle('${list}')/items?$top=${top}${select ? '&$select=' + select : ''}&$orderby=Id asc`;
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1)
      .then((r: SPHttpClientResponse) => r.json())
      .then((d: any) => d.value || d.d?.results || [])
      .catch(() => []);
  }

  // ---------------------------------------------------------------------------
  // Render entry point
  // ---------------------------------------------------------------------------
  public render(): void {
    const dashboard = this.properties.dashboard || 'MRS';
    this.currentDashboard = dashboard;

    // Build the pure HTML shell (no scripts)
    let html = '';
    if (dashboard === 'MRS') {
      html = makeHtml('MRS', 'Master Record Sheet',
        'IMP9177 &middot; 3H Pharmaceuticals LLC &middot; 21 CFR Part 111 / FSMA &middot; Phase 2A',
        MRS_BODY,
        'IMP9177 &middot; ADB Consulting &amp; CRO Inc. &middot; Master Record Sheet &middot; Read-only view &middot; Data live from SharePoint lists');
    } else if (dashboard === 'RAID') {
      html = makeHtml('RAID', 'RAID Register',
        'IMP9177 &middot; Risks &middot; Actions &middot; Issues &middot; Decisions &middot; Meetings',
        RAID_BODY,
        'IMP9177 RAID Register &middot; ADB Consulting &amp; CRO Inc. &middot; Read-only view &middot; Live data from SharePoint lists');
    } else {
      html = makeHtml('PM', 'Project Management Tracker',
        'IMP9177 &middot; 3H Pharmaceuticals LLC &middot; ADB Consulting &amp; CRO Inc.',
        PM_BODY,
        'IMP9177 PM Tracker &middot; ADB Consulting &amp; CRO Inc. &middot; Next meeting: week of Apr 21 &middot; Tina travel Apr 13&ndash;17');
    }

    // Write HTML into iframe (no scripts = no CSP issues)
    this.domElement.innerHTML = '';
    this.iframe = document.createElement('iframe');
    this.iframe.style.cssText = 'width:100%;height:980px;border:none;display:block;';
    this.iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin allow-forms allow-popups');
    this.domElement.appendChild(this.iframe);

    const ifrDoc = this.iframe.contentDocument || (this.iframe.contentWindow && this.iframe.contentWindow.document);
    if (ifrDoc) {
      ifrDoc.open();
      ifrDoc.write(html);
      ifrDoc.close();
    } else {
      this.iframe.srcdoc = html;
    }

    // Wire up refresh button and tab buttons via contentDocument DOM manipulation
    // (no inline JS needed -- we attach listeners from TypeScript)
    const attachListeners = (): void => {
      const d = this.iframe && this.iframe.contentDocument;
      if (!d) return;
      const refreshBtn = d.getElementById('refreshBtn');
      if (refreshBtn) {
        refreshBtn.addEventListener('click', () => this._loadAndRender(dashboard));
      }
      // RAID tab buttons
      ['ov', 'ac', 'is', 'de'].forEach(tab => {
        const btn = d.getElementById('tab-' + tab);
        if (btn) {
          btn.addEventListener('click', () => {
            d.querySelectorAll('.tp').forEach(e => e.classList.remove('on'));
            d.querySelectorAll('.tbn').forEach(e => e.classList.remove('on'));
            const tp = d.getElementById('tp-' + tab);
            if (tp) tp.classList.add('on');
            btn.classList.add('on');
          });
        }
      });
    };

    // Load data and render -- wait for iframe to be ready
    const d2 = this.iframe.contentDocument;
    if (d2 && d2.readyState === 'complete') {
      attachListeners();
      this._loadAndRender(dashboard);
    } else {
      this.iframe.addEventListener('load', () => {
        attachListeners();
        this._loadAndRender(dashboard);
      }, { once: true });
    }
  }

  // ---------------------------------------------------------------------------
  // Fetch all data then render into iframe DOM
  // ---------------------------------------------------------------------------
  private _loadAndRender(dashboard: string): void {
    this._fetchAll(dashboard)
      .then(data => this._renderAll(dashboard, data))
      .catch(() => this._renderFallback(dashboard));
  }

  private async _fetchAll(dashboard: string): Promise<Record<string, any[]>> {
    if (dashboard === 'MRS') {
      const [docs, gaps, suppliers, ingredients, capas] = await Promise.all([
        this.spGet('MRS_Documents', 'DocNumber,DocTitle,DocType,DocWeek,DocStatus,ADBComplete,DocURL,DocRev'),
        this.spGet('MRS_GapRegistry', 'GapID,GapSeries,GapSeverity,GapOwner,GapStatus', 100),
        this.spGet('MRS_SupplierRegistry', 'SUPNumber,SupplierName,SupplierCat,SUPQualStatus', 25),
        this.spGet('MRS_IngredientRegistry', 'INGNumber,INGName,INGCoAStatus', 110),
        this.spGet('MRS_CAPALog', 'CAPANumber,CAPAType,CAPADesc,CAPAStatus', 20),
      ]);
      return { docs, gaps, suppliers, ingredients, capas };
    }
    if (dashboard === 'RAID') {
      const [actions, issues, decisions, meetings] = await Promise.all([
        this.spGet('RAID_Actions', 'ActionID,ActionCat,ActionDesc,ActionOwner,ActionPriority,ActionStatus'),
        this.spGet('RAID_Issues', 'IssueID,IssueTitle,IssueSeverity,IssueCategory,IssueOwner,IssueStatus'),
        this.spGet('RAID_Decisions', 'DecisionID,DecisionDate,DecisionMtg,DecisionMade,DecisionBy,DecisionImpact'),
        this.spGet('RAID_Meetings', 'MeetingID,MeetingDate', 50),
      ]);
      return { actions, issues, decisions, meetings };
    }
    if (dashboard === 'PM') {
      const [milestones, openItems, deliverables, budget] = await Promise.all([
        this.spGet('PM_Milestones', 'MilestoneID,MilestonePhase,MilestoneDesc,MilestoneStatus,MilestonePct,MilestoneTarget'),
        this.spGet('PM_OpenItems', 'OIRef,OITitle,OIOwner,OIPriority,OIStatus'),
        this.spGet('PM_DocumentDeliverables', 'PMDocID,PMDocNumber,PMDocWeek,PMDocTitle,PMDocStatus'),
        this.spGet('PM_Budget', 'Title,BudgetSpent,BudgetCeiling,BudgetInvoice', 5),
      ]);
      return { milestones, openItems, deliverables, budget };
    }
    return {};
  }

  // ---------------------------------------------------------------------------
  // Render dispatchers
  // ---------------------------------------------------------------------------
  private _renderAll(dashboard: string, data: Record<string, any[]>): void {
    const keys = Object.keys(data); let hasData = false; for (let ki = 0; ki < keys.length; ki++) { if ((data[keys[ki]] || []).length > 0) { hasData = true; break; } }
    if (!hasData) {
      this._renderFallback(dashboard);
      return;
    }
    if (dashboard === 'MRS') {
      this._renderDocs(data.docs || []);
      this._renderGaps(data.gaps || []);
      this._renderSuppliers(data.suppliers || []);
      this._renderIngredients(data.ingredients || []);
      this._renderCapas(data.capas || []);
    } else if (dashboard === 'RAID') {
      this._renderActions(data.actions || []);
      this._renderIssues(data.issues || []);
      this._renderDecisions(data.decisions || []);
      this._renderMeetings(data.meetings || []);
    } else if (dashboard === 'PM') {
      this._renderMilestones(data.milestones || []);
      this._renderOpenItems(data.openItems || []);
      this._renderDocs_PM(data.deliverables || []);
      this._renderBudget(data.budget || []);
    }
    this._finalizeRender();
  }

  private _finalizeRender(): void {
    const d = this.iframe && this.iframe.contentDocument;
    if (!d) return;
    const rt = d.getElementById('rt');
    if (rt) rt.textContent = 'Refreshed ' + new Date().toLocaleTimeString();
  }

  // ---------------------------------------------------------------------------
  // Fallback data (used when lists don't exist yet)
  // ---------------------------------------------------------------------------
  private _renderFallback(dashboard: string): void {
    if (dashboard === 'MRS') {
      this._renderDocs([
        { DocNumber: 'SOP-QMS-001', DocTitle: 'Management Responsibility SOP', DocType: 'SOP', DocWeek: 'W1', DocStatus: 'PAST DUE -- Sign-Off Overdue', ADBComplete: 'Yes' },
        { DocNumber: 'QM-001', DocTitle: 'Quality Manual', DocType: 'Quality Manual', DocWeek: 'W1', DocStatus: 'PAST DUE -- Sign-Off Overdue', ADBComplete: 'Yes' },
        { DocNumber: 'DCO-0002', DocTitle: 'Document Change Order -- W2 Package', DocType: 'DCO', DocWeek: 'W2', DocStatus: 'Pending Feedback', ADBComplete: 'Yes' },
        { DocNumber: 'SOP-SUP-001', DocTitle: 'Supplier Qualification SOP', DocType: 'SOP', DocWeek: 'W2', DocStatus: 'Pending Feedback', ADBComplete: 'Yes' },
        { DocNumber: 'FM-008', DocTitle: 'Supplier CoA Requirements Checklist', DocType: 'Form', DocWeek: 'W2', DocStatus: 'Pending Feedback', ADBComplete: 'Yes' },
      ]);
      this._renderGaps([]);
      this._renderSuppliers([]);
      this._renderIngredients([]);
      this._renderCapas([{ CAPANumber: 'CAPA-0001', CAPAType: 'CAPA', CAPADesc: 'Systemic CoA Deficiency -- Supplier Qualification Program', CAPAStatus: 'OPEN -- In Progress' }]);
    } else if (dashboard === 'RAID') {
      this._renderActions([
        { ActionID: 'AC-034', ActionCat: 'Compliance', ActionDesc: 'QD Yang signatory -- BLOCKING DCO-0001', ActionOwner: 'ADB (Andre)', ActionPriority: 'Critical', ActionStatus: 'BLOCKING' },
        { ActionID: 'AC-029', ActionCat: 'Document Control', ActionDesc: 'Issue DCO-0001 sign-off package to 3H', ActionOwner: '3H -- Tina Qin', ActionPriority: 'Critical', ActionStatus: 'PAST DUE' },
        { ActionID: 'AC-023', ActionCat: 'Compliance', ActionDesc: 'Request Orkin stored product pest addendum', ActionOwner: '3H -- Tina Qin', ActionPriority: 'Critical', ActionStatus: 'Open -- Urgent' },
        { ActionID: 'ISS-010', ActionCat: 'Compliance', ActionDesc: 'Melatonin Powder -- No CoA on file (active production)', ActionOwner: '3H -- Tina Qin', ActionPriority: 'Critical', ActionStatus: 'Open -- Critical' },
      ]);
      this._renderIssues([
        { IssueID: 'ISS-010', IssueTitle: 'Melatonin Powder -- No CoA on file', IssueSeverity: 'Critical', IssueCategory: 'Compliance', IssueOwner: '3H -- Tina Qin', IssueStatus: 'Open -- Critical' },
        { IssueID: 'ISS-002', IssueTitle: 'Active Pest Infestation', IssueSeverity: 'Critical', IssueCategory: 'Facility', IssueOwner: '3H Management', IssueStatus: 'Partially Mitigated' },
      ]);
      this._renderDecisions([
        { DecisionID: 'DEC-040', DecisionMtg: 'MM-005', DecisionMade: 'Tina Qin confirmed as interim QA Approver. QD Yang = operational QA duties only.', DecisionBy: 'Andre Butler + Tina Qin + Liu Hong', DecisionImpact: 'Critical -- Blocking' },
        { DecisionID: 'DEC-035', DecisionMtg: 'MM-005', DecisionMade: 'Amazon-sourced ingredients excluded from quality scope -- R&D use only. 18 production vendors confirmed.', DecisionBy: 'Andre Butler + Tina Qin', DecisionImpact: 'High' },
      ]);
      this._renderMeetings([{}, {}, {}, {}, {}]);
    } else if (dashboard === 'PM') {
      this._renderMilestones([
        { MilestoneID: 'M-15', MilestoneDesc: 'MILESTONE 1: Phase 2A Complete', MilestonePct: '44', MilestoneTarget: 'Apr 22', MilestoneStatus: 'In Progress' },
        { MilestoneID: 'W2-DEL', MilestoneDesc: 'W2 -- 15-Document Package Delivered', MilestonePct: '100', MilestoneTarget: 'Apr 9', MilestoneStatus: 'Complete' },
        { MilestoneID: 'W1-SO', MilestoneDesc: 'CR-0001/DCO-0001 -- 3H Sign-Off', MilestonePct: '0', MilestoneTarget: 'Apr 5', MilestoneStatus: 'At Risk -- PAST DUE' },
        { MilestoneID: 'MM-005', MilestoneDesc: 'W1/W2 Review + Supplier Controls', MilestonePct: '100', MilestoneTarget: 'Apr 7', MilestoneStatus: 'Complete' },
      ]);
      this._renderBudget([{ BudgetSpent: '4617', BudgetCeiling: '25000', BudgetInvoice: 'INV-IMP9177002' }]);
      this._renderOpenItems([
        { OIRef: 'ISS-010', OITitle: 'Melatonin Powder -- No CoA. Active production. SS111.75(a)(1).', OIOwner: '3H -- Tina Qin', OIPriority: 'Critical', OIStatus: 'Open -- Critical' },
        { OIRef: 'DCO-0001', OITitle: 'W1 Sign-off PAST DUE -- Tina + Liu Hong must sign', OIOwner: '3H -- Tina Qin', OIPriority: 'Critical', OIStatus: 'At Risk -- Past Due' },
        { OIRef: 'AC-023', OITitle: 'Orkin pest addendum -- URGENT before inspection', OIOwner: '3H -- Tina Qin', OIPriority: 'Critical', OIStatus: 'Open -- Urgent' },
        { OIRef: 'W2-FB', OITitle: 'W2 feedback due Apr 14', OIOwner: '3H -- Tina Qin', OIPriority: 'High', OIStatus: 'In Progress' },
      ]);
      this._renderDocs_PM([
        { PMDocNumber: 'DCO-0001', PMDocTitle: 'W1 Package Sign-Off', PMDocWeek: 'W1', PMDocStatus: 'PAST DUE -- Sign-Off Overdue' },
        { PMDocNumber: 'QM-001', PMDocTitle: 'Quality Manual', PMDocWeek: 'W1', PMDocStatus: 'PAST DUE -- Sign-Off Overdue' },
        { PMDocNumber: 'SOP-SUP-001', PMDocTitle: 'Supplier Qualification SOP', PMDocWeek: 'W2', PMDocStatus: 'Delivered -- Pending Feedback' },
        { PMDocNumber: 'SOP-FS-001', PMDocTitle: 'Allergen Control SOP', PMDocWeek: 'W2', PMDocStatus: 'Delivered -- Pending Feedback' },
      ]);
    }
    this._finalizeRender();
  }

  // ---------------------------------------------------------------------------
  // DOM helpers
  // ---------------------------------------------------------------------------
  private _el(id: string): HTMLElement | null {
    return (this.iframe && this.iframe.contentDocument && this.iframe.contentDocument.getElementById(id)) || null;
  }
  private _set(id: string, val: string): void {
    const el = this._el(id); if (el) el.textContent = val;
  }
  private _html(id: string, val: string): void {
    const el = this._el(id); if (el) el.innerHTML = val;
  }
  private _pc(s: string): string {
    s = (s || '').toLowerCase();
    if (s.includes('past due') || s.includes('overdue') || s.includes('critical') || s.includes('blocking')) return 'pr';
    if (s.includes('pending') || s.includes('feedback') || s.includes('progress') || s.includes('urgent') || s.includes('new')) return 'pa';
    if (s.includes('complete') || s.includes('signed') || s.includes('closed')) return 'pg';
    if (s.includes('delivered') || s.includes('adb')) return 'pb';
    return 'py';
  }
  private _pill(t: string): string {
    return `<span class="pill ${this._pc(t)}">${(t || '').substring(0, 26)}</span>`;
  }
  private _tbl(heads: string[], rows: string): string {
    return `<table><thead><tr>${heads.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>${rows}</tbody></table>`;
  }
  private _grp(lbl: string, cls: string, n: number): string {
    return `<div class="gh ${cls}"><span>${lbl}</span><span style="opacity:.7">${n} item${n !== 1 ? 's' : ''}</span></div>`;
  }

  // ---------------------------------------------------------------------------
  // MRS render functions
  // ---------------------------------------------------------------------------
  private _renderDocs(items: any[]): void {
    const n = items.length;
    this._set('kD', String(n));
    this._html('dbx', n + ' docs');
    const pd = items.filter(i => (i.DocStatus || '').toUpperCase().includes('PAST DUE') || (i.DocStatus || '').toUpperCase().includes('OVERDUE'));
    const pf = items.filter(i => pd.indexOf(i) === -1 && ((i.DocStatus || '').includes('Pending') || (i.DocStatus || '').includes('Feedback')));
    const ot = items.filter(i => pd.indexOf(i) === -1 && pf.indexOf(i) === -1);
    this._set('kP', String(pd.length));
    this._set('kF', String(pf.length));
    this._set('kA', String(items.filter(i => i.ADBComplete === 'Yes').length));
    const SP = 'https://adbccro.sharepoint.com/sites/IMP9177';
    const row = (i: any): string => {
      const url = i.DocURL || `${SP}/Shared%20Documents/QMS/Documents/${encodeURIComponent(i.DocNumber || '')}_Rev${i.DocRev || 'A'}.docx`;
      return `<tr><td><a class="dlink" href="${url}" target="_blank">${i.DocNumber || ''}</a></td><td style="font-size:12px">${i.DocTitle || i.Title || ''}</td><td style="font-size:11px;color:var(--s5)">${i.DocType || ''}</td><td><span style="font-size:10px;background:var(--b1);color:#1E56A0;padding:1px 6px;border-radius:3px">${i.DocWeek || ''}</span></td><td>${this._pill(i.DocStatus || '')}</td></tr>`;
    };
    const grp = (lbl: string, cls: string, rows: any[]): string =>
      rows.length ? this._grp(lbl, cls, rows.length) + this._tbl(['Doc #', 'Title', 'Type', 'Week', 'Status'], rows.map(row).join('')) : '';
    this._html('dg', grp('Sign-Off Past Due', 'red', pd) + grp('Pending Feedback / Sign-Off', 'amb', pf) + (ot.length ? grp('Other', 'gry', ot) : ''));
  }

  private _renderGaps(items: any[]): void {
    const c = items.filter(i => i.GapSeverity === 'Critical').length || 30;
    const m = items.filter(i => i.GapSeverity === 'Major').length || 31;
    const n = items.filter(i => i.GapSeverity === 'Minor').length || 6;
    const r = items.filter(i => (i.GapSeries || '').includes('REC')).length || 10;
    this._set('gc', String(c)); this._set('gm', String(m)); this._set('gn', String(n)); this._set('gr', String(r));
    const closed = items.filter(i => (i.GapStatus || '').includes('Closed')).length;
    this._set('go', String((items.length || 77) - closed));
    this._set('gcl', String(closed));
    // SVG donut
    const total = c + m + n + r;
    const cols = ['#DC2626', '#D97706', '#059669', '#3B82D4'];
    const vals = [c, m, n, r];
    const R = 38, cx = 53, cy = 53, C = 2 * Math.PI * R;
    let off = 0;
    const arcs = vals.map((v, i) => {
      const pct = v / total, len = pct * C, rot = off * 360 - 90; off += pct;
      return `<circle cx="${cx}" cy="${cy}" r="${R}" fill="none" stroke="${cols[i]}" stroke-width="15" stroke-dasharray="${len} ${C - len}" transform="rotate(${rot} ${cx} ${cy})"/>`;
    }).join('');
    this._html('ga', arcs);
  }

  private _renderSuppliers(items: any[]): void {
    const n = items.length || 18;
    this._html('sbx', String(n));
    const pend = items.filter(i => (i.SUPQualStatus || 'PENDING').includes('PENDING')).length || n;
    const appr = items.filter(i => (i.SUPQualStatus || '').includes('APPROVED') || i.SUPQualStatus === 'Active').length;
    const cats: Record<string, number> = {};
    (items.length ? items : Array(18).fill({ SupplierCat: 'Raw Materials' }).map((x, i) => i >= 16 ? { SupplierCat: i === 16 ? 'Packaging Materials' : 'Laboratory Consumables' } : x))
      .forEach((i: any) => { const c = i.SupplierCat || 'Other'; cats[c] = (cats[c] || 0) + 1; });
    const catHtml = Object.keys(cats).sort((a, b) => (cats[b] as number) - (cats[a] as number)).map((cat: string) => {
      const cnt = cats[cat] as number;
      return `<div class="scat"><span style="font-size:11px;color:var(--s7)">${cat}</span><span style="font-size:11px;font-weight:700;color:var(--n)">${cnt}</span></div><div class="scatp"><div class="scatb" style="width:${Math.round(cnt / n * 100)}%"></div></div>`;
    }).join('');
    this._html('ss', `<div class="sup2"><div class="sbox"><div class="sbig" style="color:var(--a)">${pend}</div><div class="ssub">Pending Qualification</div></div><div class="sbox"><div class="sbig" style="color:var(--g)">${appr}</div><div class="ssub">Approved</div></div></div><div class="scats">${catHtml}</div>`);
  }

  private _renderIngredients(items: any[]): void {
    const n = items.length || 101;
    this._html('ibx', String(n));
    const onFile = items.filter(i => (i.INGCoAStatus || '').toLowerCase().includes('on file')).length || 35;
    const missing = items.filter(i => (i.INGCoAStatus || '').toLowerCase().includes('missing')).length || (n - onFile - 1);
    const crit = items.filter(i => (i.INGCoAStatus || '').includes('CRITICAL')).length || 1;
    const bars = [
      { l: 'CoA on File', n: onFile, c: '#059669' },
      { l: 'CoA Missing', n: missing, c: '#D97706' },
      { l: 'Critical Alerts', n: crit, c: '#DC2626' },
      { l: 'Spec Required', n: n, c: '#3B82D4' },
    ];
    this._html('ib', bars.map(b =>
      `<div class="dr"><span class="dlb">${b.l}</span><div class="dt"><div class="df" style="width:${Math.round(b.n / n * 100)}%;background:${b.c}"><span class="dn">${b.n > 0 ? b.n : ''}</span></div></div><span class="dnum">${b.n}</span></div>`
    ).join(''));
  }

  private _renderCapas(items: any[]): void {
    if (!items || !items.length) { items = [{ CAPANumber: 'CAPA-0001', CAPAType: 'CAPA', CAPADesc: 'Systemic CoA Deficiency -- Supplier Qualification Program Remediation', CAPAStatus: 'OPEN -- In Progress' }]; }
    const open = items.filter(i => !(i.CAPAStatus || '').toLowerCase().includes('closed')).length;
    this._set('kC', String(open));
    this._html('cbx', (items.length || 0) + ' record' + (items.length !== 1 ? 's' : ''));
    this._html('cc', this._tbl(['CAPA #', 'Type', 'Description', 'Status'],
      items.map(i => `<tr><td class="did">${i.CAPANumber || ''}</td><td style="font-size:11px">${i.CAPAType || ''}</td><td style="font-size:11px;max-width:180px">${(i.CAPADesc || '').substring(0, 60)}${(i.CAPADesc || '').length > 60 ? '...' : ''}</td><td>${this._pill(i.CAPAStatus || '')}</td></tr>`).join('')));
  }

  // ---------------------------------------------------------------------------
  // RAID render functions
  // ---------------------------------------------------------------------------
  private _renderActions(items: any[]): void {
    const crit = items.filter(i => i.ActionPriority === 'Critical' || i.ActionStatus === 'BLOCKING');
    const past = items.filter(i => (i.ActionStatus || '').includes('Past Due') || (i.ActionStatus || '').includes('PAST DUE'));
    const inprog = items.filter(i => crit.indexOf(i) === -1 && past.indexOf(i) === -1 && !(i.ActionStatus || '').toLowerCase().includes('complete') && !(i.ActionStatus || '').toLowerCase().includes('closed'));
    const done = items.filter(i => (i.ActionStatus || '').toLowerCase().includes('complete') || (i.ActionStatus || '').toLowerCase().includes('closed'));
    this._set('kAt', String(items.length));
    this._html('kAo', (items.length - done.length) + ' open');
    this._set('kAc', String(crit.length));
    this._set('hc', String(crit.length));
    this._set('hp', String(past.length));
    this._set('hd', String(done.length));
    this._html('cbx', crit.length + ' items');
    this._html('abx', items.length + ' actions');
    const row = (i: any): string => `<tr><td class="did">${i.ActionID || ''}</td><td style="font-size:11px;color:var(--s5)">${i.ActionCat || ''}</td><td style="font-size:11px;max-width:280px">${(i.ActionDesc || '').substring(0, 100)}${(i.ActionDesc || '').length > 100 ? '...' : ''}</td><td style="font-size:11px">${(i.ActionOwner || '').replace('ADB (Andre)', 'ADB').replace('3H -- ', '')}</td><td>${this._pill(i.ActionStatus || '')}</td></tr>`;
    const heads = ['ID', 'Category', 'Description', 'Owner', 'Status'];
    const critRow = (i: any): string => `<tr><td class="did">${i.ActionID || ''}</td><td style="font-size:11px;max-width:240px">${(i.ActionDesc || '').substring(0, 80)}...</td><td style="font-size:11px">${(i.ActionOwner || '').replace('ADB (Andre)', 'ADB').replace('3H -- ', '')}</td><td>${this._pill(i.ActionStatus || '')}</td></tr>`;
    this._html('ct', crit.length ? this._tbl(['ID', 'Description', 'Owner', 'Status'], crit.map(critRow).join('')) : '<div style="padding:14px;text-align:center;color:var(--s5)">No critical actions</div>');
    this._html('ag',
      (crit.length ? this._grp('Critical / Blocking', 'red', crit.length) + this._tbl(heads, crit.map(row).join('')) : '') +
      (past.filter(i => crit.indexOf(i) === -1).length ? this._grp('Past Due', 'amb', past.filter(i => crit.indexOf(i) === -1).length) + this._tbl(heads, past.filter(i => crit.indexOf(i) === -1).map(row).join('')) : '') +
      (inprog.length ? this._grp('In Progress / Open', 'blu', inprog.length) + this._tbl(heads, inprog.map(row).join('')) : '') +
      (done.length ? this._grp('Completed / Closed', 'grn', done.length) + this._tbl(heads, done.map(row).join('')) : ''));
  }

  private _renderIssues(items: any[]): void {
    const open = items.filter(i => !(i.IssueStatus || '').toLowerCase().includes('closed')).length;
    this._set('kIo', String(open));
    this._set('hi', String(open));
    this._html('ibx', items.length + ' issues');
    const crit = items.filter(i => i.IssueSeverity === 'Critical');
    const maj = items.filter(i => i.IssueSeverity === 'Major' && crit.indexOf(i) === -1);
    const oth = items.filter(i => crit.indexOf(i) === -1 && maj.indexOf(i) === -1);
    const row = (i: any): string => `<tr><td class="did">${i.IssueID || ''}</td><td style="font-weight:600;font-size:12px">${i.IssueTitle || ''}</td><td style="font-size:11px;color:var(--s5)">${i.IssueCategory || ''}</td><td style="font-size:11px">${(i.IssueOwner || '').replace('3H -- ', '')}</td><td>${this._pill(i.IssueStatus || '')}</td></tr>`;
    const heads = ['ID', 'Issue', 'Category', 'Owner', 'Status'];
    this._html('ig',
      (crit.length ? this._grp('Critical', 'red', crit.length) + this._tbl(heads, crit.map(row).join('')) : '') +
      (maj.length ? this._grp('Major', 'amb', maj.length) + this._tbl(heads, maj.map(row).join('')) : '') +
      (oth.length ? this._grp('Other', 'gry', oth.length) + this._tbl(heads, oth.map(row).join('')) : '') ||
      '<div style="padding:16px;text-align:center;color:var(--s5)">No issues logged</div>');
  }

  private _renderDecisions(items: any[]): void {
    this._set('kDl', String(items.length));
    this._html('dbx', items.length + ' decisions');
    const byMtg: Record<string, any[]> = {};
    items.forEach(i => { const m = i.DecisionMtg || 'General'; (byMtg[m] = byMtg[m] || []).push(i); });
    const mtgOrder = ['MM-005', 'MM-004', 'MM-003', 'MM-002', 'MM-001', 'General'];
    const row = (i: any): string => `<tr><td class="did">${i.DecisionID || ''}</td><td style="font-size:11px;max-width:320px">${(i.DecisionMade || '').substring(0, 120)}${(i.DecisionMade || '').length > 120 ? '...' : ''}</td><td style="font-size:11px">${(i.DecisionBy || '').substring(0, 40)}</td><td>${this._pill(i.DecisionImpact || '')}</td></tr>`;
    const heads = ['ID', 'Decision', 'Made By', 'Impact'];
    const sorted = [...mtgOrder.filter(m => byMtg[m]), ...Object.keys(byMtg).filter(m => mtgOrder.indexOf(m) === -1)];
    this._html('dg', sorted.map(m =>
      this._grp(m, m === 'MM-005' || m === 'MM-004' ? 'grn' : 'blu', byMtg[m].length) + this._tbl(heads, byMtg[m].map(row).join(''))
    ).join('') || '<div style="padding:16px;text-align:center;color:var(--s5)">No decisions logged</div>');
  }

  private _renderMeetings(items: any[]): void {
    this._set('kMh', String(items.length || 5));
  }

  // ---------------------------------------------------------------------------
  // PM render functions
  // ---------------------------------------------------------------------------
  private _fmtDate(val: string): string {
    if (!val) return '';
    const m = val.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) {
      const months = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
      return months[parseInt(m[2], 10) - 1] + ' ' + parseInt(m[3], 10);
    }
    return val;
  }

  private _renderMilestones(items: any[]): void {
    this._html('mbx', items.length + ' milestones');
    const pd = items.filter(i => (i.MilestoneStatus || '').includes('Past Due') || (i.MilestoneStatus || '').includes('Risk'));
    this._set('kPD', String(pd.length));
    const pcts = items.map(i => parseFloat(i.MilestonePct) || 0).filter(n => n > 0);
    const avgPct = pcts.length ? Math.round(pcts.reduce((a, b) => a + b, 0) / pcts.length) : 44;
    this._set('kPct', avgPct + '%');
    const m1 = new Date('2026-04-22'), now = new Date();
    const days = Math.ceil((m1.getTime() - now.getTime()) / (1000 * 60 * 60 * 24));
    this._set('kDays', String(Math.max(0, days)));
    const hdays = this._el('hdays');
    if (hdays) { hdays.textContent = days > 0 ? days + ' days to M-1' : 'M-1 PAST DUE'; }
    this._html('mt', this._tbl(['ID', 'Milestone', 'Progress', 'Target', 'Status'],
      [...items].sort((a, b) => (b.MilestoneID || '').localeCompare(a.MilestoneID || '')).map(i => {
        const pct = parseFloat(i.MilestonePct) || 0;
        const sc = pct >= 100 ? 'pg' : i.MilestoneStatus && (i.MilestoneStatus.includes('Past Due') || i.MilestoneStatus.includes('Risk')) ? 'pr' : i.MilestoneStatus && i.MilestoneStatus.includes('Progress') ? 'pa' : 'py';
        return `<tr><td class="did">${i.MilestoneID || ''}</td><td style="font-size:12px;font-weight:600">${(i.MilestoneDesc || '').substring(0, 40)}${(i.MilestoneDesc || '').length > 40 ? '...' : ''}</td><td style="min-width:90px"><div style="font-size:11px;color:var(--s5);margin-bottom:2px">${pct}%</div><div class="pbar"><div class="pbf" style="width:${pct}%;background:${pct >= 100 ? 'var(--g)' : pct >= 50 ? 'var(--b5)' : 'var(--a)'}"></div></div></td><td style="font-size:11px;color:var(--s5)">${this._fmtDate(i.MilestoneTarget || '')}</td><td><span class="pill ${sc}">${(i.MilestoneStatus || '').substring(0, 22)}</span></td></tr>`;
      }).join('')));
  }

  private _renderOpenItems(items: any[]): void {
    this._set('kOI', String(items.filter(i => !(i.OIStatus || '').toLowerCase().includes('closed')).length));
    this._html('oibx', items.length + ' items');
    if (!items.length) { this._html('oig', '<div style="padding:14px;text-align:center;color:var(--s5)">No open items</div>'); return; }
    const crit = items.filter(i => i.OIPriority === 'Critical');
    const high = items.filter(i => i.OIPriority === 'High' && crit.indexOf(i) === -1);
    const oth = items.filter(i => crit.indexOf(i) === -1 && high.indexOf(i) === -1);
    const row = (i: any): string => `<tr><td class="did">${i.OIRef || ''}</td><td style="font-size:11px;max-width:220px">${(i.OITitle || '').substring(0, 80)}${(i.OITitle || '').length > 80 ? '...' : ''}</td><td style="font-size:11px">${(i.OIOwner || '').replace('3H -- ', '').replace('ADB (Andre)', 'ADB')}</td><td>${this._pill(i.OIStatus || '')}</td></tr>`;
    const heads = ['Ref', 'Item', 'Owner', 'Status'];
    this._html('oig',
      (crit.length ? this._grp('Critical', 'red', crit.length) + this._tbl(heads, crit.map(row).join('')) : '') +
      (high.length ? this._grp('High Priority', 'amb', high.length) + this._tbl(heads, high.map(row).join('')) : '') +
      (oth.length ? this._grp('Other', 'gry', oth.length) + this._tbl(heads, oth.map(row).join('')) : ''));
  }

  private _renderDocs_PM(items: any[]): void {
    if (!items || !items.length) { this._html('ddg', '<div style="padding:14px;text-align:center;color:var(--s5)">No deliverables loaded</div>'); return; }
    const delivered = items.filter(i => (i.PMDocStatus || '').toLowerCase().includes('delivered') || (i.PMDocStatus || '').toLowerCase().includes('complete')).length;
    this._set('kDD', delivered + '/' + items.length);
    this._html('ddbx', items.length + ' docs');
    const pd = items.filter(i => (i.PMDocStatus || '').toUpperCase().includes('PAST DUE') || (i.PMDocStatus || '').toUpperCase().includes('OVERDUE'));
    const pf = items.filter(i => pd.indexOf(i) === -1 && (i.PMDocStatus || '').includes('Pending Feedback'));
    const del = items.filter(i => pd.indexOf(i) === -1 && pf.indexOf(i) === -1 && (i.PMDocStatus || '').toLowerCase().includes('delivered'));
    const ns = items.filter(i => pd.indexOf(i) === -1 && pf.indexOf(i) === -1 && del.indexOf(i) === -1);
    const row = (i: any): string => `<tr><td class="did">${i.PMDocNumber || ''}</td><td style="font-size:11px">${(i.PMDocTitle || '').substring(0, 45)}${(i.PMDocTitle || '').length > 45 ? '...' : ''}</td><td><span style="font-size:10px;background:var(--b1);color:#1E56A0;padding:1px 6px;border-radius:3px">${i.PMDocWeek || ''}</span></td><td>${this._pill(i.PMDocStatus || '')}</td></tr>`;
    const heads = ['Doc #', 'Title', 'Week', 'Status'];
    this._html('ddg',
      (pd.length ? this._grp('Past Due -- Sign-Off Overdue', 'red', pd.length) + this._tbl(heads, pd.map(row).join('')) : '') +
      (pf.length ? this._grp('Pending Feedback', 'amb', pf.length) + this._tbl(heads, pf.map(row).join('')) : '') +
      (del.length ? this._grp('Delivered / Complete', 'grn', del.length) + this._tbl(heads, del.map(row).join('')) : '') +
      (ns.length ? this._grp('Planned', 'gry', ns.length) + this._tbl(heads, ns.map(row).join('')) : ''));
  }

  private _renderBudget(items: any[]): void {
    if (!items || !items.length) { items = [{ BudgetSpent: '4617', BudgetCeiling: '25000', BudgetInvoice: 'INV-IMP9177I002' }]; }
    const b = items[0];
    const spent = parseFloat((b.BudgetSpent || b.Title || '4617').toString().replace(/[^0-9.]/g, '')) || 4617;
    const ceiling = parseFloat((b.BudgetCeiling || '25000').toString().replace(/[^0-9.]/g, '')) || 25000;
    const pct = Math.round(spent / ceiling * 100);
    const rem = ceiling - spent;
    this._set('kBud', '$' + (spent / 1000).toFixed(1) + 'K');
    // PSO data: ceiling from contract, gaps = additional scope identified, invoiced = billed to date
    const gapAdded = parseFloat((b.BudgetGapsAdded || '0').toString().replace(/[^0-9.]/g, '')) || 0;
    const totalAuth = ceiling + gapAdded;
    const invoice1 = parseFloat((b.BudgetInv1 || '4617').toString().replace(/[^0-9.]/g, '')) || 4617;
    const invoice2 = parseFloat((b.BudgetInv2 || '0').toString().replace(/[^0-9.]/g, '')) || 0;
    const totalInvoiced = invoice1 + invoice2;
    const pctInv = Math.round(totalInvoiced / totalAuth * 100);
    const remAuth = totalAuth - totalInvoiced;
    this._set('kBud', '$' + (totalInvoiced / 1000).toFixed(1) + 'K');
    this._html('bd', `<div class="bcard">
      <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--s5);margin-bottom:10px">PSO-V7 &mdash; Budget Summary</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:14px">
        <div style="background:var(--b0);border-radius:6px;padding:10px 12px">
          <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:var(--b)">PSO Approved</div>
          <div style="font-size:20px;font-weight:700;color:var(--n);font-family:'Libre Baskerville',serif">$${ceiling.toLocaleString()}</div>
          <div style="font-size:10px;color:var(--s5);margin-top:2px">Signed Mar 18, 2026</div>
        </div>
        <div style="background:${gapAdded > 0 ? '#FEF3C7' : 'var(--s0)'};border-radius:6px;padding:10px 12px">
          <div style="font-size:9px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:${gapAdded > 0 ? '#92400E' : 'var(--s5)'}">Scope Gaps Added</div>
          <div style="font-size:20px;font-weight:700;color:${gapAdded > 0 ? 'var(--a)' : 'var(--s5)'};font-family:'Libre Baskerville',serif">${gapAdded > 0 ? '$' + gapAdded.toLocaleString() : 'None'}</div>
          <div style="font-size:10px;color:var(--s5);margin-top:2px">PSO-CO-001 pending</div>
        </div>
      </div>
      <div style="font-size:11px;font-weight:600;color:var(--s7);margin-bottom:6px">Invoiced to Date</div>
      <div class="btrack"><div class="bfill" style="width:${pctInv}%;background:var(--g)"></div></div>
      <div style="display:flex;justify-content:space-between;font-size:11px;color:var(--s5);margin-bottom:12px;margin-top:3px">
        <span>$${totalInvoiced.toLocaleString()} invoiced (${pctInv}%)</span><span>$${remAuth.toLocaleString()} remaining</span>
      </div>
      <div style="font-size:11px;font-weight:600;color:var(--s7);margin-bottom:6px">Invoice Detail</div>
      <div style="display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px;padding:6px 0;border-bottom:1px solid var(--s1)">
        <span style="color:var(--s7)">${b.BudgetInvoice || 'INV-IMP9177I002'}</span>
        <span style="display:flex;align-items:center;gap:8px"><span style="font-size:10px;color:var(--s5)">Phase 1 &mdash; Paid</span><span style="font-weight:700;color:var(--g)">$${invoice1.toLocaleString()}</span></span>
      </div>
      ${invoice2 > 0 ? `<div style="display:flex;justify-content:space-between;font-size:11px;margin-bottom:4px;padding:6px 0;border-bottom:1px solid var(--s1)"><span style="color:var(--s7)">${b.BudgetInv2Ref || 'INV-IMP9177I003'}</span><span style="display:flex;align-items:center;gap:8px"><span style="font-size:10px;color:var(--s5)">Phase 2A &mdash; Paid</span><span style="font-weight:700;color:var(--g)">$${invoice2.toLocaleString()}</span></span></div>` : '<div style="display:flex;justify-content:space-between;font-size:11px;padding:6px 0;color:var(--s5)"><span>Phase 2A</span><span>Not yet invoiced</span></div>'}
      <div style="display:flex;justify-content:space-between;font-size:12px;border-top:2px solid var(--s2);padding-top:8px;margin-top:8px"><span style="font-weight:600">Total Authorized</span><span style="font-weight:700;color:var(--n)">$${totalAuth.toLocaleString()}</span></div>
      <div class="bpso" style="margin-top:10px"><span class="bptag">PSO-CO-001</span><div style="font-size:11px;color:var(--s7);margin-top:2px">Phase 2 Change Order &mdash; Pending Execution</div><div style="font-size:10px;color:var(--s5);margin-top:2px">Phase 2B authorization blocked until CO executed</div></div>
    </div>`);
  }

  // ---------------------------------------------------------------------------
  // SPFx boilerplate
  // ---------------------------------------------------------------------------
  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [{
        header: { description: 'IMP9177 Dashboard' },
        groups: [{
          groupName: 'Dashboard Selection',
          groupFields: [
            PropertyPaneChoiceGroup('dashboard', {
              label: 'Select Dashboard to Display',
              options: [
                { key: 'MRS', text: 'Master Record Sheet (MRS)' },
                { key: 'RAID', text: 'RAID Register' },
                { key: 'PM', text: 'PM Tracker' },
              ]
            })
          ]
        }]
      }]
    };
  }
}
