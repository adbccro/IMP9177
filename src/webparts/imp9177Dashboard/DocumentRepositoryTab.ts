/**
 * DocumentRepositoryTab.ts
 * IMP9177 QMS Portal — Document Repository Matrix tab
 *
 * DROP-IN: Add renderDocumentRepositoryTab() call to the portal's tab switch block.
 * DATA:    Fetches live from Microsoft Graph (SP file metadata + version history).
 * ZONES:   Drafts   → /sites/IMP9177/Shared Documents/QMS/Documents/Drafts
 *                      /sites/IMP9177/Shared Documents/QMS/Forms/Drafts
 *          Published → /sites/IMP9177/Shared Documents/Published/QMS/Documents
 *                      /sites/IMP9177/Shared Documents/Published/QMS/Forms
 *                      /sites/IMP9177/Shared Documents/Published/QMS/Quality Manual
 *          Official  → /sites/IMP9177/Shared Documents/Official/QMS/Documents
 *                      /sites/IMP9177/Shared Documents/Official/QMS/Forms
 *
 * VERSIONING MODEL:
 *   Rev (A/B/C…)  = QMS lifecycle letter, changes via DCO
 *   Version (1.0) = SharePoint major version (publish event, whole number)
 *   Version (1.3) = SharePoint minor version (working save, decimal)
 *   Major check-in → publish → lands in Published zone as new major version
 *   DCO sign-off  → Power Automate promote → lands in Official zone
 */

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { WebPartContext } from "@microsoft/sp-webpart-base";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

interface ZoneEntry {
  rev: string;           // e.g. "A", "B"
  version: string;       // e.g. "1.0", "1.3"
  versionType: "major" | "minor";
  status: string;        // derived from DCO list or file metadata
  dco: string | null;    // e.g. "DCO-0001"
  dcoStatus: string;     // "blocked" | "open" | "signed" | "n/a"
  checkedOutTo: string | null;
  file: string;
  path: string;
  modified: string;
  webUrl: string;
  warn?: string;
  history: VersionEntry[];
}

interface VersionEntry {
  version: string;
  type: "major" | "minor";
  who: string;
  when: string;
  comment: string;
}

interface DocRecord {
  id: string;
  title: string;
  type: "SOP" | "FM" | "QM" | "FPS" | string;
  group: string;
  zones: {
    drafts: ZoneEntry | null;
    published: ZoneEntry | null;
    official: ZoneEntry | null;
  };
}

interface SpFileItem {
  Name: string;
  ServerRelativeUrl: string;
  TimeLastModified: string;
  CheckOutType: number;   // 0=none, 1=online, 2=offline
  CheckedOutByUser?: { LoginName: string; Title: string };
  MajorVersion: number;
  MinorVersion: number;
  UIVersion: number;
}

// ---------------------------------------------------------------------------
// Zone folder paths (relative to site root)
// ---------------------------------------------------------------------------

const SITE_URL = "https://adbccro.sharepoint.com/sites/IMP9177";

const ZONE_PATHS: Record<string, string[]> = {
  drafts: [
    "Shared Documents/QMS/Documents/Drafts",
    "Shared Documents/QMS/Forms/Drafts",
  ],
  published: [
    "Shared Documents/Published/QMS/Documents",
    "Shared Documents/Published/QMS/Forms",
    "Shared Documents/Published/QMS/Quality Manual",
  ],
  official: [
    "Shared Documents/Official/QMS/Documents",
    "Shared Documents/Official/QMS/Forms",
  ],
};

// DCO status map — sourced from QMS_DCOs list
// Updated here until live list query is wired
const DCO_STATUS: Record<string, string> = {
  "DCO-0001": "blocked",
  "DCO-0002": "open",
};

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function extractDocId(filename: string): string {
  // e.g. "SOP-QMS-001_RevA_Management_Responsibility.docx" → "SOP-QMS-001"
  const m = filename.match(/^([A-Z]{2,5}-[A-Z]{2,5}-\d{3}|[A-Z]{2,5}-\d{3}|FM-[A-Z]{2,5}|QM-\d{3}|FPS-\d{3})/i);
  return m ? m[1].toUpperCase() : filename.replace(/\.[^.]+$/, "");
}

function extractRev(filename: string): string {
  const m = filename.match(/_Rev([A-Z])/i);
  return m ? m[1].toUpperCase() : "A";
}

function extractDco(filename: string, content?: string): string | null {
  // Prefer content scan; fall back to filename conventions
  if (content) {
    const m = content.match(/DCO[:\s-]+(DCO-\d{4})/i);
    if (m) return m[1];
  }
  return null;
}

function spVersionToString(major: number, minor: number): string {
  if (minor === 0) return `${major}.0`;
  return `${major}.${minor}`;
}

function versionType(minor: number): "major" | "minor" {
  return minor === 0 ? "major" : "minor";
}

function docType(docId: string): string {
  if (docId.startsWith("QM-")) return "QM";
  if (docId.startsWith("SOP-")) return "SOP";
  if (docId.startsWith("FM-")) return "FM";
  if (docId.startsWith("FPS-")) return "FPS";
  return "DOC";
}

function docGroup(docId: string): string {
  if (docId.startsWith("QM-")) return "Quality Manual";
  if (docId.startsWith("SOP-QMS-")) return "SOPs — QMS Core";
  if (docId.startsWith("SOP-SUP-")) return "SOPs — Supplier & Receiving";
  if (docId.startsWith("SOP-FS-")) return "SOPs — Food Safety & Sanitation";
  if (docId.startsWith("SOP-PC-")) return "SOPs — Pest Control";
  if (docId.startsWith("SOP-PRD-") || docId.startsWith("SOP-FRS-") || docId.startsWith("SOP-RCL-")) return "SOPs — Production";
  if (docId.startsWith("FPS-")) return "Finished Product Specifications";
  if (["FM-001","FM-002","FM-003"].includes(docId)) return "Forms — Change Control";
  if (["FM-004","FM-005","FM-006","FM-007"].includes(docId)) return "Forms — Supplier & Receiving";
  if (["FM-027","FM-030"].includes(docId)) return "Forms — Quality Unit";
  if (docId.startsWith("FM-ALG")) return "Forms — Allergen";
  return "Other";
}

function warnCheck(docId: string, filename: string, dco: string | null): string | undefined {
  if (!dco && (docId.startsWith("SOP-") || docId.startsWith("QM-") || docId.startsWith("FM-"))) {
    return "No DCO assigned";
  }
  return undefined;
}

// ---------------------------------------------------------------------------
// Graph fetch helpers
// ---------------------------------------------------------------------------

async function fetchSpFiles(
  context: WebPartContext,
  folderPath: string
): Promise<SpFileItem[]> {
  const encoded = encodeURIComponent(folderPath);
  const url =
    `${SITE_URL}/_api/web/GetFolderByServerRelativeUrl('${encoded}')/Files` +
    `?$select=Name,ServerRelativeUrl,TimeLastModified,CheckOutType,CheckedOutByUser/Title,MajorVersion,MinorVersion,UIVersion` +
    `&$expand=CheckedOutByUser`;

  try {
    const resp: SPHttpClientResponse = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    return (data.value || []) as SpFileItem[];
  } catch {
    return [];
  }
}

async function fetchVersionHistory(
  context: WebPartContext,
  serverRelativeUrl: string
): Promise<VersionEntry[]> {
  const encoded = encodeURIComponent(serverRelativeUrl);
  const url =
    `${SITE_URL}/_api/web/GetFileByServerRelativeUrl('${encoded}')/Versions` +
    `?$select=VersionLabel,Created,CreatedBy/Title,CheckInComment` +
    `&$expand=CreatedBy`;

  try {
    const resp: SPHttpClientResponse = await context.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );
    if (!resp.ok) return [];
    const data = await resp.json();
    return ((data.value || []) as any[]).map((v: any) => {
      const label: string = v.VersionLabel || "1.0";
      const parts = label.split(".");
      const minor = parseInt(parts[1] || "0", 10);
      return {
        version: label,
        type: versionType(minor),
        who: v.CreatedBy?.Title || "Unknown",
        when: v.Created ? v.Created.substring(0, 10) : "",
        comment: v.CheckInComment || "",
      } as VersionEntry;
    }).reverse(); // newest first
  } catch {
    return [];
  }
}

async function fetchDcoStatuses(context: WebPartContext): Promise<Record<string, string>> {
  const url =
    `${SITE_URL}/_api/web/lists/getbytitle('QMS_DCOs')/items` +
    `?$select=Title,DCOStatus&$top=100`;
  try {
    const resp = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!resp.ok) return DCO_STATUS;
    const data = await resp.json();
    const map: Record<string, string> = {};
    for (const item of (data.value || [])) {
      map[item.Title] = (item.DCOStatus || "open").toLowerCase();
    }
    return Object.keys(map).length > 0 ? map : DCO_STATUS;
  } catch {
    return DCO_STATUS;
  }
}

// ---------------------------------------------------------------------------
// Main data loader
// ---------------------------------------------------------------------------

async function loadDocumentMatrix(context: WebPartContext): Promise<DocRecord[]> {
  const dcoStatusMap = await fetchDcoStatuses(context);

  // Collect all files per zone
  const zoneFiles: Record<string, SpFileItem[]> = {
    drafts: [],
    published: [],
    official: [],
  };

  for (const zone of ["drafts", "published", "official"] as const) {
    for (const folderPath of ZONE_PATHS[zone]) {
      const files = await fetchSpFiles(context, folderPath);
      for (const f of files) {
        // Skip non-docx and archive subfolder files
        if (!f.Name.toLowerCase().endsWith(".docx")) continue;
        if (f.ServerRelativeUrl.toLowerCase().includes("/archive/")) continue;
        zoneFiles[zone].push(f);
      }
    }
  }

  // Build a map: docId → DocRecord
  const docMap = new Map<string, DocRecord>();

  const processFile = async (
    f: SpFileItem,
    zone: "drafts" | "published" | "official"
  ): Promise<void> => {
    const docId = extractDocId(f.Name);
    const rev = extractRev(f.Name);
    const major = f.MajorVersion ?? 1;
    const minor = f.MinorVersion ?? 0;
    const ver = spVersionToString(major, minor);
    const vType = versionType(minor);

    // Parse path for folder label
    const pathParts = f.ServerRelativeUrl.split("/");
    const folderLabel = pathParts.slice(4, -1).join("/");

    // Version history (skip for official — usually empty)
    const history = zone !== "official"
      ? await fetchVersionHistory(context, f.ServerRelativeUrl)
      : [];

    // DCO — try to read from history comments or fall back to known map
    let dco: string | null = null;
    for (const h of history) {
      const m = h.comment.match(/DCO-\d{4}/i);
      if (m) { dco = m[0].toUpperCase(); break; }
    }
    // Also check filename
    const fnDco = f.Name.match(/DCO-(\d{4})/i);
    if (!dco && fnDco) dco = `DCO-${fnDco[1]}`;

    // Infer DCO from known mapping if still null
    if (!dco) {
      if (["QM-001","SOP-QMS-001","SOP-QMS-002","SOP-QMS-003","SOP-PRD-108",
           "SOP-PRD-432","SOP-FRS-549","FM-001","FM-002","FM-003","FM-027","FM-030"]
          .includes(docId)) dco = "DCO-0001";
      else if (docId.startsWith("SOP-SUP-") || docId.startsWith("SOP-FS-") ||
               docId.startsWith("SOP-PC-") || docId.startsWith("FM-0") ||
               docId === "FM-ALG") dco = "DCO-0002";
    }

    const dcoStatus = dco ? (dcoStatusMap[dco] || "open") : "n/a";
    const coUser = f.CheckOutType > 0 ? (f.CheckedOutByUser?.Title || "Unknown") : null;
    const warn = warnCheck(docId, f.Name, dco);

    const entry: ZoneEntry = {
      rev,
      version: ver,
      versionType: vType,
      status: dcoStatus === "signed" ? "implemented" : dcoStatus === "blocked" ? "inreview" : "draft",
      dco,
      dcoStatus,
      checkedOutTo: coUser,
      file: f.Name,
      path: folderLabel,
      modified: f.TimeLastModified.substring(0, 10),
      webUrl: `${SITE_URL}/${f.ServerRelativeUrl}`,
      warn,
      history,
    };

    if (!docMap.has(docId)) {
      docMap.set(docId, {
        id: docId,
        title: docId, // will refine below
        type: docType(docId),
        group: docGroup(docId),
        zones: { drafts: null, published: null, official: null },
      });
    }
    const rec = docMap.get(docId)!;
    rec.zones[zone] = entry;
  };

  // Process all zones (parallel per zone, sequential across zones to avoid rate limits)
  for (const zone of ["drafts", "published", "official"] as const) {
    await Promise.all(zoneFiles[zone].map(f => processFile(f, zone)));
  }

  // Enrich titles from known map
  const TITLES: Record<string, string> = {
    "QM-001": "Quality Manual",
    "SOP-QMS-001": "Management Responsibility",
    "SOP-QMS-002": "Document Control",
    "SOP-QMS-003": "Change Control",
    "SOP-SUP-001": "Supplier Qualification",
    "SOP-SUP-002": "Receiving Inspection",
    "SOP-FS-001": "Allergen Control",
    "SOP-FS-002": "Equipment Cleaning",
    "SOP-FS-003": "Facility Sanitation",
    "SOP-FS-004": "Environmental Monitoring",
    "SOP-PC-001": "Pest Sighting Response",
    "SOP-PRD-108": "Finished Product Release",
    "SOP-PRD-432": "Finished Product Spec & Testing",
    "SOP-FRS-549": "Product Specification Sheet",
    "SOP-RCL-321": "Recall Procedure",
    "FPS-001": "Lychee VD3 Gummy Spec",
    "FM-001": "Master Document Log",
    "FM-002": "Change Request Form",
    "FM-003": "Document Change Order Form",
    "FM-004": "Approved Supplier List",
    "FM-005": "Receiving Log",
    "FM-006": "Raw Material Spec Sheet",
    "FM-007": "Material Hold Label",
    "FM-027": "QU/QS Designation Record",
    "FM-030": "Finished Product Spec Sheet",
    "FM-ALG": "Allergen Status Record",
  };

  for (const [id, rec] of docMap) {
    rec.title = TITLES[id] || id;
  }

  // Sort: by group then by id
  const GROUP_ORDER = [
    "Quality Manual",
    "SOPs — QMS Core",
    "SOPs — Supplier & Receiving",
    "SOPs — Food Safety & Sanitation",
    "SOPs — Pest Control",
    "SOPs — Production",
    "Finished Product Specifications",
    "Forms — Change Control",
    "Forms — Supplier & Receiving",
    "Forms — Quality Unit",
    "Forms — Allergen",
    "Other",
  ];

  return [...docMap.values()].sort((a, b) => {
    const ga = GROUP_ORDER.indexOf(a.group);
    const gb = GROUP_ORDER.indexOf(b.group);
    if (ga !== gb) return ga - gb;
    return a.id.localeCompare(b.id);
  });
}

// ---------------------------------------------------------------------------
// HTML renderer — returns full innerHTML for the tab container
// ---------------------------------------------------------------------------

export function renderDocumentRepositoryTab(
  container: HTMLElement,
  context: WebPartContext
): void {
  container.innerHTML = getShell();
  attachStyles(container);
  showLoading(container);

  loadDocumentMatrix(context)
    .then(docs => {
      hideLoading(container);
      mountMatrix(container, docs);
    })
    .catch(err => {
      hideLoading(container);
      showError(container, err);
    });
}

// ---------------------------------------------------------------------------
// Shell HTML
// ---------------------------------------------------------------------------

function getShell(): string {
  return `
<div class="drm-wrap">
  <div class="drm-toolbar">
    <select id="drm-tf"><option value="all">All types</option><option value="SOP">SOP</option><option value="FM">FM</option><option value="QM">QM</option><option value="FPS">FPS</option></select>
    <select id="drm-df"><option value="all">All DCOs</option><option value="DCO-0001">DCO-0001</option><option value="DCO-0002">DCO-0002</option><option value="none">No DCO</option></select>
    <select id="drm-zf"><option value="all">All zones</option><option value="drafts">Has Drafts</option><option value="published">Has Published</option><option value="drafts-only">Drafts only</option><option value="both">Both zones</option></select>
    <input id="drm-qf" type="text" placeholder="Search ID or title…">
    <span class="drm-live-badge" id="drm-badge">Loading…</span>
  </div>
  <div class="drm-sbar" id="drm-sbar"></div>
  <div class="drm-flow">
    <div class="drm-flow-blank"></div>
    <div class="drm-flow-zone">Drafts<br><span>QMS/…/Drafts</span></div>
    <div class="drm-flow-arr"><div class="drm-arr-line"></div><div class="drm-arr-lbl">major check-in → publish</div></div>
    <div class="drm-flow-zone">Published<br><span>Published/QMS/…</span></div>
    <div class="drm-flow-arr"><div class="drm-arr-line"></div><div class="drm-arr-lbl">DCO sign-off → promote</div></div>
    <div class="drm-flow-zone">Official<br><span>Official/QMS/…</span></div>
  </div>
  <div class="drm-loading" id="drm-loading">
    <div class="drm-spinner"></div>
    <span>Loading document matrix from SharePoint…</span>
  </div>
  <div class="drm-error" id="drm-error" style="display:none"></div>
  <div id="drm-table-wrap" style="display:none">
    <table class="drm-table" id="drm-table">
      <colgroup>
        <col class="drm-col-id">
        <col class="drm-col-title">
        <col class="drm-col-zone">
        <col class="drm-col-zone">
        <col class="drm-col-zone">
      </colgroup>
      <thead>
        <tr>
          <th>Doc ID</th>
          <th>Title</th>
          <th class="drm-zh drm-zh-d">Drafts</th>
          <th class="drm-zh drm-zh-p">Published</th>
          <th class="drm-zh drm-zh-o">Official</th>
        </tr>
      </thead>
      <tbody id="drm-tbody"></tbody>
    </table>
  </div>
  <div class="drm-detail" id="drm-detail" style="display:none"></div>
</div>`;
}

// ---------------------------------------------------------------------------
// CSS injection
// ---------------------------------------------------------------------------

function attachStyles(container: HTMLElement): void {
  if (document.getElementById("drm-styles")) return;
  const style = document.createElement("style");
  style.id = "drm-styles";
  style.textContent = `
.drm-wrap{font-family:'Segoe UI',sans-serif;font-size:13px;color:#323130;padding:12px 0}
.drm-toolbar{display:flex;align-items:center;gap:8px;margin-bottom:10px;flex-wrap:wrap}
.drm-toolbar select,.drm-toolbar input{font-size:12px;padding:3px 8px;border:1px solid #C8C6C4;border-radius:2px;height:28px;background:#fff;color:#323130}
.drm-toolbar input{min-width:160px}
.drm-live-badge{font-size:10px;padding:2px 8px;border-radius:10px;background:#DFF6DD;color:#107C10;border:1px solid #A9D3A8;font-weight:600}
.drm-live-badge.loading{background:#F3F2F1;color:#605E5C;border-color:#C8C6C4}
.drm-sbar{display:flex;gap:6px;margin-bottom:10px;flex-wrap:wrap}
.drm-pill{font-size:11px;padding:2px 9px;border-radius:10px;border:1px solid #C8C6C4;background:#F3F2F1;color:#605E5C}
.drm-pill strong{font-weight:600;color:#323130}
.drm-pill.warn{border-color:#D83B01;color:#A4262C;background:#FDE7E9}
.drm-flow{display:grid;grid-template-columns:100px 144px 1fr 48px 1fr 48px 1fr;align-items:center;margin-bottom:4px;gap:0}
.drm-flow-blank{grid-column:span 2}
.drm-flow-zone{text-align:center;font-size:10px;color:#605E5C;line-height:1.4}
.drm-flow-zone span{font-size:9px;opacity:.7}
.drm-flow-arr{display:flex;flex-direction:column;align-items:center;gap:2px}
.drm-arr-line{width:100%;height:1px;background:#C8C6C4;position:relative}
.drm-arr-line::after{content:'';position:absolute;right:-1px;top:-3px;border-left:4px solid #C8C6C4;border-top:3px solid transparent;border-bottom:3px solid transparent}
.drm-arr-lbl{font-size:9px;color:#A19F9D;text-align:center;line-height:1.2}
.drm-loading{display:flex;align-items:center;gap:10px;padding:24px 0;color:#605E5C;font-size:13px}
.drm-spinner{width:20px;height:20px;border:2px solid #C8C6C4;border-top-color:#0078D4;border-radius:50%;animation:drm-spin .8s linear infinite;flex-shrink:0}
@keyframes drm-spin{to{transform:rotate(360deg)}}
.drm-error{padding:12px;background:#FDE7E9;color:#A4262C;border-radius:2px;font-size:12px}
.drm-table{width:100%;border-collapse:collapse;table-layout:fixed}
.drm-col-id{width:104px}
.drm-col-title{width:148px}
.drm-col-zone{width:calc((100% - 252px)/3)}
.drm-table thead th{font-size:11px;font-weight:600;padding:6px 9px;text-align:left;border-bottom:1px solid #C8C6C4;background:#F3F2F1;color:#323130;position:sticky;top:0;z-index:1}
.drm-zh{text-align:center!important;font-size:12px;padding:7px 9px!important;border-bottom:2px solid!important}
.drm-zh-d{color:#0078D4;border-color:#0078D4!important}
.drm-zh-p{color:#8A4B08;border-color:#C19C00!important}
.drm-zh-o{color:#107C10;border-color:#107C10!important}
.drm-table tbody tr{border-bottom:1px solid #F3F2F1;cursor:pointer}
.drm-table tbody tr:hover td{background:#F3F2F1}
.drm-table tbody tr.drm-sel td{background:#DEECF9}
.drm-table tbody tr td{padding:0;vertical-align:top}
.drm-grp td{padding:4px 9px 3px!important;font-size:10px;font-weight:600;color:#605E5C;background:#FAF9F8;border-top:1px solid #E1DFDD;cursor:default!important;letter-spacing:.3px;text-transform:uppercase}
.drm-grp:hover td{background:#FAF9F8!important}
.drm-id-cell{padding:7px 9px}
.drm-doc-id{font-size:11px;font-weight:600;font-family:'Cascadia Code','Courier New',monospace;color:#323130;white-space:nowrap}
.drm-type-badge{font-size:9px;padding:1px 4px;border-radius:2px;margin-top:2px;display:inline-block;font-weight:600}
.drm-type-SOP{background:#DEECF9;color:#0078D4}
.drm-type-FM{background:#FFF4CE;color:#7A4F01}
.drm-type-QM{background:#DFF6DD;color:#107C10}
.drm-type-FPS{background:#F0E6FF;color:#5C2D91}
.drm-title-cell{padding:7px 9px;font-size:11px;color:#605E5C;line-height:1.4;vertical-align:middle!important}
.drm-zone-cell{padding:6px 8px;border-left:1px solid #F3F2F1;vertical-align:top}
.drm-zone-cell.drm-empty{background:repeating-linear-gradient(45deg,transparent,transparent 5px,#F3F2F1 5px,#F3F2F1 5.5px);opacity:.5}
.drm-vc{display:flex;flex-direction:column;gap:3px}
.drm-vb{display:flex;align-items:center;gap:4px;flex-wrap:wrap}
.drm-rev{font-size:10px;font-family:'Cascadia Code','Courier New',monospace;font-weight:600;padding:1px 5px;border-radius:2px}
.drm-rev-A{background:#DEECF9;color:#0078D4}
.drm-rev-B{background:#DFF6DD;color:#107C10}
.drm-rev-C{background:#FFF4CE;color:#7A4F01}
.drm-ver{font-size:10px;font-family:'Cascadia Code','Courier New',monospace;color:#605E5C;padding:1px 4px;border-radius:2px;background:#F3F2F1;display:inline-flex;align-items:center;gap:3px}
.drm-ver.major{background:#DFF6DD;color:#107C10;font-weight:600}
.drm-ver.minor{background:#F3F2F1;color:#605E5C}
.drm-dco{font-size:9px;font-family:'Cascadia Code','Courier New',monospace;padding:1px 4px;border-radius:2px;white-space:nowrap}
.drm-dco-blocked{background:#FDE7E9;color:#A4262C}
.drm-dco-open{background:#FFF4CE;color:#7A4F01}
.drm-dco-signed{background:#DFF6DD;color:#107C10}
.drm-dco-na{background:#F3F2F1;color:#A19F9D}
.drm-co{font-size:9px;color:#8A4B08;display:flex;align-items:center;gap:3px}
.drm-co-dot{width:6px;height:6px;border-radius:50%;background:#C19C00;flex-shrink:0}
.drm-co-in{color:#107C10}
.drm-co-dot-in{background:#107C10}
.drm-mod{font-size:9px;color:#A19F9D}
.drm-path{font-size:9px;color:#A19F9D;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:120px}
.drm-warn{font-size:9px;color:#A4262C}
.drm-detail{margin-top:10px;border:1px solid #C8C6C4;border-radius:2px;background:#fff}
.drm-detail-head{padding:8px 14px;background:#F3F2F1;border-bottom:1px solid #C8C6C4;display:flex;align-items:center;gap:10px}
.drm-detail-id{font-family:'Cascadia Code','Courier New',monospace;font-size:13px;font-weight:600}
.drm-detail-title{font-size:12px;color:#605E5C;flex:1}
.drm-detail-close{cursor:pointer;color:#605E5C;font-size:16px;padding:0 4px;border:none;background:transparent}
.drm-detail-body{padding:10px 14px;display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px}
.drm-dz{border:1px solid #E1DFDD;border-radius:2px;padding:10px}
.drm-dz-name{font-size:11px;font-weight:600;margin-bottom:7px;padding-bottom:5px;border-bottom:1px solid #E1DFDD}
.drm-dz-name.drafts{color:#0078D4}
.drm-dz-name.published{color:#8A4B08}
.drm-dz-name.official{color:#107C10}
.drm-df-row{font-size:11px;margin-bottom:4px;display:flex;gap:5px;align-items:flex-start}
.drm-df-lbl{color:#A19F9D;min-width:58px;flex-shrink:0;font-size:10px;padding-top:1px}
.drm-df-val{color:#323130;word-break:break-all;font-size:10px;font-family:'Cascadia Code','Courier New',monospace}
.drm-absent{font-size:11px;color:#A19F9D;font-style:italic}
.drm-vh{margin-top:8px;border-top:1px solid #E1DFDD;padding-top:8px}
.drm-vh-title{font-size:10px;font-weight:600;color:#605E5C;margin-bottom:5px}
.drm-vh-row{display:flex;align-items:flex-start;gap:5px;padding:3px 0;border-bottom:1px solid #F3F2F1;font-size:10px}
.drm-vh-row:last-child{border:none}
.drm-vh-ver{font-family:'Cascadia Code','Courier New',monospace;min-width:28px;font-weight:600}
.drm-vh-type{padding:1px 4px;border-radius:2px;font-size:9px;flex-shrink:0}
.drm-vh-type.major{background:#DFF6DD;color:#107C10;font-weight:600}
.drm-vh-type.minor{background:#F3F2F1;color:#605E5C}
.drm-vh-who{color:#605E5C;flex:1}
.drm-vh-when{color:#A19F9D;flex-shrink:0}
.drm-open-btn{display:inline-flex;align-items:center;gap:4px;font-size:10px;color:#0078D4;text-decoration:none;margin-top:4px;padding:2px 6px;border:1px solid #C8C6C4;border-radius:2px;background:#fff}
.drm-open-btn:hover{background:#DEECF9}
`;
  document.head.appendChild(style);
}

// ---------------------------------------------------------------------------
// Mount the interactive matrix
// ---------------------------------------------------------------------------

function mountMatrix(container: HTMLElement, allDocs: DocRecord[]): void {
  let filtered = [...allDocs];
  let selected: string | null = null;

  const badge = container.querySelector("#drm-badge") as HTMLElement;
  badge.textContent = `Live · SP · ${new Date().toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}`;
  badge.classList.remove("loading");

  const tableWrap = container.querySelector("#drm-table-wrap") as HTMLElement;
  tableWrap.style.display = "";

  function go(): void {
    const t = (container.querySelector("#drm-tf") as HTMLSelectElement).value;
    const d = (container.querySelector("#drm-df") as HTMLSelectElement).value;
    const z = (container.querySelector("#drm-zf") as HTMLSelectElement).value;
    const q = (container.querySelector("#drm-qf") as HTMLInputElement).value.toLowerCase();

    filtered = allDocs.filter(doc => {
      const allZ = [doc.zones.drafts, doc.zones.published, doc.zones.official].filter(Boolean) as ZoneEntry[];
      const mt = t === "all" || doc.type === t;
      const mq = !q || doc.id.toLowerCase().includes(q) || doc.title.toLowerCase().includes(q);
      const md = d === "all"
        || (d === "none" && allZ.every(z => !z.dco))
        || (d !== "none" && allZ.some(z => z.dco === d));
      const hasd = !!doc.zones.drafts, hasp = !!doc.zones.published;
      const mz = z === "all"
        || (z === "drafts" && hasd)
        || (z === "published" && hasp)
        || (z === "drafts-only" && hasd && !hasp)
        || (z === "both" && hasd && hasp);
      return mt && mq && md && mz;
    });

    renderSummary();
    renderTable();
  }

  function renderSummary(): void {
    const tot = filtered.length;
    const hasp = filtered.filter(d => d.zones.published).length;
    const both = filtered.filter(d => d.zones.drafts && d.zones.published).length;
    const donly = filtered.filter(d => d.zones.drafts && !d.zones.published).length;
    const ponly = filtered.filter(d => !d.zones.drafts && d.zones.published).length;
    const bl = filtered.filter(d =>
      [d.zones.drafts, d.zones.published, d.zones.official].filter(Boolean)
        .some((z: any) => z.dcoStatus === "blocked")).length;
    const co = filtered.filter(d =>
      [d.zones.drafts, d.zones.published, d.zones.official].filter(Boolean)
        .some((z: any) => z.checkedOutTo)).length;
    const warns = filtered.filter(d =>
      [d.zones.drafts, d.zones.published, d.zones.official].filter(Boolean)
        .some((z: any) => z.warn)).length;

    const sbar = container.querySelector("#drm-sbar")!;
    sbar.innerHTML =
      `<span class="drm-pill"><strong>${tot}</strong> documents</span>` +
      `<span class="drm-pill"><strong>${both}</strong> in both zones</span>` +
      `<span class="drm-pill"><strong>${donly}</strong> drafts only</span>` +
      `<span class="drm-pill"><strong>${ponly}</strong> published only</span>` +
      `<span class="drm-pill"><strong>0</strong> in Official</span>` +
      (bl ? `<span class="drm-pill warn"><strong>${bl}</strong> DCO blocked</span>` : "") +
      (co ? `<span class="drm-pill warn"><strong>${co}</strong> checked out</span>` : "") +
      (warns ? `<span class="drm-pill warn"><strong>${warns}</strong> warnings</span>` : "");
  }

  function dcoClass(ds: string): string {
    const map: Record<string, string> = { blocked: "drm-dco-blocked", open: "drm-dco-open", signed: "drm-dco-signed", "n/a": "drm-dco-na" };
    return map[ds] || "drm-dco-na";
  }

  function revClass(rev: string): string {
    return `drm-rev drm-rev-${rev}`;
  }

  function zoneCell(z: ZoneEntry | null): string {
    if (!z) return `<td class="drm-zone-cell drm-empty"></td>`;
    const verClass = z.versionType === "major" ? "major" : "minor";
    const verIcon = z.versionType === "major" ? "●" : "○";
    const coHtml = z.checkedOutTo
      ? `<div class="drm-co"><span class="drm-co-dot"></span>${z.checkedOutTo}</div>`
      : `<div class="drm-co drm-co-in"><span class="drm-co-dot drm-co-dot-in"></span>Checked in</div>`;
    return `<td class="drm-zone-cell"><div class="drm-vc">
      <div class="drm-vb">
        <span class="${revClass(z.rev)}">Rev ${z.rev}</span>
        <span class="drm-ver ${verClass}">${verIcon} ${z.version}</span>
        ${z.dco ? `<span class="drm-dco ${dcoClass(z.dcoStatus)}">${z.dco}</span>` : `<span class="drm-dco drm-dco-na">—</span>`}
      </div>
      ${coHtml}
      <span class="drm-mod">${z.modified}</span>
      <span class="drm-path" title="${z.path}">${z.path}</span>
      ${z.warn ? `<span class="drm-warn">${z.warn}</span>` : ""}
    </div></td>`;
  }

  function renderTable(): void {
    const tbody = container.querySelector("#drm-tbody")!;
    const groups = [...new Set(filtered.map(d => d.group))];
    let html = "";
    groups.forEach(g => {
      const docs = filtered.filter(d => d.group === g);
      html += `<tr class="drm-grp"><td colspan="5">${g}</td></tr>`;
      docs.forEach(doc => {
        html += `<tr class="${selected === doc.id ? "drm-sel" : ""}" data-id="${doc.id}">
          <td class="drm-id-cell">
            <div class="drm-doc-id">${doc.id}</div>
            <span class="drm-type-badge drm-type-${doc.type}">${doc.type}</span>
          </td>
          <td class="drm-title-cell">${doc.title}</td>
          ${zoneCell(doc.zones.drafts)}
          ${zoneCell(doc.zones.published)}
          ${zoneCell(doc.zones.official)}
        </tr>`;
      });
    });
    tbody.innerHTML = html;

    // Bind row clicks
    tbody.querySelectorAll("tr[data-id]").forEach(row => {
      row.addEventListener("click", () => {
        const id = (row as HTMLElement).dataset.id!;
        selected = selected === id ? null : id;
        renderTable();
        selected ? showDetail(id) : hideDetail();
      });
    });
  }

  function showDetail(id: string): void {
    const doc = allDocs.find(d => d.id === id)!;
    const panel = container.querySelector("#drm-detail") as HTMLElement;
    panel.style.display = "";

    const zkeys = ["drafts", "published", "official"] as const;
    const znames = { drafts: "Drafts", published: "Published", official: "Official" };

    const zonesHtml = zkeys.map(zk => {
      const z = doc.zones[zk];
      if (!z) return `<div class="drm-dz"><div class="drm-dz-name ${zk}">${znames[zk]}</div><div class="drm-absent">Not present in this zone</div></div>`;

      const histHtml = z.history.length > 0
        ? `<div class="drm-vh">
            <div class="drm-vh-title">Version history</div>
            ${z.history.slice(0, 8).map(h => `
              <div class="drm-vh-row">
                <span class="drm-vh-ver">${h.version}</span>
                <span class="drm-vh-type ${h.type}">${h.type}</span>
                <span class="drm-vh-who">${h.who}</span>
                <span class="drm-vh-when">${h.when}</span>
              </div>`).join("")}
           </div>`
        : "";

      return `<div class="drm-dz">
        <div class="drm-dz-name ${zk}">${znames[zk]}</div>
        <div class="drm-df-row"><span class="drm-df-lbl">Revision</span><span class="drm-rev drm-rev-${z.rev}">Rev ${z.rev}</span></div>
        <div class="drm-df-row"><span class="drm-df-lbl">Version</span><span class="drm-ver ${z.versionType}">${z.versionType === "major" ? "●" : "○"} ${z.version}</span></div>
        <div class="drm-df-row"><span class="drm-df-lbl">DCO</span><span class="drm-dco ${dcoClass(z.dcoStatus)}">${z.dco || "none"} · ${z.dcoStatus}</span></div>
        <div class="drm-df-row"><span class="drm-df-lbl">Check-out</span><span class="drm-df-val" style="font-family:inherit">${z.checkedOutTo ? `⚠ ${z.checkedOutTo}` : "✓ Checked in"}</span></div>
        <div class="drm-df-row"><span class="drm-df-lbl">Modified</span><span class="drm-df-val">${z.modified}</span></div>
        <div class="drm-df-row" style="align-items:flex-start"><span class="drm-df-lbl">File</span><span class="drm-df-val">${z.file}</span></div>
        ${z.warn ? `<div class="drm-df-row"><span class="drm-df-lbl" style="color:#A4262C">Warning</span><span class="drm-df-val" style="color:#A4262C;font-family:inherit">${z.warn}</span></div>` : ""}
        <a class="drm-open-btn" href="${z.webUrl}" target="_blank">Open in SharePoint ↗</a>
        ${histHtml}
      </div>`;
    }).join("");

    panel.innerHTML = `
      <div class="drm-detail-head">
        <span class="drm-detail-id">${doc.id}</span>
        <span class="drm-detail-title">${doc.title}</span>
        <button class="drm-detail-close" id="drm-close-btn">✕</button>
      </div>
      <div class="drm-detail-body">${zonesHtml}</div>`;

    panel.querySelector("#drm-close-btn")!.addEventListener("click", () => {
      selected = null;
      hideDetail();
      renderTable();
    });

    panel.scrollIntoView({ behavior: "smooth", block: "nearest" });
  }

  function hideDetail(): void {
    const panel = container.querySelector("#drm-detail") as HTMLElement;
    panel.style.display = "none";
    panel.innerHTML = "";
  }

  // Wire toolbar
  ["#drm-tf", "#drm-df", "#drm-zf"].forEach(sel => {
    container.querySelector(sel)?.addEventListener("change", go);
  });
  container.querySelector("#drm-qf")?.addEventListener("input", go);

  go();
}

// ---------------------------------------------------------------------------
// Loading / error helpers
// ---------------------------------------------------------------------------

function showLoading(container: HTMLElement): void {
  const el = container.querySelector("#drm-loading") as HTMLElement;
  if (el) el.style.display = "flex";
}

function hideLoading(container: HTMLElement): void {
  const el = container.querySelector("#drm-loading") as HTMLElement;
  if (el) el.style.display = "none";
}

function showError(container: HTMLElement, err: unknown): void {
  const el = container.querySelector("#drm-error") as HTMLElement;
  if (el) {
    el.style.display = "";
    el.textContent = `Error loading document matrix: ${err instanceof Error ? err.message : String(err)}`;
  }
}
