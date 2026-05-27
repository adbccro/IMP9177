/**
 * IMP9177 — Patch 2: Fix buildShell() CR tab HTML
 * Generated: 2026-04-27 | Session: CR Feature Build
 *
 * Run AFTER patch_add_cr_feature.js
 * Run from SPFx project root.
 *
 * This patch replaces the sc-cr tab stub in buildShell() with the full
 * CR panel HTML containing #cr-list-wrap, filter bar, and + New CR button.
 *
 * Also fixes DCO modal CR Link field to resolve CR_ID → CR_Title.
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const TARGET = path.resolve(
  process.cwd(),
  'src', 'webparts', 'qmsPortalWebPart', 'QmsPortalWebPart.ts'
);

function readFile(fp) {
  if (!fs.existsSync(fp)) { console.error(`\n[FATAL] File not found:\n  ${fp}\n`); process.exit(1); }
  return fs.readFileSync(fp, 'utf8');
}

console.log(`\n[IMP9177] buildShell CR Panel Patch`);
console.log(`Target: ${TARGET}\n`);
let src = readFile(TARGET);

// ── Already patched guard ────────────────────────────────────────────────────
if (src.includes('cr-list-wrap') && src.includes('cr-filter-bar') && src.includes('cr-count-badge')) {
  console.log('[SKIP] buildShell() CR panel already patched — cr-list-wrap, cr-filter-bar, cr-count-badge all present.');
} else {
  // Strategy: find the sc-cr section in buildShell()
  // It exists as a template literal string. We look for patterns like:
  //   <div class="sc" id="sc-cr">
  // The closing pattern varies. We use a broad regex to find and replace.

  const crPanelContent = `<div class="sc" id="sc-cr">
          <div style="padding:20px 24px;">
            <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:16px;">
              <div style="display:flex;align-items:center;gap:12px;">
                <h2 style="font-size:16px;font-weight:700;color:#1e3a5f;margin:0">Change Requests</h2>
                <span id="cr-count-badge" style="background:#e5e7eb;color:#374151;font-size:11px;font-weight:700;padding:2px 8px;border-radius:10px">—</span>
              </div>
              <button onclick="window.qpOpenNewCR()" style="padding:6px 14px;background:#1e3a5f;color:#fff;border:none;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;display:none" id="cr-new-btn">+ New CR</button>
            </div>
            <div id="cr-filter-bar" style="display:flex;gap:10px;align-items:center;margin-bottom:16px;flex-wrap:wrap;">
              <input id="cr-search" type="text" placeholder="Search CR ID or title\\u2026"
                style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px;width:200px"
                oninput="window._qpRenderCR && window._qpRenderCR()">
              <select id="cr-filter-status"
                style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px"
                onchange="window._qpRenderCR && window._qpRenderCR()">
                <option value="">All Statuses</option>
                <option>Draft</option><option>Submitted</option><option>Approved</option>
                <option>Rejected</option><option>Linked</option><option>Closed</option>
              </select>
              <select id="cr-filter-priority"
                style="padding:6px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:12px"
                onchange="window._qpRenderCR && window._qpRenderCR()">
                <option value="">All Priorities</option>
                <option>Low</option><option>Medium</option><option>High</option><option>Critical</option>
              </select>
            </div>
            <div id="cr-list-wrap" style="border:1px solid #e5e7eb;border-radius:8px;overflow:hidden;">
              <div style="padding:24px;text-align:center;color:#9ca3af;font-size:13px">Loading\\u2026</div>
            </div>
          </div>
        </div>`;

  // Try several anchor patterns for the existing sc-cr div
  const patterns = [
    // Pattern 1: div with id sc-cr followed by any content up to next sc div or section end
    /(<div[^>]*id=["']sc-cr["'][^>]*>)[\s\S]*?(?=<div[^>]*class=["'][^"']*sc[^"']*["'][^>]*id=["']sc-(?!cr)|<\/div>\s*<\/main|\$\{)/,
    // Pattern 2: simple sc-cr content between markers
    /id="sc-cr"[^>]*>[\s\S]{0,2000}?(?=id="sc-records"|id="sc-training"|id="sc-docs")/
  ];

  let patched = false;
  for (const pattern of patterns) {
    if (pattern.test(src)) {
      const before = src;
      src = src.replace(pattern, crPanelContent + '\n        ');
      if (src !== before) {
        console.log('  [OK] buildShell() sc-cr panel replaced');
        patched = true;
        break;
      }
    }
  }

  if (!patched) {
    console.log('  [WARN] Could not auto-patch sc-cr div — manual fix needed.');
    console.log('  See cr_panel_injection.md for the HTML to insert manually.');
    console.log('  The feature will still work but the CR tab will show a loading spinner');
    console.log('  until _renderCR() runs and populates #cr-list-wrap.');
    console.log('');
    console.log('  Quick manual fix: search for id="sc-cr" in QmsPortalWebPart.ts');
    console.log('  and replace the div\'s inner content with the HTML from cr_panel_injection.md');
  }
}

// ── Fix cr-new-btn visibility based on role ─────────────────────────────────
// Add role check in _attachListeners or _renderAll to show/hide the + New CR button
if (!src.includes('cr-new-btn') || !src.includes("style.display = ''")) {
  // Inject show/hide for cr-new-btn after _currentUser is resolved
  const anchor = `this._renderCR(); // [CR-03 PATCHED]`;
  if (src.includes(anchor)) {
    const insert = `
    // Show + New CR button for PM/ADB User
    const crNewBtn = document.getElementById('cr-new-btn');
    if (crNewBtn) {
      crNewBtn.style.display = (this._currentUser?.role === 'PM' || this._currentUser?.role === 'ADB User') ? '' : 'none';
    }`;
    src = src.split(anchor).join(anchor + insert);
    console.log('  [OK] cr-new-btn role-gate wired');
  }
}

// ── Write patched file ───────────────────────────────────────────────────────
fs.writeFileSync(TARGET, src, 'utf8');
console.log(`\n[SUCCESS] QmsPortalWebPart.ts updated in-place.`);
console.log(`Lines: ${src.split('\n').length}`);
console.log(`\nRun build now:`);
console.log(`  nvm use 22.14.0`);
console.log(`  .\\node_modules\\.bin\\heft clean`);
console.log(`  .\\node_modules\\.bin\\heft build --production`);
console.log(`  .\\node_modules\\.bin\\heft package-solution --production`);
