/**
 * patch_drm_draft_rev_final.js
 * IMP9177 — QMS Portal  |  Doc Repo: DRAFT revision display — all three patches
 *
 * Run from SPFx root:
 *   node .\patch_drm_draft_rev_final.js
 *
 * Patches:
 *   P-01  _drmRev() — return 'DRAFT' for _DRAFT_ filenames
 *   P-02  zc() cell renderer — DRAFT-aware Rev label (may already be applied)
 *   P-03  showD() detail panel — DRAFT-aware Rev label
 *
 * All searches use regex with \s* to handle newlines and varying indentation.
 * Already-applied patches are detected and skipped gracefully.
 */

'use strict';

const fs   = require('fs');
const path = require('path');

const TARGET = path.resolve(
  __dirname,
  'src', 'webparts', 'qmsPortalWebPart', 'QmsPortalWebPart.ts'
);

function readFile(fp) {
  if (!fs.existsSync(fp)) {
    console.error('\n[FATAL] File not found:\n  ' + fp + '\n');
    process.exit(1);
  }
  return fs.readFileSync(fp, 'utf8');
}

function applyPatch(src, patchId, alreadyAppliedCheck, search, replacement) {
  // Skip if already applied
  if (alreadyAppliedCheck && src.includes(alreadyAppliedCheck)) {
    console.log('[SKIP] ' + patchId + ' — already applied');
    return src;
  }
  if (!search.test(src)) {
    console.error('\n[FAIL] ' + patchId + ' — search pattern not found in source.');
    console.error('  Pattern: ' + search);
    process.exit(1);
  }
  const result = src.replace(search, replacement);
  if (result === src) {
    console.error('\n[FAIL] ' + patchId + ' — replacement produced no change.\n');
    process.exit(1);
  }
  console.log('[OK]   ' + patchId);
  return result;
}

console.log('\n=== patch_drm_draft_rev_final.js ===');
console.log('Target: ' + TARGET + '\n');

let src = readFile(TARGET);

// ===========================================================================
// P-01 — _drmRev(): return 'DRAFT' for _DRAFT_ filenames
//
// REAL source (line 2418, with possible newlines):
//   private _drmRev(n:string):string{const m=n.match(/_Rev([A-Z])/i);return
//     m?m[1].toUpperCase():'A';}
// ===========================================================================
src = applyPatch(
  src,
  'P-01 — _drmRev: DRAFT support',
  "return'DRAFT'",   // already-applied sentinel
  /private _drmRev\(n:string\):string\{const m=n\.match\(\/_Rev\(\[A-Z\]\)\/i\);\s*return\s*m\?m\[1\]\.toUpperCase\(\):'A';\}/,
  "private _drmRev(n:string):string{if(/_DRAFT_/i.test(n))return'DRAFT';const m=n.match(/_Rev([A-Z])/i);return m?m[1].toUpperCase():'A';}"
);

// ===========================================================================
// P-02 — zc() cell renderer: DRAFT-aware Rev label
//
// Targets the span inside zc() with font-weight:700 and background:var(--b1)
// that reads: >Rev ${z.rev}</span>
// Uses unique prefix `eight:700;padding:1px 5px` to distinguish from other spans.
// ===========================================================================
src = applyPatch(
  src,
  'P-02 — zc() cell: DRAFT-aware Rev label',
  "z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev}</span><span s",  // already-applied sentinel
  /eight:700;padding:1px 5px;border-radius:2px;background:var\(--b1\);color:#1565c0">Rev\s*\$\{z\.rev\}<\/span>/g,
  "eight:700;padding:1px 5px;border-radius:2px;background:var(--b1);color:#1565c0\">${z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev}</span>"
);

// ===========================================================================
// P-03 — showD() detail panel: DRAFT-aware Rev label
//
// Targets the Revision row span inside zonesH map.
// Unique context: preceded by `Revision</span>` label and followed by `</div><div`
// with Version row. Uses \s* to handle newlines anywhere in the span.
// ===========================================================================
src = applyPatch(
  src,
  'P-03 — showD() detail panel: DRAFT-aware Rev label',
  "z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev}</span></div><div style=\"font-size:11px;margin-bottom:4px;display:flex;gap:5px\"><span style=\"color:var(--s5);min-width:58px;font-size:10px\">Version",  // already-applied sentinel
  /Revision<\/span><span style="font-size:10px;font-family:var\(--mono\);font-weight:700;padding:1px 5px;border-radius:2px;background:var\(--b1\);color:#1565c0">\s*Rev\s*\$\{z\.rev\}\s*<\/span><\/div>/,
  'Revision</span><span style="font-size:10px;font-family:var(--mono);font-weight:700;padding:1px 5px;border-radius:2px;background:var(--b1);color:#1565c0">${z.rev===\'DRAFT\'?\'DRAFT\':\'Rev \'+z.rev}</span></div>'
);

// ===========================================================================
// Final verification — no bare `Rev ${z.rev}` should remain
// ===========================================================================
const remaining = (src.match(/Rev \$\{z\.rev\}/g) || []).length;
if (remaining > 0) {
  console.warn('\n[WARN] ' + remaining + ' occurrence(s) of bare `Rev ${z.rev}` still remain.');
  console.warn('       These may be in other web parts — review manually.\n');
} else {
  console.log('[OK]   Verification: no bare `Rev ${z.rev}` remaining ✅');
}

fs.writeFileSync(TARGET, src, 'utf8');

console.log('\n✅  All patches applied successfully.');
console.log('    File written: ' + TARGET);
console.log('\nNext steps:');
console.log('  1. .\\node_modules\\.bin\\heft clean');
console.log('  2. .\\node_modules\\.bin\\heft build --production');
console.log('  3. .\\node_modules\\.bin\\heft package-solution --production');
console.log('  4. Add-PnPApp -Path .\\sharepoint\\solution\\imp-9177-spfx.sppkg -Scope Site -Overwrite -Publish');
console.log('  5. Ctrl+Shift+R on QMS-Portal.aspx\n');
