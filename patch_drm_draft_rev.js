/**
 * patch_drm_draft_rev.js
 * IMP9177 — QMS Portal  |  Doc Repo: show DRAFT instead of Rev A for draft-zone files
 *
 * Run from SPFx root:
 *   node .\patch_drm_draft_rev.js
 *
 * Anchors verified against real QmsPortalWebPart.ts (Select-String output 2026-04-25).
 *
 * Patches:
 *   P-01  _drmRev() — return 'DRAFT' when filename contains _DRAFT_, else parse _RevX, else 'A'
 *   P-02  zc() cell renderer — replace hardcoded `Rev ${z.rev}` with DRAFT-aware label
 *   P-03  showD() detail panel — replace hardcoded `Rev ${z.rev}` with DRAFT-aware label
 *
 * STRICT MODE: anchor must exist before replacement; exits with error on any failure.
 * File is never written if any patch fails.
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

function applyPatch(src, patchId, anchor, search, replacement) {
  const anchorFound = (anchor instanceof RegExp) ? anchor.test(src) : src.includes(anchor);
  if (!anchorFound) {
    console.error('\n[FAIL] ' + patchId + ' — anchor not found.');
    console.error('  Anchor: ' + anchor);
    console.error('  Source may already be patched or anchor has changed.\n');
    process.exit(1);
  }
  const result = (search instanceof RegExp)
    ? src.replace(search, replacement)
    : src.split(search).join(replacement);
  if (result === src) {
    console.error('\n[FAIL] ' + patchId + ' — replacement produced no change (already patched?)\n');
    process.exit(1);
  }
  console.log('[OK]   ' + patchId);
  return result;
}

console.log('\n=== patch_drm_draft_rev.js ===');
console.log('Target: ' + TARGET + '\n');

let src = readFile(TARGET);

// ===========================================================================
// P-01 — Fix _drmRev() to return 'DRAFT' for _DRAFT_ filenames
//
// REAL source (line 2418):
//   private _drmRev(n:string):string{const m=n.match(/_Rev([A-Z])/i);return m?m[1].toUpperCase():'A';}
//
// AFTER:
//   Check for _DRAFT_ first → return 'DRAFT'
//   Then check for _RevX → return that letter
//   Else fall back to 'A'
// ===========================================================================
src = applyPatch(
  src,
  'P-01 — _drmRev: return DRAFT for _DRAFT_ filenames',
  "private _drmRev(n:string):string{const m=n.match(/_Rev([A-Z])/i);return m?m[1].toUpperCase():'A';}",
  "private _drmRev(n:string):string{const m=n.match(/_Rev([A-Z])/i);return m?m[1].toUpperCase():'A';}",
  "private _drmRev(n:string):string{if(/_DRAFT_/i.test(n))return'DRAFT';const m=n.match(/_Rev([A-Z])/i);return m?m[1].toUpperCase():'A';}"
);

// ===========================================================================
// P-02 — Fix zc() cell renderer: DRAFT-aware revision label
//
// REAL source (inside zc in _drmMount, line 2422):
//   `Rev ${z.rev}`
// appears as:
//   d:var(--b1);color:#1565c0">Rev ${z.rev}</span>
//
// There are TWO occurrences of `Rev ${z.rev}` in the file —
//   one in zc() (table cell)
//   one in showD() (detail panel)
// We patch them individually using their unique surrounding context.
//
// zc() anchor: the background:var(--b1);color:#1565c0 span that precedes it
// in the zone cell — unique to the compact table cell renderer.
// ===========================================================================
src = applyPatch(
  src,
  'P-02 — zc() cell: DRAFT-aware Rev label',
  'd:var(--b1);color:#1565c0">Rev ${z.rev}</span>',
  'd:var(--b1);color:#1565c0">Rev ${z.rev}</span>',
  "d:var(--b1);color:#1565c0\">${z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev}</span>"
);

// ===========================================================================
// P-03 — Fix showD() detail panel: DRAFT-aware revision label
//
// REAL source (inside showD in _drmMount, line 2422):
//   background:var(--b1);color:#1565c0">Rev ${z.rev}</span>
// The detail panel uses a slightly different inline style string —
// it has font-weight:700 and a min-width span before it.
// Unique anchor: the "Revision" label row context.
// ===========================================================================
src = applyPatch(
  src,
  'P-03 — showD() detail panel: DRAFT-aware Rev label',
  ';color:#1565c0">Rev\n           ${z.rev}</span>',
  // The detail panel rev span — unique context is the preceding `background:var(--b1);color:#1565c0` 
  // inside the zonesH map, which has font-weight:700 not present in zc()
  /background:var\(--b1\);color:#1565c0">Rev \$\{z\.rev\}<\/span><\/div><div style="font-size:11px;margin-bottom/,
  "background:var(--b1);color:#1565c0\">${z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev}</span></div><div style=\"font-size:11px;margin-bottom"
);

// ===========================================================================
// Write
// ===========================================================================
fs.writeFileSync(TARGET, src, 'utf8');

console.log('\n✅  All patches applied successfully.');
console.log('    File written: ' + TARGET);
console.log('\nNext steps:');
console.log('  1. .\\node_modules\\.bin\\heft clean');
console.log('  2. .\\node_modules\\.bin\\heft build --production');
console.log('  3. .\\node_modules\\.bin\\heft package-solution --production');
console.log('  4. Add-PnPApp -Path .\\sharepoint\\solution\\imp-9177-spfx.sppkg -Scope Site -Overwrite -Publish');
console.log('  5. Ctrl+Shift+R on QMS-Portal.aspx\n');
