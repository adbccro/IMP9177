/**
 * patch_drm_draft_rev_p02p03.js
 * IMP9177 — QMS Portal  |  Doc Repo: DRAFT label fix — P-02 and P-03 only
 *
 * P-01 already applied in prior run. This script handles only the two
 * Rev ${z.rev} occurrences that have a literal newline inside the template literal.
 *
 * Run from SPFx root:
 *   node .\patch_drm_draft_rev_p02p03.js
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

function applyPatch(src, patchId, search, replacement) {
  if (!src.includes(search)) {
    console.error('\n[FAIL] ' + patchId + ' — search string not found in source.');
    console.error('  Search (repr): ' + JSON.stringify(search));
    process.exit(1);
  }
  const count = src.split(search).length - 1;
  if (count > 1) console.warn('[WARN] ' + patchId + ' — ' + count + ' occurrences; replacing all.');
  const result = src.split(search).join(replacement);
  if (result === src) {
    console.error('\n[FAIL] ' + patchId + ' — replacement produced no change.\n');
    process.exit(1);
  }
  console.log('[OK]   ' + patchId + ' (' + count + ' occurrence' + (count > 1 ? 's' : '') + ')');
  return result;
}

console.log('\n=== patch_drm_draft_rev_p02p03.js ===');
console.log('Target: ' + TARGET + '\n');

let src = readFile(TARGET);

// ===========================================================================
// P-02 — zc() cell renderer
//
// REAL source (inside zc template literal) — note the newline before ${z.rev}:
//   d:var(--b1);color:#1565c0">Rev ${z.rev}</span>
//
// From the Select-String output the line wraps as:
//   ...d:var(--b1);color:#1565c0">Rev ${z.rev}</span><span s
// — it is on one line with no embedded newline in the zc() cell.
// But P-02 previously failed, so the anchor string used was wrong.
// Use the tightest unique surrounding context to be safe.
// ===========================================================================
src = applyPatch(
  src,
  'P-02 — zc() cell: DRAFT-aware Rev label',
  'eight:700;padding:1px 5px;border-radius:2px;background:var(--b1);color:#1565c0">Rev ${z.rev}</span>',
  "eight:700;padding:1px 5px;border-radius:2px;background:var(--b1);color:#1565c0\">${z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev}</span>"
);

// ===========================================================================
// P-03 — showD() detail panel
//
// REAL source — the detail panel has a newline inside the template literal:
//   color:#1565c0">Rev\n           ${z.rev}</span></div>
//
// Exact string with the newline and leading spaces as they appear in the file.
// ===========================================================================
src = applyPatch(
  src,
  'P-03 — showD() detail panel: DRAFT-aware Rev label',
  'radius:2px;background:var(--b1);color:#1565c0">Rev\n           ${z.rev}</span></div>',
  "radius:2px;background:var(--b1);color:#1565c0\">${z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev}</span></div>"
);

// ===========================================================================
// Write
// ===========================================================================
fs.writeFileSync(TARGET, src, 'utf8');

console.log('\n✅  P-02 and P-03 applied successfully.');
console.log('    File written: ' + TARGET);
console.log('\nNext steps:');
console.log('  1. .\\node_modules\\.bin\\heft clean');
console.log('  2. .\\node_modules\\.bin\\heft build --production');
console.log('  3. .\\node_modules\\.bin\\heft package-solution --production');
console.log('  4. Add-PnPApp -Path .\\sharepoint\\solution\\imp-9177-spfx.sppkg -Scope Site -Overwrite -Publish');
console.log('  5. Ctrl+Shift+R on QMS-Portal.aspx\n');
