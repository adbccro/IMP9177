/**
 * patch_drm_draft_rev_p03.js
 * IMP9177 — QMS Portal  |  Doc Repo: DRAFT label fix — P-03 only (showD detail panel)
 *
 * P-01 and P-02 already applied. This fixes the one remaining `Rev ${z.rev}`
 * in the showD() detail panel, which has a literal newline inside the template literal.
 *
 * Run from SPFx root:
 *   node .\patch_drm_draft_rev_p03.js
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

console.log('\n=== patch_drm_draft_rev_p03.js ===');
console.log('Target: ' + TARGET + '\n');

let src = readFile(TARGET);

// Confirm P-01 and P-02 are already in place
if (!src.includes("return'DRAFT'")) {
  console.error('[FAIL] P-01 does not appear to be applied — _drmRev still missing DRAFT logic.');
  console.error('       Run patch_drm_draft_rev.js first.\n');
  process.exit(1);
}
console.log('[OK]   P-01 pre-check passed');

if (!src.includes("z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev")) {
  console.error('[FAIL] P-02 does not appear to be applied — zc() still missing DRAFT logic.');
  console.error('       Run patch_drm_draft_rev_p02p03.js first.\n');
  process.exit(1);
}
console.log('[OK]   P-02 pre-check passed');

// ===========================================================================
// P-03 — showD() detail panel
//
// The exact string in the file (line 2422, inside the zonesH template literal):
//   color:#1565c0">Rev
//            ${z.rev}</span></div>
//
// PowerShell wraps at ~60 chars from column 12, so the actual indentation
// is whatever whitespace sits between the newline and "${z.rev}".
// We use a regex to handle any amount of whitespace after the newline
// so we don't get burned by tab vs space or column count differences.
// ===========================================================================
const search  = /color:#1565c0">Rev\s*\n\s*\$\{z\.rev\}<\/span><\/div>/;
const replace  = "color:#1565c0\">${z.rev==='DRAFT'?'DRAFT':'Rev '+z.rev}</span></div>";

if (!search.test(src)) {
  // Fallback: try without the newline — maybe the file is on one line after all
  const searchOneLine = 'color:#1565c0">Rev ${z.rev}</span></div>';
  if (!src.includes(searchOneLine)) {
    console.error('\n[FAIL] P-03 — cannot locate detail panel Rev ${z.rev} span.');
    console.error('       Neither multiline nor single-line form found.');
    console.error('       Run this to inspect the file:');
    console.error('         Select-String -Path .\\src\\webparts\\qmsPortalWebPart\\QmsPortalWebPart.ts -Pattern "1565c0.*Rev|Rev.*z\\.rev"');
    process.exit(1);
  }
  // Single-line fallback
  const result = src.split(searchOneLine).join(replace);
  if (result === src) {
    console.error('\n[FAIL] P-03 — single-line replacement produced no change.\n');
    process.exit(1);
  }
  src = result;
  console.log('[OK]   P-03 — showD() detail panel: DRAFT-aware Rev label (single-line form)');
} else {
  src = src.replace(search, replace);
  console.log('[OK]   P-03 — showD() detail panel: DRAFT-aware Rev label (multiline form)');
}

// Verify no bare `Rev ${z.rev}` remains anywhere
const remaining = (src.match(/Rev \$\{z\.rev\}/g) || []).length;
if (remaining > 0) {
  console.warn('[WARN] ' + remaining + ' occurrence(s) of `Rev ${z.rev}` still remain — review manually.');
} else {
  console.log('[OK]   Verification: no bare `Rev ${z.rev}` remaining ✅');
}

fs.writeFileSync(TARGET, src, 'utf8');

console.log('\n✅  P-03 applied successfully.');
console.log('    File written: ' + TARGET);
console.log('\nNext steps:');
console.log('  1. .\\node_modules\\.bin\\heft clean');
console.log('  2. .\\node_modules\\.bin\\heft build --production');
console.log('  3. .\\node_modules\\.bin\\heft package-solution --production');
console.log('  4. Add-PnPApp -Path .\\sharepoint\\solution\\imp-9177-spfx.sppkg -Scope Site -Overwrite -Publish');
console.log('  5. Ctrl+Shift+R on QMS-Portal.aspx\n');
