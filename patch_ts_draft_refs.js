/**
 * patch_ts_draft_refs.js
 * IMP9177 — Fix all remaining Rev A / Rev B / _RevA_ / _RevB_ references in TypeScript
 *
 * Run from SPFx root:
 *   node .\patch_ts_draft_refs.js
 *
 * Patches (in order):
 *   P-01  _allDocs: all rev:'Rev A' and rev:'Rev B' → rev:'DRAFT'
 *   P-02  _FM map (lines 1521-1528): _RevA_ filenames → correct _DRAFT_ filenames
 *   P-03  _fn fallback (line 1545): docId + '_RevA.docx' → docId + '_DRAFT_.docx'
 *   P-04  Download button (line 1563): data-dl-name="${docId}_RevA.docx" → _DRAFT_.docx
 *   P-05  docFileMap fallback (line 1653): docId + '_RevA.docx' → docId + '_DRAFT_.docx'
 *   P-06  docIdAudit replace (line 1742): .replace("_RevA.docx","") → .replace("_DRAFT_.docx","")
 *   P-07  fileMap fallbacks (lines 2452, 2488): docId + '_RevA.docx' → docId + '_DRAFT_.docx'
 *
 * STRICT MODE: anchor must exist; exits on failure; file written only on full success.
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

function applyPatch(src, patchId, search, replacement, expectedCount) {
  const count = (search instanceof RegExp)
    ? (src.match(search) || []).length
    : src.split(search).length - 1;

  if (count === 0) {
    console.error('\n[FAIL] ' + patchId + ' — search not found.');
    if (!(search instanceof RegExp)) console.error('  Search: ' + search);
    process.exit(1);
  }
  if (expectedCount && count !== expectedCount) {
    console.warn('[WARN] ' + patchId + ' — expected ' + expectedCount + ' occurrence(s), found ' + count + '. Replacing all.');
  }

  const result = (search instanceof RegExp)
    ? src.replace(search, replacement)
    : src.split(search).join(replacement);

  if (result === src) {
    console.error('\n[FAIL] ' + patchId + ' — replacement produced no change.\n');
    process.exit(1);
  }

  console.log('[OK]   ' + patchId + ' (' + count + ' occurrence' + (count !== 1 ? 's' : '') + ')');
  return result;
}

console.log('\n=== patch_ts_draft_refs.js ===');
console.log('Target: ' + TARGET + '\n');

let src = readFile(TARGET);

// ===========================================================================
// P-01 — _allDocs: rev:'Rev A' and rev:'Rev B' → rev:'DRAFT'
// All 26 entries in the _allDocs block use rev:'Rev A' or rev:'Rev B'.
// Simple global replace — safe because rev:'Rev A' and rev:'Rev B' only
// appear in this hardcoded catalogue, not in any dynamic logic.
// ===========================================================================
src = applyPatch(src, "P-01a — _allDocs rev:'Rev A' → 'DRAFT'",
  /rev:'Rev A'/g, "rev:'DRAFT'");

src = applyPatch(src, "P-01b — _allDocs rev:'Rev B' → 'DRAFT'",
  /rev:'Rev B'/g, "rev:'DRAFT'");

// ===========================================================================
// P-02 — _FM map filenames: _RevA_ → _DRAFT_ for FM-004 through FM-ALG
//
// REAL source (lines 1521-1528):
//   'FM-004':'FM-004_Supplier_Evaluation_Form_RevA.docx',
//   'FM-005':'FM-005_Ingredient_Approval_Form_RevA.docx',
//   'FM-006':'FM-006_Material_Receipt_Log_RevA.docx',
//   'FM-007':'FM-007_CoA_Review_Checklist_RevA.docx',
//   'FM-ALG':'FM-ALG_Allergen_Status_Record_RevA.docx',
//
// Note: these are the OLD names that don't match actual SP filenames.
// We align them with the actual Draft filenames from the SP listing.
// ===========================================================================
const fmReplacements = [
  ["'FM-004':'FM-004_Supplier_Evaluation_Form_RevA.docx'",
   "'FM-004':'FM-004_DRAFT_Approved_Supplier_List_TEMPLATE.docx'"],
  ["'FM-005':'FM-005_Ingredient_Approval_Form_RevA.docx'",
   "'FM-005':'FM-005_DRAFT_Receiving_Log_TEMPLATE.docx'"],
  ["'FM-006':'FM-006_Material_Receipt_Log_RevA.docx'",
   "'FM-006':'FM-006_DRAFT_RawMaterial_Spec_Sheet_TEMPLATE.docx'"],
  ["'FM-007':'FM-007_CoA_Review_Checklist_RevA.docx'",
   "'FM-007':'FM-007_DRAFT_Material_Hold_Label_TEMPLATE.docx'"],
  ["'FM-ALG':'FM-ALG_Allergen_Status_Record_RevA.docx'",
   "'FM-ALG':'FM-ALG_DRAFT_Allergen_Status_Record_TEMPLATE.docx'"],
];

fmReplacements.forEach(([search, replacement], i) => {
  src = applyPatch(src, 'P-02' + String.fromCharCode(97 + i) + ' — _FM map: ' + search.split(':')[0].replace(/'/g,''), search, replacement, 1);
});

// ===========================================================================
// P-03 — _fn fallback: docId + '_RevA.docx' → docId + '_DRAFT_.docx'
// Line 1545: const _fn = _FM[docId] || (docId + '_RevA.docx');
// ===========================================================================
src = applyPatch(src, "P-03 — _fn fallback '_RevA.docx'",
  "const _fn   = _FM[docId] || (docId + '_RevA.docx')",
  "const _fn   = _FM[docId] || (docId + '_DRAFT_.docx')", 1);

// ===========================================================================
// P-04 — Download button data-dl-name: "${docId}_RevA.docx" → _DRAFT_.docx
// Line 1563
// ===========================================================================
src = applyPatch(src, "P-04 — download button data-dl-name",
  'data-dl-name="${docId}_RevA.docx"',
  'data-dl-name="${docId}_DRAFT_.docx"', 1);

// ===========================================================================
// P-05 — docFileMap fallback: docId + '_RevA.docx' → docId + '_DRAFT_.docx'
// Line 1653: const docFileName = docFileMap[docId] || (docId + '_RevA.docx');
// ===========================================================================
src = applyPatch(src, "P-05 — docFileMap fallback '_RevA.docx'",
  "docFileMap[docId] || (docId + '_RevA.docx')",
  "docFileMap[docId] || (docId + '_DRAFT_.docx')", 1);

// ===========================================================================
// P-06 — docIdAudit: name.replace("_RevA.docx", "") → name.replace("_DRAFT_.docx", "")
// Line 1742
// ===========================================================================
src = applyPatch(src, 'P-06 — docIdAudit replace "_RevA.docx"',
  'name.replace("_RevA.docx", "")',
  'name.replace("_DRAFT_.docx", "")', 1);

// ===========================================================================
// P-07 — fileMap fallbacks at lines 2452 and 2488
// Both read: const fn = fileMap[docId] || (docId + '_RevA.docx');
// ===========================================================================
src = applyPatch(src, "P-07 — fileMap fallbacks '_RevA.docx' (×2)",
  "fileMap[docId] || (docId + '_RevA.docx')",
  "fileMap[docId] || (docId + '_DRAFT_.docx')", 2);

// ===========================================================================
// Final verification — no bare _RevA_ or rev:'Rev A'/'Rev B' should remain
// (excluding Official zone which legitimately uses RevA)
// ===========================================================================
const revACount = (src.match(/rev:'Rev [AB]'/g) || []).length;
const revAFallback = (src.match(/_RevA\.docx/g) || []).length;
const revBFallback = (src.match(/_RevB\.docx/g) || []).length;

if (revACount > 0)    console.warn('[WARN] ' + revACount + ' rev:\'Rev A/B\' still remain — check manually');
if (revAFallback > 0) console.warn('[WARN] ' + revAFallback + ' _RevA.docx references remain — may be in Official zone or _addDownloadBar (review)');
if (revBFallback > 0) console.warn('[WARN] ' + revBFallback + ' _RevB.docx references remain — may be in Official zone (review)');

if (revACount === 0 && revAFallback === 0 && revBFallback === 0) {
  console.log('[OK]   Verification: no stale Rev A/B references remaining ✅');
}

fs.writeFileSync(TARGET, src, 'utf8');

console.log('\n✅  All patches applied successfully.');
console.log('    File written: ' + TARGET);
console.log('\nNext steps:');
console.log('  1. .\\node_modules\\.bin\\heft clean');
console.log('  2. .\\node_modules\\.bin\\heft build --production');
console.log('  3. .\\node_modules\\.bin\\heft package-solution --production');
console.log('  4. Add-PnPApp -Path .\\sharepoint\\solution\\imp-9177-spfx.sppkg -Scope Site -Overwrite -Publish\n');
