/**
 * patch_remove_iframe.js
 * IMP9177 — QMS Portal  |  Option 3: Remove iframe, render directly into domElement
 *
 * Run from SPFx root:
 *   node .\patch_remove_iframe.js
 *
 * Anchors verified against real QmsPortalWebPart.ts (Select-String output 2026-04-25).
 *
 * Patches (in order):
 *   P-01  Remove iframe creation / style / appendChild block  (lines 714-717)
 *   P-02  Remove ifrDoc open/write/close + srcdoc fallback    (lines 719-722)
 *         Inject this.domElement.innerHTML = _extractBodyContent(buildShell())
 *   P-03  Remove load-event listener; call attach() directly  (line 727)
 *   P-04  this._iframe.contentDocument  → document            (line 725)
 *   P-05  this._iframe?.contentDocument → document            (lines 778,1109,2375,2380,2423,2425)
 *   P-06  this._iframe.contentWindow    → window              (residual non-optional-chained)
 *   P-07  this._iframe?.contentWindow   → window              (lines 1111,2377,2427)
 *   P-08  Remove stale if(!w)return null-guard in _generateDCOReport
 *   P-09  Belt-and-suspenders: replace any surviving 1100px height with 100vh
 *   P-10  Add private _extractBodyContent() helper to class
 *
 * STRICT MODE: every patch verifies its anchor exists BEFORE replacement.
 * Missing anchor → report which patch failed → process.exit(1).
 * File is never written if any patch fails.
 */

'use strict';

const fs   = require('fs');
const path = require('path');

// ---------------------------------------------------------------------------
// Target
// ---------------------------------------------------------------------------
const TARGET = path.resolve(
  __dirname,
  'src', 'webparts', 'qmsPortalWebPart', 'QmsPortalWebPart.ts'
);

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
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
    console.error('\n[FAIL] ' + patchId + ' — anchor not found in source.');
    console.error('  Anchor: ' + anchor);
    console.error('  Source may already be patched or anchor string has changed.\n');
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

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
console.log('\n=== patch_remove_iframe.js ===');
console.log('Target: ' + TARGET + '\n');

let src = readFile(TARGET);

// ===========================================================================
// P-01 — Remove iframe creation block
//
// REAL source (lines 714-717):
//   this._iframe = document.createElement('iframe');
//   this._iframe.style.cssText = 'width:100%;height:1100px;border:none;display:block;';
//   this._iframe.setAttribute('sandbox', 'allow-scripts allow-same-origin allow-forms allow-popups');
//   this.domElement.appendChild(this._iframe);
// ===========================================================================
src = applyPatch(
  src,
  'P-01 — remove iframe creation block',
  "this._iframe = document.createElement('iframe')",
  /this\._iframe\s*=\s*document\.createElement\(\s*['"]iframe['"]\s*\);\s*\n\s*this\._iframe\.style\.cssText\s*=\s*['"][^'"]*['"];\s*\n\s*this\._iframe\.setAttribute\(\s*['"]sandbox['"][^)]*\);\s*\n\s*this\.domElement\.appendChild\(\s*this\._iframe\s*\);/,
  "// [PATCH P-01] iframe removed — application renders directly in domElement\n    this.domElement.style.height = 'auto';"
);

// ===========================================================================
// P-02 — Remove ifrDoc block + srcdoc fallback; inject into domElement
//
// REAL source (lines 719-722):
//   const ifrDoc = this._iframe.contentDocument ||
//     (this._iframe.contentWindow && this._iframe.contentWindow.document);
//   if (ifrDoc) { ifrDoc.open(); ifrDoc.write(buildShell()); ifrDoc.close(); }
//   else { this._iframe.srcdoc = buildShell(); }
// ===========================================================================
src = applyPatch(
  src,
  'P-02 — remove ifrDoc/srcdoc block; inject into domElement',
  'ifrDoc.write(buildShell())',
  /const ifrDoc\s*=\s*this\._iframe\.contentDocument\s*\|\|\s*\(\s*this\._iframe\.contentWindow\s*&&\s*this\._iframe\.contentWindow\.document\s*\);\s*\n\s*if\s*\(\s*ifrDoc\s*\)\s*\{\s*ifrDoc\.open\(\);\s*ifrDoc\.write\(buildShell\(\)\);\s*ifrDoc\.close\(\);\s*\}\s*\n\s*else\s*\{\s*this\._iframe\.srcdoc\s*=\s*buildShell\(\);\s*\}/,
  '// [PATCH P-02] Render shell directly — no iframe\n    this.domElement.innerHTML = this._extractBodyContent(buildShell());'
);

// ===========================================================================
// P-03 — Remove load-event listener; call attach() directly
//
// REAL source (line 727):
//   else this._iframe.addEventListener('load', attach, { once: true });
// ===========================================================================
src = applyPatch(
  src,
  'P-03 — remove load listener; call attach() directly',
  "this._iframe.addEventListener('load', attach",
  /else\s+this\._iframe\.addEventListener\(\s*['"]load['"]\s*,\s*attach\s*,\s*\{\s*once\s*:\s*true\s*\}\s*\);/,
  '// [PATCH P-03] No load event needed — DOM is already populated\n    attach();'
);

// ===========================================================================
// P-04 — this._iframe.contentDocument → document  (non-optional-chained)
//
// REAL source: line 725 — const d = this._iframe.contentDocument;
// (ifrDoc line already removed by P-02, but P-02 regex consumed those;
//  line 725 is a separate standalone assignment)
// ===========================================================================
if (src.includes('this._iframe.contentDocument')) {
  src = applyPatch(
    src,
    'P-04 — this._iframe.contentDocument → document',
    'this._iframe.contentDocument',
    /this\._iframe\.contentDocument/g,
    'document'
  );
} else {
  console.log('[SKIP] P-04 — no non-optional-chained contentDocument remaining');
}

// ===========================================================================
// P-05 — this._iframe?.contentDocument → document  (optional-chained)
//
// REAL source: lines 778, 1109, 2375, 2380, 2423, 2425
// ===========================================================================
src = applyPatch(
  src,
  'P-05 — this._iframe?.contentDocument → document',
  'this._iframe?.contentDocument',
  /this\._iframe\?\.contentDocument/g,
  'document'
);

// ===========================================================================
// P-06 — this._iframe.contentWindow → window  (non-optional-chained)
//
// Belt-and-suspenders — the ifrDoc block (lines 719-720) was removed by P-02,
// but catch any other non-optional-chained usages if present.
// ===========================================================================
if (src.includes('this._iframe.contentWindow')) {
  src = applyPatch(
    src,
    'P-06 — this._iframe.contentWindow → window',
    'this._iframe.contentWindow',
    /this\._iframe\.contentWindow/g,
    'window'
  );
} else {
  console.log('[SKIP] P-06 — no non-optional-chained contentWindow remaining');
}

// ===========================================================================
// P-07 — this._iframe?.contentWindow → window  (optional-chained)
//
// REAL source: lines 1111, 2377, 2427
// ===========================================================================
src = applyPatch(
  src,
  'P-07 — this._iframe?.contentWindow → window',
  'this._iframe?.contentWindow',
  /this\._iframe\?\.contentWindow/g,
  'window'
);

// ===========================================================================
// P-08 — Remove stale null-guard on w in _generateDCOReport
//
// After P-07, line 2377 becomes:
//   const w=window as any;if(!w)return;
// `window` is never null — the guard is a dead branch left over from the iframe.
// Handle both minified (no spaces) and formatted (with spaces) variants.
// ===========================================================================
const guardMinified  = 'const w=window as any;if(!w)return;';
const guardFormatted = 'const w = window as any; if (!w) return;';
const guardFormatted2 = 'const w = window as any;\n    if (!w) return;';

if (src.includes(guardMinified)) {
  src = applyPatch(
    src,
    'P-08 — remove stale w null-guard (minified)',
    guardMinified,
    guardMinified,
    'const w=window as any; // [PATCH P-08] window is never null — guard removed'
  );
} else if (src.includes(guardFormatted)) {
  src = applyPatch(
    src,
    'P-08 — remove stale w null-guard (formatted)',
    guardFormatted,
    guardFormatted,
    'const w = window as any; // [PATCH P-08] window is never null — guard removed'
  );
} else if (src.includes(guardFormatted2)) {
  src = applyPatch(
    src,
    'P-08 — remove stale w null-guard (multiline)',
    guardFormatted2,
    guardFormatted2,
    'const w = window as any; // [PATCH P-08] window is never null — guard removed'
  );
} else {
  console.log('[SKIP] P-08 — w null-guard not found in expected forms; verify _generateDCOReport manually');
}

// ===========================================================================
// P-09 — Belt-and-suspenders: replace any surviving 1100px height
//         P-01 removed the iframe style block, but buildShell() may contain
//         a matching rule for the app shell container div.
//         Replace with 100vh so the app fills the viewport naturally.
// ===========================================================================
if (src.includes('1100px')) {
  src = applyPatch(
    src,
    'P-09 — replace remaining 1100px with 100vh',
    '1100px',
    /height\s*:\s*1100px/g,
    'height:100vh'
  );
} else {
  console.log('[SKIP] P-09 — no remaining 1100px height references');
}

// ===========================================================================
// P-10 — Add private _extractBodyContent() helper to the class
//
// buildShell() returns a full HTML document string (<html><head>...<body>).
// We extract the <style> blocks + <body> inner content for domElement injection.
//
// Primary anchor: `public getPropertyPaneConfiguration` (standard SPFx method).
// Fallback: before the last `\n}` in the file.
// ===========================================================================
const extractHelper = [
  '',
  '  /**',
  '   * [PATCH P-10] Extract <style> blocks and <body> content from buildShell() output',
  '   * for direct injection into this.domElement — no iframe wrapper.',
  '   */',
  '  private _extractBodyContent(fullHtml: string): string {',
  '    const styles: string[] = [];',
  '    const styleRx = /<style[^>]*>([\\s\\S]*?)<\\/style>/gi;',
  '    let m: RegExpExecArray | null;',
  '    while ((m = styleRx.exec(fullHtml)) !== null) {',
  "      styles.push('<style>' + m[1] + '</style>');",
  '    }',
  '    const bodyMatch = fullHtml.match(/<body[^>]*>([\\s\\S]*?)<\\/body>/i);',
  '    const body = bodyMatch ? bodyMatch[1] : fullHtml;',
  "    return styles.join('\\n') + '\\n' + body;",
  '  }',
  '',
  ''
].join('\n');

const P10_ANCHOR = 'public getPropertyPaneConfiguration';
if (src.includes(P10_ANCHOR)) {
  src = applyPatch(
    src,
    'P-10 — insert _extractBodyContent() helper',
    P10_ANCHOR,
    P10_ANCHOR,
    extractHelper + 'public getPropertyPaneConfiguration'
  );
} else {
  const lastBrace = src.lastIndexOf('\n}');
  if (lastBrace === -1) {
    console.error('\n[FAIL] P-10 — cannot locate insertion point for _extractBodyContent helper.\n');
    process.exit(1);
  }
  src = src.slice(0, lastBrace) + '\n' + extractHelper + src.slice(lastBrace);
  console.log('[OK]   P-10 (fallback — inserted before final closing brace)');
}

// ===========================================================================
// Write — only reached if every patch above succeeded
// ===========================================================================
fs.writeFileSync(TARGET, src, 'utf8');

console.log('\n✅  All patches applied successfully.');
console.log('    File written: ' + TARGET);
console.log('\nNext steps:');
console.log('  1. git diff src\\webparts\\qmsPortalWebPart\\QmsPortalWebPart.ts');
console.log('  2. .\\node_modules\\.bin\\heft clean');
console.log('  3. .\\node_modules\\.bin\\heft build --production');
console.log('  4. .\\node_modules\\.bin\\heft package-solution --production');
console.log('  5. Add-PnPApp -Path .\\sharepoint\\solution\\imp-9177-spfx.sppkg -Scope Site -Overwrite -Publish');
console.log('  6. Ctrl+Shift+R on QMS-Portal.aspx\n');
