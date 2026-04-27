// generate_pj_all.js
// Generates Purpose & Justification for all docs in DCO-0001 and DCO-0002
// using the same Anthropic API pattern as the QMS portal.
//
// Usage:
//   ANTHROPIC_API_KEY=sk-ant-... node generate_pj_all.js
//
// Or with SharePoint token (to read key from SP and save results back):
//   SP_ACCESS_TOKEN=<token> node generate_pj_all.js
//   (Get token: Connect-PnPOnline ... then Get-PnPAccessToken)

'use strict';
const fs = require('fs');

const SP_SITE  = 'https://adbccro.sharepoint.com/sites/IMP9177';
const SP_TOKEN = process.env.SP_ACCESS_TOKEN || '';
const MODEL    = 'claude-sonnet-4-20250514';

// ── Document lists ────────────────────────────────────────────────────────────

const DCO_DOCS = {
  'DCO-0001': [
    { id: 'QM-001',       title: 'Quality Manual',                              type: 'QM'  },
    { id: 'SOP-QMS-001',  title: 'Management Responsibility',                   type: 'SOP' },
    { id: 'SOP-QMS-002',  title: 'Document Control',                            type: 'SOP' },
    { id: 'SOP-QMS-003',  title: 'Change Control',                              type: 'SOP' },
    { id: 'SOP-PRD-108',  title: 'Finished Product Release Management',         type: 'SOP' },
    { id: 'SOP-RCL-321',  title: 'Product Recall Procedure',                    type: 'SOP' },
    { id: 'SOP-PRD-432',  title: 'Finished Product Specifications',             type: 'SOP' },
    { id: 'SOP-FRS-549',  title: 'Product Specification Sheet',                 type: 'SOP' },
    { id: 'FM-001',       title: 'Master Document Log',                         type: 'FM'  },
    { id: 'FM-002',       title: 'Change Request Form',                         type: 'FM'  },
    { id: 'FM-003',       title: 'Document Change Order Form',                  type: 'FM'  },
    { id: 'FM-025',       title: 'Finished Product Release Review Record',      type: 'FM'  },
    { id: 'FM-027',       title: 'QU/QS Designation Record',                   type: 'FM'  },
    { id: 'FM-030',       title: 'Finished Product Spec Sheet Template',        type: 'FM'  },
    { id: 'FPS-001',      title: 'Lychee VD3 Gummy Finished Product Spec',     type: 'FPS' },
  ],
  'DCO-0002': [
    { id: 'SOP-SUP-001',  title: 'Supplier Qualification',                      type: 'SOP' },
    { id: 'SOP-SUP-002',  title: 'Receiving Inspection',                        type: 'SOP' },
    { id: 'SOP-FS-001',   title: 'Allergen Control',                            type: 'SOP' },
    { id: 'SOP-FS-002',   title: 'Equipment Cleaning',                          type: 'SOP' },
    { id: 'SOP-FS-003',   title: 'Facility Sanitation',                         type: 'SOP' },
    { id: 'SOP-FS-004',   title: 'Environmental Monitoring',                    type: 'SOP' },
    { id: 'SOP-PC-001',   title: 'Pest Sighting Response',                      type: 'SOP' },
    { id: 'FM-004',       title: 'Approved Supplier List',                      type: 'FM'  },
    { id: 'FM-005',       title: 'Receiving Log',                               type: 'FM'  },
    { id: 'FM-006',       title: 'Raw Material Specification Sheet',            type: 'FM'  },
    { id: 'FM-007',       title: 'Material Hold Label',                         type: 'FM'  },
    { id: 'FM-ALG',       title: 'Allergen Status Record',                      type: 'FM'  },
    { id: 'FM-008',       title: 'Supplier CoA Requirements Checklist',         type: 'FM'  },
  ],
};

// ── Helpers ───────────────────────────────────────────────────────────────────

const DELAY_MS = 13000; // 5 req/min limit → 12 s minimum; 13 s for safety

const sleep = ms => new Promise(r => setTimeout(r, ms));

function spHeaders(extra = {}) {
  return {
    'Authorization': `Bearer ${SP_TOKEN}`,
    'Accept': 'application/json;odata=nometadata',
    ...extra,
  };
}

async function spGet(path) {
  const r = await fetch(`${SP_SITE}/_api/${path}`, { headers: spHeaders() });
  if (!r.ok) throw new Error(`SP GET failed ${r.status}: ${path}`);
  return r.json();
}

async function spPatch(listTitle, itemId, body) {
  const r = await fetch(
    `${SP_SITE}/_api/web/lists/getbytitle('${listTitle}')/items(${itemId})`,
    {
      method: 'POST',
      headers: spHeaders({
        'Content-Type': 'application/json;odata=nometadata',
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      }),
      body: JSON.stringify(body),
    }
  );
  if (!r.ok) {
    const txt = await r.text();
    throw new Error(`SP PATCH failed ${r.status}: ${txt.slice(0, 200)}`);
  }
}

async function readApiKeyFromSP() {
  const data = await spGet(
    `web/lists/getbytitle('QMS_Config')/items?$filter=Title eq 'anthropicApiKey'&$select=Title,Cfg_Value&$top=1`
  );
  const key = (data.value || [])[0]?.Cfg_Value || '';
  if (!key) {
    // fallback key name
    const data2 = await spGet(
      `web/lists/getbytitle('QMS_Config')/items?$filter=Title eq 'anthropic_api_key'&$select=Title,Cfg_Value&$top=1`
    ).catch(() => ({ value: [] }));
    return (data2.value || [])[0]?.Cfg_Value || '';
  }
  return key;
}

async function getDcoItemId(dcoTitle) {
  const data = await spGet(
    `web/lists/getbytitle('QMS_DCOs')/items?$filter=Title eq '${dcoTitle}'&$select=Id,Title&$top=1`
  );
  return (data.value || [])[0]?.Id || null;
}

async function generatePJWithRetry(apiKey, doc, dcoId, retries = 3) {
  for (let attempt = 0; attempt < retries; attempt++) {
    try {
      return await generatePJ(apiKey, doc, dcoId);
    } catch (err) {
      if (err.message.includes('429') && attempt < retries - 1) {
        const wait = 60000 + attempt * 15000; // 60s, 75s, 90s
        process.stdout.write(`[RATE LIMIT — waiting ${wait/1000}s] `);
        await sleep(wait);
      } else {
        throw err;
      }
    }
  }
}

async function generatePJ(apiKey, doc, dcoId) {
  const typeLabel = {
    'QM':  'Quality Manual',
    'SOP': 'Standard Operating Procedure',
    'FM':  'Form / Record',
    'FPS': 'Finished Product Specification',
  }[doc.type] || doc.type;

  const prompt =
    `For a QMS document titled '${doc.id} — ${doc.title}' of type '${typeLabel}' ` +
    `for a gummy dietary supplement contract manufacturer (3H Pharmaceuticals LLC) ` +
    `subject to 21 CFR Part 111 and FSMA, this is a first-issue DCO (${dcoId}). ` +
    `Generate a concise Purpose of Change and Justification of Change. ` +
    `Return JSON only: {"purpose": "...", "justification": "..."}`;

  const resp = await fetch('https://api.anthropic.com/v1/messages', {
    method: 'POST',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json',
      'anthropic-dangerous-direct-browser-access': 'true',
    },
    body: JSON.stringify({
      model: MODEL,
      max_tokens: 512,
      messages: [{ role: 'user', content: prompt }],
    }),
  });

  if (!resp.ok) {
    const txt = await resp.text();
    throw new Error(`Anthropic API ${resp.status}: ${txt.slice(0, 200)}`);
  }

  const data = await resp.json();
  const text = data?.content?.[0]?.text || '';
  const match = text.match(/\{[\s\S]*\}/);
  if (!match) throw new Error(`No JSON in response: ${text.slice(0, 100)}`);
  return JSON.parse(match[0]);
}

// ── Main ──────────────────────────────────────────────────────────────────────

async function main() {
  // 1. Resolve API key
  let apiKey = process.env.ANTHROPIC_API_KEY || '';

  if (!apiKey && SP_TOKEN) {
    console.log('Reading API key from QMS_Config...');
    apiKey = await readApiKeyFromSP();
    if (apiKey) console.log('  API key loaded from SharePoint.');
  }

  if (!apiKey) {
    console.error(
      'ERROR: No API key found.\n' +
      '  Set ANTHROPIC_API_KEY env var, or set SP_ACCESS_TOKEN so the key can be read from QMS_Config.\n' +
      '  To get SP token:  Connect-PnPOnline ...  then  $env:SP_ACCESS_TOKEN = Get-PnPAccessToken'
    );
    process.exit(1);
  }

  // Resume from existing results file if present
  let allResults = {};
  if (fs.existsSync('pj_results.json')) {
    try {
      allResults = JSON.parse(fs.readFileSync('pj_results.json', 'utf8'));
      console.log('Resuming from existing pj_results.json...');
    } catch (_) {}
  }

  let totalDocs = 0;
  let firstCall = true;

  for (const [dcoId, docs] of Object.entries(DCO_DOCS)) {
    if (!allResults[dcoId]) allResults[dcoId] = {};
    const pending = docs.filter(d => !allResults[dcoId][d.id]?.purpose);
    const already = docs.length - pending.length;
    console.log(`\n═══ ${dcoId} (${docs.length} docs — ${already} already done, ${pending.length} remaining) ═══`);

    for (const doc of pending) {
      process.stdout.write(`  ${doc.id.padEnd(14)} `);
      if (!firstCall) await sleep(DELAY_MS);
      firstCall = false;
      try {
        const pj = await generatePJWithRetry(apiKey, doc, dcoId);
        allResults[dcoId][doc.id] = pj;
        console.log(`[OK]`);
        console.log(`    Purpose:        ${pj.purpose}`);
        console.log(`    Justification:  ${pj.justification}`);
        totalDocs++;
        // Save progress after every successful call
        fs.writeFileSync('pj_results.json', JSON.stringify(allResults, null, 2));
      } catch (err) {
        console.log(`[FAIL] ${err.message}`);
        allResults[dcoId][doc.id] = { purpose: '', justification: '', error: err.message };
      }
    }

    // Count already-done toward total
    totalDocs += already;
  }

  // 2. Write local JSON backup
  fs.writeFileSync('pj_results.json', JSON.stringify(allResults, null, 2));
  console.log(`\nResults written to pj_results.json`);

  // 3. Save to SharePoint if token is available
  if (SP_TOKEN) {
    console.log('\nSaving to SharePoint...');
    for (const dcoId of Object.keys(allResults)) {
      const itemId = await getDcoItemId(dcoId);
      if (!itemId) {
        console.log(`  [SKIP] ${dcoId} — item not found in QMS_DCOs`);
        continue;
      }
      await spPatch('QMS_DCOs', itemId, {
        DCO_DocPurposes: JSON.stringify(allResults[dcoId]),
      });
      console.log(`  [SAVED] ${dcoId} (item ${itemId})`);
    }
  } else {
    console.log('\nSP_ACCESS_TOKEN not set — skipping SharePoint save.');
    console.log('To save, run:');
    console.log('  $env:SP_ACCESS_TOKEN = Get-PnPAccessToken');
    console.log('  node generate_pj_all.js');
  }

  console.log(`\nDone. ${totalDocs}/28 documents generated.`);
}

main().catch(err => {
  console.error('Fatal:', err.message);
  process.exit(1);
});
