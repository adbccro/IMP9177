'use strict';
const fs = require('fs');

const SP_SITE  = 'https://adbccro.sharepoint.com/sites/IMP9177';
const SP_TOKEN = process.env.SP_ACCESS_TOKEN || '';

if (!SP_TOKEN) {
  console.error(
    'ERROR: SP_ACCESS_TOKEN not set.\n' +
    '  Connect-PnPOnline -Url "https://adbccro.sharepoint.com/sites/IMP9177" -ClientId "ba48ac81-6f23-43bd-9797-ec2866071102" -Interactive\n' +
    '  $env:SP_ACCESS_TOKEN = Get-PnPAccessToken\n' +
    '  node .\\save_pj_to_sp.js'
  );
  process.exit(1);
}

const results = JSON.parse(fs.readFileSync('pj_results.json', 'utf8'));

const PATCHES = [
  { dcoId: 'DCO-0001', itemId: 1, data: results['DCO-0001'] },
  { dcoId: 'DCO-0002', itemId: 4, data: results['DCO-0002'] },
];

async function patch(itemId, body) {
  const url = `${SP_SITE}/_api/web/lists/getbytitle('QMS_DCOs')/items(${itemId})`;
  const r = await fetch(url, {
    method: 'POST',
    headers: {
      'Authorization':  `Bearer ${SP_TOKEN}`,
      'Accept':         'application/json;odata=nometadata',
      'Content-Type':   'application/json;odata=nometadata',
      'IF-MATCH':       '*',
      'X-HTTP-Method':  'MERGE',
    },
    body: JSON.stringify(body),
  });
  if (!r.ok) {
    const txt = await r.text();
    throw new Error(`HTTP ${r.status}: ${txt.slice(0, 300)}`);
  }
}

async function main() {
  for (const { dcoId, itemId, data } of PATCHES) {
    const docCount = Object.keys(data).length;
    process.stdout.write(`Saving ${dcoId} (item ${itemId}, ${docCount} docs)... `);
    await patch(itemId, { DCO_DocPurposes: JSON.stringify(data) });
    console.log('OK');
  }
  console.log('\nBoth DCO records updated successfully.');
}

main().catch(err => {
  console.error('FAILED:', err.message);
  process.exit(1);
});
