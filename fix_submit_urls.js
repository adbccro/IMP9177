// fix_submit_urls.js
const fs = require('fs');
const f = './src/webparts/qmsPortalWebPart/QmsPortalWebPart.ts';
let src = fs.readFileSync(f, 'utf8');

// Fix DCOs URL - unquoted template literal
src = src.replace(
  "base2 + /_api/web/lists/getbytitle('QMS_DCOs')/items( + dcoItem2.Id + ),",
  "base2 + \"/_api/web/lists/getbytitle('QMS_DCOs')/items(\" + dcoItem2.Id + \")\","
);

// Fix RoutingHistory URL - unquoted template literal
src = src.replace(
  "base2 + /_api/web/lists/getbytitle('QMS_RoutingHistory')/items,",
  "base2 + \"/_api/web/lists/getbytitle('QMS_RoutingHistory')/items\","
);

fs.writeFileSync(f, src, 'utf8');

// Verify
const idx = src.indexOf('dcoItem2.Id');
console.log('Result:', src.substring(idx - 80, 200));
console.log('Done');
