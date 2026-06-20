const fs = require('fs');
const path = require('path');

const srcHtml = fs.readFileSync(path.join(__dirname, '..', 'EC_FieldService_Tracker.html'), 'utf8');
// The SheetJS block is the first <script> that contains the xlsx.js banner.
const start = srcHtml.indexOf('<script>/*! xlsx.js');
if (start === -1) throw new Error('SheetJS banner not found in v1 HTML');
const open = srcHtml.indexOf('>', start) + 1;
const end = srcHtml.indexOf('</script>', open);
if (end === -1) throw new Error('SheetJS closing tag not found');
const js = srcHtml.slice(open, end);
const out = path.join(__dirname, '..', 'vendor', 'xlsx.full.min.js');
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, js, 'utf8');
console.log('Extracted SheetJS:', js.length, 'bytes ->', out);
