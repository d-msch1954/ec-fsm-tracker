// scripts/extract.js  — run once: node scripts/extract.js
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { dirname, join } from 'path';
import { fileURLToPath } from 'url';

const root = dirname(fileURLToPath(import.meta.url));
const html = readFileSync(join(root, '..', 'EC_FieldService_Tracker.html'), 'utf8');

const scriptRe = /<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/gi;
const scripts = [];
let m;
while ((m = scriptRe.exec(html)) !== null) scripts.push(m[1].trim());

const styleRe = /<style(?:\s[^>]*)?>([\s\S]*?)<\/style>/gi;
const styles = [];
while ((m = styleRe.exec(html)) !== null) styles.push(m[1].trim());

// SheetJS is the script block containing the XLSX global
const xlsxScript = scripts.find(s =>
  s.includes('var XLSX') || s.includes('XLSX.version') || s.includes('"SheetJS"')
);
const appScripts = scripts.filter(s => s !== xlsxScript);

mkdirSync(join(root, '..', 'vendor'), { recursive: true });
mkdirSync(join(root, '..', 'src'), { recursive: true });

if (!xlsxScript) { console.error('ERROR: SheetJS not found — check the HTML file'); process.exit(1); }

writeFileSync(join(root, '..', 'vendor', 'xlsx.full.min.js'), xlsxScript, 'utf8');
writeFileSync(join(root, '..', 'src', 'core.ui.extracted.js'),
  appScripts.join('\n\n// ─── next script block ───\n\n'), 'utf8');
writeFileSync(join(root, '..', 'src', 'styles.css.bak'),
  styles.join('\n\n/* ─── next style block ─── */\n\n'), 'utf8');

console.log(`Scripts: ${scripts.length} blocks | SheetJS: ${Math.round(xlsxScript.length/1024)}KB | App JS: ${Math.round(appScripts.reduce((a,s)=>a+s.length,0)/1024)}KB`);
console.log(`Styles: ${styles.length} blocks`);
console.log('Done. Review src/core.ui.extracted.js then split into core.js + ui.js');
