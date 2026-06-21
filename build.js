// build.js — inlines src/ files into single self-contained HTML
import { readFileSync, writeFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const root = dirname(fileURLToPath(import.meta.url));

const template = readFileSync(join(root, 'src', 'app.template.html'), 'utf8');
const styles   = readFileSync(join(root, 'src', 'styles.css'), 'utf8');
const xlsx     = readFileSync(join(root, 'vendor', 'xlsx.full.min.js'), 'utf8');
const core     = readFileSync(join(root, 'src', 'core.js'), 'utf8');
const ui       = readFileSync(join(root, 'src', 'ui.js'), 'utf8');

const out = template
  .replace('<!--INLINE:styles-->', `<style>\n${styles}\n</style>`)
  .replace('<!--INLINE:xlsx-->',   `<script>\n${xlsx}\n</script>`)
  .replace('<!--INLINE:core-->',   `<script>\n${core}\n</script>`)
  .replace('<!--INLINE:ui-->',     `<script>\n${ui}\n</script>`);

mkdirSync(join(root, 'docs'), { recursive: true });
writeFileSync(join(root, 'EC_FieldService_Tracker.html'), out, 'utf8');
writeFileSync(join(root, 'docs', 'index.html'), out, 'utf8');
console.log(`Built → ${Math.round(out.length / 1024)}KB`);
