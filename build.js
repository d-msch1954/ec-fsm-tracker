#!/usr/bin/env node
// Copies the self-contained HTML to docs/index.html for GitHub Pages.
// When src/ decomposition is added later, update this to inline src/core.js,
// src/ui.js, src/styles.css, and vendor/xlsx.full.min.js into the template.

import { copyFileSync, mkdirSync } from 'fs';
import { join, dirname } from 'path';
import { fileURLToPath } from 'url';

const root = dirname(fileURLToPath(import.meta.url));
const src  = join(root, 'EC_FieldService_Tracker.html');
const dest = join(root, 'docs', 'index.html');

mkdirSync(join(root, 'docs'), { recursive: true });
copyFileSync(src, dest);
console.log(`Built → docs/index.html`);
