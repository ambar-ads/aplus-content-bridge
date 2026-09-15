/**
 * Builds the demo bundle: the tool at the root, the A+ Content page under /site/.
 *
 * Run after `vite build`. Two things it deliberately does NOT do:
 *
 * It does not copy `legacy-site/A+ Content Source(ForWeb).csv`, which is git-ignored and may be
 * stale on whoever's machine runs this. The CSV is generated from the catalogue instead, so the
 * demo always shows what the current pipeline actually produces.
 *
 * And it does not ship the A+ page as-is. A second copy of a live site is exactly the thing
 * someone bookmarks by mistake, so the demo copy carries a banner saying what it is and linking
 * to the real one.
 */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, loadModule } from './_bundle.mjs';

const DIST = path.join(ROOT, 'dist');
const SITE = path.join(DIST, 'site');
const LEGACY = path.join(ROOT, 'legacy-site');

const { serializeForWebCsv } = await loadModule('src/lib/legacyCsv.ts');
const { catalogToRows } = await loadModule('src/lib/catalogStore.ts');
const { readMasterWorkbook, MASTER_FILENAME } = await loadModule('src/lib/masterWorkbook.ts');
const { FORWEB_FILENAME } = await loadModule('src/lib/config.ts');

if (!fs.existsSync(DIST)) {
  console.error('dist/ belum ada. Jalankan dulu: npm run build');
  process.exit(1);
}

// Prefer the master workbook when it exists, so the demo reflects the real round trip.
const masterPath = path.join(ROOT, 'data', MASTER_FILENAME);
const catalogPath = path.join(ROOT, 'data/catalog.json');

let rows;
if (fs.existsSync(masterPath)) {
  rows = (await readMasterWorkbook(fs.readFileSync(masterPath), MASTER_FILENAME)).rows;
  console.log(`Sumber data demo : ${MASTER_FILENAME}`);
} else if (fs.existsSync(catalogPath)) {
  rows = catalogToRows(JSON.parse(fs.readFileSync(catalogPath, 'utf8')));
  console.log('Sumber data demo : data/catalog.json');
} else {
  console.error('Tidak ada sumber data. Jalankan: npm run seed:catalog');
  process.exit(1);
}

fs.mkdirSync(SITE, { recursive: true });

for (const asset of ['asus-logo.png', 'tutor.jpg']) {
  fs.copyFileSync(path.join(LEGACY, asset), path.join(SITE, asset));
}

const BANNER = `  <div style="background:#fff8e5;border-bottom:1px solid #f0d98c;color:#7a5c00;padding:.6rem 1rem;text-align:center;font-size:.875rem;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Arial,sans-serif;">
    <strong>Demo copy</strong> — not the live site. The live page is at
    <a href="https://asus-commercial.github.io/AplusContent/" style="color:#0072c6;">asus-commercial.github.io/AplusContent</a>.
  </div>
`;

let html = fs.readFileSync(path.join(LEGACY, 'index.html'), 'utf8');
if (!html.includes('<body>')) {
  console.error('Tidak menemukan <body> di legacy-site/index.html');
  process.exit(1);
}
html = html.replace('<body>', '<body>\n' + BANNER);
fs.writeFileSync(path.join(SITE, 'index.html'), html);

const csv = serializeForWebCsv(rows);
fs.writeFileSync(path.join(SITE, FORWEB_FILENAME), csv, 'utf8');

console.log(`Halaman A+ demo  : dist/site/ (${rows.length} produk, ${Buffer.byteLength(csv, 'utf8').toLocaleString('id-ID')} byte CSV)`);
console.log('');
console.log('Isi dist/:');
for (const f of fs.readdirSync(DIST)) console.log('  ' + f + (fs.statSync(path.join(DIST, f)).isDirectory() ? '/' : ''));
console.log('\nAlat di / dan halaman A+ di /site/');
