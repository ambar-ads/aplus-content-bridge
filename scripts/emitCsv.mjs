/**
 * Menulis CSV ForWeb dari katalog ke legacy-site/, untuk diuji di salinan situs lama sebelum
 * file production di GitHub ditimpa.
 *
 * Marketing memakai tombol unduh di UI; script ini untuk developer dan untuk pemeriksaan
 * otomatis.
 */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, loadModule } from './_bundle.mjs';

const { serializeForWebCsv } = await loadModule('src/lib/legacyCsv.ts');
const { catalogToRows } = await loadModule('src/lib/catalogStore.ts');
const { FORWEB_FILENAME } = await loadModule('src/lib/config.ts');
const { readMasterWorkbook, MASTER_FILENAME } = await loadModule('src/lib/masterWorkbook.ts');

// Alur sebenarnya berangkat dari file Excel master; catalog.json hanya cadangan untuk kondisi
// awal, sebelum file master pertama dibangkitkan.
const masterPath = path.join(ROOT, 'data', MASTER_FILENAME);
const catalogPath = path.join(ROOT, 'data/catalog.json');

let rows;
if (fs.existsSync(masterPath)) {
  const master = await readMasterWorkbook(fs.readFileSync(masterPath));
  rows = master.rows;
  console.log(`Sumber: ${MASTER_FILENAME}`);
  master.warnings.forEach((w) => console.log(`  ! ${w}`));
} else if (fs.existsSync(catalogPath)) {
  rows = catalogToRows(JSON.parse(fs.readFileSync(catalogPath, 'utf8')));
  console.log('Sumber: data/catalog.json');
} else {
  console.error('Tidak ada sumber data. Jalankan: npm run seed:catalog lalu npm run build:master');
  process.exit(1);
}
const csv = serializeForWebCsv(rows);

const out = path.join(ROOT, 'legacy-site', FORWEB_FILENAME);
fs.writeFileSync(out, csv, 'utf8');

console.log(`${rows.length} produk -> ${path.relative(ROOT, out)}`);
console.log(`${Buffer.byteLength(csv, 'utf8').toLocaleString('id-ID')} byte, UTF-8 ber-BOM`);
console.log('\nBuka legacy-site/index.html untuk mengujinya sebelum diunggah ke GitHub.');
