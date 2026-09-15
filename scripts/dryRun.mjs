/**
 * Menjalankan alur bulanan yang sama persis dengan UI, tapi dari terminal.
 *
 * Bukan sekadar alat bantu: ini yang menguji jalur file master -> price list -> CSV terhadap
 * file ASUS sungguhan, termasuk hal yang paling penting untuk dijaga — jumlah produk tidak
 * pernah berkurang setelah import.
 */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, loadModule } from './_bundle.mjs';
import { newestPriceList } from './_files.mjs';

const ASSETS = path.join(ROOT, 'assets');

const { readPriceList } = await loadModule('src/lib/priceListImport.ts');
const { readMasterWorkbook, buildMasterWorkbook, MASTER_FILENAME } = await loadModule('src/lib/masterWorkbook.ts');
const { deriveForWebRow } = await loadModule('src/lib/deriveForWeb.ts');
const { buildCatalogFromRows, mergePriceList, catalogToRows } = await loadModule('src/lib/catalogStore.ts');
const { serializeForWebCsv, isBlankValue } = await loadModule('src/lib/legacyCsv.ts');
const { normalizeLabel } = await loadModule('src/vendor/pmCell.ts');

const masterPath = path.join(ROOT, 'data', MASTER_FILENAME);
if (!fs.existsSync(masterPath)) {
  console.error(`${MASTER_FILENAME} belum ada. Jalankan: npm run build:master`);
  process.exit(1);
}

const master = await readMasterWorkbook(fs.readFileSync(masterPath), MASTER_FILENAME);
console.log(`File master : ${MASTER_FILENAME} — ${master.rows.length} produk, ${master.series.length} kode seri`);
master.warnings.forEach((w) => console.log(`  ! ${w}`));

const plName = newestPriceList(ASSETS);
console.log(`Price list  : ${plName ?? '(tidak ada)'}\n`);

const { catalog: baseCatalog } = buildCatalogFromRows(master.rows, {
  origin: 'manual',
  sourceFile: MASTER_FILENAME,
});

let catalog = baseCatalog;
let merge = null;
const newPn90 = new Set();

if (plName) {
  const pl = await readPriceList(fs.readFileSync(path.join(ASSETS, plName)), plName);
  const series = new Map(master.series.map((s) => [normalizeLabel(s.code), s.marketingName]));
  const incoming = [];
  const deriveWarnings = [];
  for (const m of pl.models) {
    if (!m.line) continue;
    const { row, warnings } = deriveForWebRow(m.record, { mktNameRule: series });
    deriveWarnings.push(...warnings);
    incoming.push({
      pn90: m.pn90,
      row,
      provenance: { origin: 'priceList', sourceFile: plName, sheetName: m.sheetName, group: m.group },
    });
  }
  const merged = mergePriceList(catalog, incoming);
  catalog = merged.catalog;
  merge = merged.report;
  merged.report.added.forEach((e) => newPn90.add(e.pn90));

  console.log(`Model terbaca dari price list : ${incoming.length}`);
  console.log(`  produk baru ditambahkan     : ${merge.added.length}`);
  console.log(`  sudah ada, tidak disentuh   : ${merge.keptUntouched}`);
  console.log(`  ada di katalog, tidak di PL : ${merge.notInSource}  <- dipertahankan`);
  if (deriveWarnings.length) console.log(`  catatan derivasi            : ${deriveWarnings.length}`);
}

const rows = catalogToRows(catalog);
const csv = serializeForWebCsv(rows);
const rebuilt = await buildMasterWorkbook({ rows, series: master.series, newPn90 });
const rebuiltRead = await readMasterWorkbook(rebuilt);

console.log(`\nHasil akhir: ${rows.length} produk`);
console.log(`  CSV         : ${Buffer.byteLength(csv, 'utf8').toLocaleString('id-ID')} byte`);
console.log(`  File master : ${Buffer.byteLength(Buffer.from(rebuilt)).toLocaleString('id-ID')} byte`);

let failures = 0;
const assert = (ok, msg) => {
  console.log(`  ${ok ? 'ok  ' : 'FAIL'}  ${msg}`);
  if (!ok) failures += 1;
};

console.log('\nPemeriksaan:');
assert(rows.length >= master.rows.length, `jumlah produk tidak berkurang (${master.rows.length} -> ${rows.length})`);
assert(csv.startsWith('﻿'), 'CSV diawali BOM');
assert(csv.split('\r\n')[0].split(',').length === 25, 'header CSV 25 kolom');
assert(new Set(rows.map((r) => r.partNumber)).size === rows.length, '90PN unik di seluruh baris');
assert(rows.every((r) => !isBlankValue(r.productName)), 'setiap baris punya Product Name');
assert(rebuiltRead.rows.length === rows.length, 'file master hasil regenerasi bisa dibaca kembali utuh');
assert(rebuiltRead.unknownColumns.length === 0 && rebuiltRead.missingColumns.length === 0, 'kolom file master tetap lengkap');

const noCopy = rows.filter((r) => isBlankValue(r.tagline) && isBlankValue(r.ksp) && isBlankValue(r.link));
const noLink = rows.filter((r) => isBlankValue(r.link));
console.log(`\n  ${noCopy.length} produk tanpa Tagline/KSP/Link sama sekali, ${noLink.length} tanpa Link`);

if (merge && merge.added.length) {
  console.log('\nContoh produk baru:');
  merge.added.slice(0, 5).forEach((e) => {
    console.log(`  ${e.pn90}  ${e.row.productName}`);
    console.log(`    ${e.row.title}`);
  });
}

console.log('');
if (failures) { console.error(`GAGAL — ${failures} pemeriksaan tidak lulus.`); process.exit(1); }
console.log('LULUS — alur bulanan berjalan utuh dari file master sampai CSV.');
