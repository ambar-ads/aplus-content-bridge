/**
 * Mengisi data/catalog.json dari CSV yang SEDANG tayang.
 *
 * Ini langkah sekali jalan yang menjadi fondasi seluruh pendekatan. ABP Price List hanya
 * memuat produk yang sedang dijual bulan itu (155 dari 392 produk yang tayang), sementara
 * `RAW.KSP.LInk` di workbook lama bersifat akumulatif sejak Agustus 2025. Membangun ulang CSV
 * dari price list saja akan menghapus ratusan produk dari situs. Jadi katalog di-seed dulu dari
 * apa yang sudah terbit, lalu price list berikutnya hanya menambahi.
 */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, loadModule } from './_bundle.mjs';

const SOURCE = process.argv[2]
  ? path.resolve(process.argv[2])
  : path.join(ROOT, 'legacy-site/.original/A+ Content Source(ForWeb).csv');

const { parseForWebCsv } = await loadModule('src/lib/legacyCsv.ts');
const { buildCatalogFromRows } = await loadModule('src/lib/catalogStore.ts');

if (!fs.existsSync(SOURCE)) {
  console.error(`CSV sumber tidak ditemukan: ${SOURCE}`);
  process.exit(1);
}

const text = fs.readFileSync(SOURCE, 'utf8');
const { rows, ghostRows, headerMatches } = parseForWebCsv(text);

console.log(`Sumber      : ${path.relative(ROOT, SOURCE)}`);
console.log(`Header cocok: ${headerMatches ? 'ya' : 'TIDAK — periksa kontrak kolom'}`);
console.log(`Produk      : ${rows.length}`);
console.log(`Baris hantu : ${ghostRows} (dibuang)`);

const { catalog, warnings } = buildCatalogFromRows(rows, {
  origin: 'seed',
  sourceFile: path.basename(SOURCE),
});

if (warnings.length) {
  console.log(`\nCatatan kualitas data (${warnings.length}):`);
  warnings.forEach((w) => console.log(`  - ${w}`));
}

const out = path.join(ROOT, 'data/catalog.json');
fs.mkdirSync(path.dirname(out), { recursive: true });
fs.writeFileSync(out, JSON.stringify(catalog, null, 2) + '\n');
console.log(`\nDitulis: ${path.relative(ROOT, out)} (${catalog.entries.length} entri)`);
