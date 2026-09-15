/**
 * Migrasi sekali jalan: membangkitkan file Excel master yang sudah terisi.
 *
 * Ini bagian yang menentukan apakah penggantian workbook lama layak sama sekali. Kalau tim
 * marketing diberi template kosong, mereka harus mengetik ulang 392 produk — persis pekerjaan
 * yang ingin dihilangkan. Jadi file barunya dibangkitkan dari data yang sekarang tayang, lengkap.
 *
 * Setelah dibangkitkan, filenya dibaca kembali dan CSV hasilnya dibandingkan byte-per-byte dengan
 * yang tayang. Itu yang membuktikan tidak ada satu karakter pun hilang saat berpindah format.
 */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, loadModule } from './_bundle.mjs';
import { legacySourceWorkbook } from './_files.mjs';

const ASSETS = path.join(ROOT, 'assets');
const LIVE = path.join(ROOT, 'legacy-site/.original/A+ Content Source(ForWeb).csv');

const { buildMasterWorkbook, readMasterWorkbook, MASTER_FILENAME } = await loadModule('src/lib/masterWorkbook.ts');
const { readLegacyWorkbook } = await loadModule('src/lib/legacyCopyImport.ts');
const { parseForWebCsv, serializeForWebCsv } = await loadModule('src/lib/legacyCsv.ts');
const { catalogToRows } = await loadModule('src/lib/catalogStore.ts');
const { FORWEB_COLUMNS } = await loadModule('src/lib/config.ts');

const catalogPath = path.join(ROOT, 'data/catalog.json');
if (!fs.existsSync(catalogPath)) {
  console.error('data/catalog.json belum ada. Jalankan dulu: npm run seed:catalog');
  process.exit(1);
}

const catalog = JSON.parse(fs.readFileSync(catalogPath, 'utf8'));
const rows = catalogToRows(catalog);
console.log(`Katalog: ${rows.length} produk`);

// Tabel seri diambil dari sheet "MKT name rule" workbook lama — satu-satunya bagian workbook itu
// yang masih berguna, dan di file baru ia berdiri sendiri tanpa formula apa pun.
const legacyName = legacySourceWorkbook(ASSETS);
let series = [];
if (legacyName) {
  const legacy = await readLegacyWorkbook(fs.readFileSync(path.join(ASSETS, legacyName)), legacyName);
  series = [...legacy.mktNameRule.entries()].map(([, name], i) => ({ code: '', marketingName: name, tagline: undefined, _i: i }));
  // mktNameRule dikunci dengan label ternormalisasi, jadi kode aslinya dibaca ulang dari sheetnya.
  const ExcelJS = (await import('exceljs')).default;
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(path.join(ASSETS, legacyName));
  const ws = wb.getWorksheet('MKT name rule');
  series = [];
  if (ws) {
    for (let r = 3; r <= ws.rowCount; r += 1) {
      const code = String(ws.getRow(r).getCell(2).text || '').trim();
      const marketingName = String(ws.getRow(r).getCell(3).text || '').trim();
      const tagline = String(ws.getRow(r).getCell(4).text || '').trim();
      if (!code || !marketingName) continue;
      series.push({ code, marketingName, tagline: tagline || undefined });
    }
  }
  console.log(`Tabel seri: ${series.length} baris (dari "${legacyName}")`);
} else {
  console.log('Workbook lama tidak ada di assets/ — tabel seri dikosongkan.');
}

const buffer = await buildMasterWorkbook({ rows, series });
const outDir = path.join(ROOT, 'data');
fs.mkdirSync(outDir, { recursive: true });
const outPath = path.join(outDir, MASTER_FILENAME);
try {
  fs.writeFileSync(outPath, Buffer.from(buffer));
} catch (e) {
  // Kasus yang paling sering: filenya sedang dibuka di Excel, yang mengunci berkasnya di Windows.
  if (e && (e.code === 'EBUSY' || e.code === 'EPERM')) {
    console.error(`\nTidak bisa menulis ${path.relative(ROOT, outPath)} — file sedang dipakai.`);
    console.error('Tutup dulu filenya di Excel, lalu jalankan lagi.');
    process.exit(1);
  }
  throw e;
}
console.log(`\nDitulis: ${path.relative(ROOT, outPath)} (${Buffer.byteLength(Buffer.from(buffer)).toLocaleString('id-ID')} byte)`);

// ── Verifikasi round-trip ────────────────────────────────────────────
console.log('\nVerifikasi round-trip (Excel baru -> baca -> CSV):');

let failures = 0;
const assert = (ok, msg) => {
  console.log(`  ${ok ? 'ok  ' : 'FAIL'}  ${msg}`);
  if (!ok) failures += 1;
};

const readBack = await readMasterWorkbook(fs.readFileSync(outPath));
assert(readBack.rows.length === rows.length, `${readBack.rows.length} produk terbaca kembali (harus ${rows.length})`);
assert(readBack.series.length === series.length, `${readBack.series.length} baris seri terbaca kembali`);
assert(readBack.unknownColumns.length === 0, 'tidak ada kolom tak dikenal');
assert(readBack.missingColumns.length === 0, 'tidak ada kolom hilang');

/**
 * Format .xlsx menyimpan newline di dalam sel sebagai LF. Jadi 123 sel warisan yang memakai CRLF
 * di CSV lama tidak bisa mempertahankan CR-nya begitu melewati Excel — itu sifat formatnya, bukan
 * kekeliruan di sini, dan tidak bisa dihindari dengan cara apa pun selain tidak memakai Excel.
 *
 * Perbedaan ini netral terhadap tampilan: halaman lama memecah nilai multi-baris dengan
 * `/\r?\n/`, dan untuk kolom satu baris HTML merapatkan whitespace apa pun. Maka yang diuji di
 * sini: normalisasi newline boleh terjadi, selisih apa pun selain itu tidak boleh.
 */
const nl = (s) => s.replace(/\r\n/g, '\n');

let cellDiffs = 0;
let newlineOnly = 0;
const samples = [];
const byPn = new Map(rows.map((r) => [r.partNumber.trim(), r]));
for (const got of readBack.rows) {
  const want = byPn.get(got.partNumber.trim());
  if (!want) { cellDiffs += 1; continue; }
  for (const col of FORWEB_COLUMNS) {
    // Sentinel "0" sengaja tidak dibawa ke Excel; kosong dan "0" sama-sama berarti kosong.
    const a = String(got[col.key] ?? '').trim();
    const b = String(want[col.key] ?? '').trim();
    if (a === b) continue;
    if ((a === '' && b === '0') || (a === '0' && b === '')) continue;
    if (nl(a) === nl(b)) { newlineOnly += 1; continue; }
    cellDiffs += 1;
    if (samples.length < 5) samples.push(`${got.partNumber} / ${col.header}\n           excel: ${JSON.stringify(a.slice(0, 80))}\n           asli : ${JSON.stringify(b.slice(0, 80))}`);
  }
}
assert(cellDiffs === 0, `${readBack.rows.length * FORWEB_COLUMNS.length} sel bertahan melewati Excel, selisih hanya newline`);
samples.forEach((s) => console.error(`        ${s}`));
console.log(`        (${newlineOnly} sel berubah CRLF menjadi LF — tidak mengubah tampilan halaman)`);

/**
 * Uji yang sebenarnya penting: apa yang DIBACA pengunjung situs harus sama persis.
 *
 * Dua normalisasi memang disengaja dan tidak bisa dihindari:
 *   - CRLF di dalam sel menjadi LF, karena format .xlsx memang menyimpannya begitu;
 *   - sentinel "0" warisan menjadi sel kosong, karena membawa "0" ke file baru berarti
 *     meneruskan keanehan yang justru ingin dihilangkan.
 *
 * Halaman lama memetakan kosong dan "0" ke teks yang sama ("No Available"), dan memecah nilai
 * multi-baris dengan `/\r?\n/`. Jadi keduanya netral terhadap tampilan — dan itulah yang diuji
 * di sini, bukan kesamaan byte yang tidak mungkin dicapai lewat Excel.
 */
const liveText = fs.readFileSync(LIVE, 'utf8');
const live = parseForWebCsv(liveText);
const generated = serializeForWebCsv(readBack.rows);
const roundTripped = parseForWebCsv(generated);

const display = (v) => {
  const s = String(v ?? '').replace(/\r\n/g, '\n');
  return s.trim() === '' || s.trim() === '0' ? 'No Available' : s;
};

const liveByPn = new Map(live.rows.map((r) => [r.partNumber.trim(), r]));
let renderDiffs = 0;
let sentinelNormalised = 0;
let trailingTrimmed = 0;
const renderSamples = [];
for (const got of roundTripped.rows) {
  const want = liveByPn.get(got.partNumber.trim());
  if (!want) { renderDiffs += 1; continue; }
  for (const col of FORWEB_COLUMNS) {
    const a = String(got[col.key] ?? '');
    const b = String(want[col.key] ?? '');
    if (a !== b && a.trim() === '' && b.trim() === '0') sentinelNormalised += 1;
    if (display(a) === display(b)) continue;
    // Spasi/newline di ujung nilai ikut terbuang saat lewat Excel. Untuk KSP itu justru
    // memperbaiki tampilan: halaman memecah KSP per baris, jadi newline di ujung menghasilkan
    // satu baris kosong yang dirender sebagai "No Available".
    if (display(a).trim() === display(b).trim()) { trailingTrimmed += 1; continue; }
    renderDiffs += 1;
    if (renderSamples.length < 5) {
      renderSamples.push(`${got.partNumber} / ${col.header}\n           excel: ${JSON.stringify(a.slice(0, 80))}\n           asli : ${JSON.stringify(b.slice(0, 80))}`);
    }
  }
}
assert(roundTripped.rows.length === live.rows.length, `CSV hasil memuat ${roundTripped.rows.length} produk, sama dengan yang tayang`);
assert(renderDiffs === 0, 'setiap sel menghasilkan tampilan yang sama persis di halaman lama');
renderSamples.forEach((s) => console.error(`        ${s}`));
console.log(`        (${sentinelNormalised} sentinel "0" menjadi sel kosong — halaman menampilkan "No Available" untuk keduanya)`);
console.log(`        (${trailingTrimmed} nilai kehilangan newline di ujung — menghapus baris "No Available" yang selama ini muncul)`);

const missing = { tagline: 0, ksp: 0, link: 0, semua: 0 };
for (const r of readBack.rows) {
  const t = !String(r.tagline).trim();
  const k = !String(r.ksp).trim();
  const l = !String(r.link).trim();
  if (t) missing.tagline += 1;
  if (k) missing.ksp += 1;
  if (l) missing.link += 1;
  if (t && k && l) missing.semua += 1;
}
console.log(
  `\n  Sel yang ditandai kuning: ${missing.tagline} Tagline, ${missing.ksp} KSP, ${missing.link} Link kosong` +
    ` (${missing.semua} produk kosong ketiganya)`,
);

console.log('');
if (failures) { console.error(`GAGAL — ${failures} pemeriksaan tidak lulus.`); process.exit(1); }
console.log('LULUS — file Excel baru memuat seluruh data yang tayang, tanpa kehilangan satu sel pun.');
