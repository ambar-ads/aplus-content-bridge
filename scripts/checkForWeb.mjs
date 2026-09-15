/**
 * Membuktikan bahwa CSV yang dihasilkan katalog identik dengan yang SEDANG tayang.
 *
 * Ini pengaman utama proyek ini. Teks 392 produk itu sudah dipakai marketing untuk listing
 * e-commerce; refactor apa pun di sisi kita tidak boleh mengubahnya diam-diam. Perbedaan yang
 * memang diharapkan hanya satu: 1611 baris hantu dari spill range Excel yang dibuang.
 *
 * Exit non-zero begitu ada satu sel pun yang menyimpang.
 */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, loadModule } from './_bundle.mjs';

const LIVE = path.join(ROOT, 'legacy-site/.original/A+ Content Source(ForWeb).csv');
const CATALOG = path.join(ROOT, 'data/catalog.json');

const { parseForWebCsv, serializeForWebCsv } = await loadModule('src/lib/legacyCsv.ts');
const { catalogToRows } = await loadModule('src/lib/catalogStore.ts');
const { FORWEB_COLUMNS } = await loadModule('src/lib/config.ts');

for (const [label, file] of [['CSV live', LIVE], ['katalog', CATALOG]]) {
  if (!fs.existsSync(file)) {
    console.error(`${label} tidak ditemukan: ${path.relative(ROOT, file)}`);
    if (file === CATALOG) console.error('Jalankan dulu: npm run seed:catalog');
    process.exit(1);
  }
}

const liveText = fs.readFileSync(LIVE, 'utf8');
const catalog = JSON.parse(fs.readFileSync(CATALOG, 'utf8'));

const live = parseForWebCsv(liveText);
const actualRows = catalogToRows(catalog);

let failures = 0;
const fail = (msg) => { failures += 1; console.error(`  FAIL  ${msg}`); };
const pass = (msg) => console.log(`  ok    ${msg}`);

console.log('1. Kontrak kolom');
live.headerMatches ? pass('25 kolom header cocok persis (termasuk "Part Number" ganda)')
                   : fail('header CSV live tidak cocok dengan FORWEB_COLUMNS');

console.log('\n2. Jumlah baris');
if (actualRows.length === live.rows.length) pass(`${actualRows.length} produk, sama dengan CSV live`);
else fail(`katalog ${actualRows.length} produk vs CSV live ${live.rows.length}`);
pass(`${live.ghostRows} baris hantu dibuang (memang diharapkan)`);

console.log('\n3. Isi sel, dibanding per 90PN');
const liveByPn = new Map(live.rows.map((r) => [(r.partNumber || r.partNumberAuto || '').trim(), r]));
let cellDiffs = 0;
const shown = [];
for (const row of actualRows) {
  const pn = (row.partNumber || row.partNumberAuto || '').trim();
  const expected = liveByPn.get(pn);
  if (!expected) { fail(`90PN ${pn} tidak ada di CSV live`); continue; }
  for (const col of FORWEB_COLUMNS) {
    if ((row[col.key] ?? '') !== (expected[col.key] ?? '')) {
      cellDiffs += 1;
      if (shown.length < 10) {
        shown.push(`${pn} / ${col.header}\n         live: ${JSON.stringify(String(expected[col.key]).slice(0, 90))}\n         kita: ${JSON.stringify(String(row[col.key]).slice(0, 90))}`);
      }
    }
  }
}
if (cellDiffs === 0) pass(`${actualRows.length * FORWEB_COLUMNS.length} sel identik`);
else { fail(`${cellDiffs} sel berbeda`); shown.forEach((s) => console.error(`        ${s}`)); }

console.log('\n4. Round-trip serialize -> parse');
const generated = serializeForWebCsv(actualRows);
const reparsed = parseForWebCsv(generated);
if (reparsed.rows.length !== actualRows.length) {
  fail(`re-parse menghasilkan ${reparsed.rows.length} baris, bukan ${actualRows.length}`);
} else {
  let rt = 0;
  reparsed.rows.forEach((r, i) => {
    for (const col of FORWEB_COLUMNS) if ((r[col.key] ?? '') !== (actualRows[i][col.key] ?? '')) rt += 1;
  });
  rt === 0 ? pass('serialize lalu parse mengembalikan data yang sama persis')
           : fail(`${rt} sel berubah setelah round-trip`);
}

console.log('\n5. Bentuk byte');
generated.startsWith('﻿') ? pass('UTF-8 BOM ada') : fail('BOM hilang');
generated.endsWith('\r\n') ? pass('diakhiri CRLF') : fail('tidak diakhiri CRLF');

/** Pecah jadi record utuh, menghormati newline di dalam sel. */
function splitRecords(text) {
  const src = text.startsWith('﻿') ? text.slice(1) : text;
  const out = [];
  let cur = '';
  let q = false;
  for (let i = 0; i < src.length; i += 1) {
    const c = src[i];
    if (q) { cur += c; if (c === '"') { if (src[i + 1] === '"') { cur += '"'; i += 1; } else q = false; } continue; }
    if (c === '"') { q = true; cur += c; continue; }
    if (c === '\r' && src[i + 1] === '\n') { out.push(cur); cur = ''; i += 1; continue; }
    if (c === '\n') { out.push(cur); cur = ''; continue; }
    cur += c;
  }
  if (cur !== '') out.push(cur);
  return out;
}

const liveRecords = splitRecords(liveText);
const genRecords = splitRecords(generated);
// Baris hantu di file live: Product Name kosong atau "0".
const liveReal = liveRecords.filter((rec, i) => {
  if (i === 0) return true;
  const m = rec.split(',');
  const name = (m[1] ?? '').trim();
  return name !== '' && name !== '0';
});

console.log('\n6. Perbandingan byte terhadap record CSV live (setelah baris hantu dibuang)');
if (liveReal.length !== genRecords.length) {
  fail(`jumlah record ${genRecords.length} vs ${liveReal.length}`);
} else {
  let byteDiffs = 0;
  const firstDiffs = [];
  for (let i = 0; i < liveReal.length; i += 1) {
    if (liveReal[i] !== genRecords[i]) {
      byteDiffs += 1;
      if (firstDiffs.length < 3) {
        firstDiffs.push(`record #${i}\n         live: ${JSON.stringify(liveReal[i].slice(0, 120))}\n         kita: ${JSON.stringify(genRecords[i].slice(0, 120))}`);
      }
    }
  }
  if (byteDiffs === 0) pass(`${genRecords.length} record identik byte-per-byte (termasuk quoting)`);
  else { fail(`${byteDiffs} record berbeda secara byte`); firstDiffs.forEach((d) => console.error(`        ${d}`)); }
}

console.log('');
if (failures) {
  console.error(`GAGAL — ${failures} pemeriksaan tidak lulus.`);
  process.exit(1);
}
console.log('LULUS — CSV hasil generate identik dengan yang sedang tayang.');
