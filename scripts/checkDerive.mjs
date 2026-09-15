/**
 * Menguji aturan derivasi terhadap kebenaran lapangan.
 *
 * 155 produk ada di dua tempat sekaligus: di ABP Price List (sebagai datasheet PM mentah) dan di
 * CSV yang sedang tayang (sebagai teks hasil olahan sheet Auto-*). Menderivasi yang mentah lalu
 * membandingkannya dengan yang tayang adalah cara langsung untuk mengukur seberapa setia
 * terjemahan formula Excel ke TypeScript.
 *
 * Dijalankan sebagai laporan, bukan gerbang lulus/gagal: beberapa selisih memang diharapkan
 * karena datasheet PM Januari 2026 lebih baru daripada sebagian baris yang tayang.
 */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, loadModule } from './_bundle.mjs';
import { newestPriceList, legacySourceWorkbook } from './_files.mjs';

const ASSETS = path.join(ROOT, 'assets');
const LIVE = path.join(ROOT, 'legacy-site/.original/A+ Content Source(ForWeb).csv');

const { parseForWebCsv } = await loadModule('src/lib/legacyCsv.ts');
const { readPriceList } = await loadModule('src/lib/priceListImport.ts');
const { readLegacyWorkbook } = await loadModule('src/lib/legacyCopyImport.ts');
const { deriveForWebRow } = await loadModule('src/lib/deriveForWeb.ts');
const { FORWEB_COLUMNS } = await loadModule('src/lib/config.ts');

const priceListName = newestPriceList(ASSETS);
const legacyFile = legacySourceWorkbook(ASSETS);

if (!priceListName) { console.error('Tidak ada file price list (*.xlsx) di assets/'); process.exit(1); }
if (!legacyFile) { console.error('Tidak ada file "A+ Content Source*.xlsx" di assets/'); process.exit(1); }

console.log(`Price list : ${priceListName}`);
console.log(`Workbook   : ${legacyFile}\n`);

const legacy = await readLegacyWorkbook(fs.readFileSync(path.join(ASSETS, legacyFile)), legacyFile);
console.log(`MKT name rule : ${legacy.stats.mktRules} aturan`);
console.log(`RAW.KSP.LInk  : ${legacy.stats.rawRows} baris, ${legacy.stats.withCopy} punya Tagline/KSP/Link`);
legacy.warnings.slice(0, 5).forEach((w) => console.log(`  ! ${w}`));

const pl = await readPriceList(fs.readFileSync(path.join(ASSETS, priceListName)), priceListName);
console.log(`\nDatasheet ditemukan: ${pl.sheets.length}`);
pl.sheets.forEach((s) => console.log(`  ${String(s.models).padStart(4)} model  ${String(s.labels).padStart(3)} label  ${s.group.padEnd(4)}  ${s.name}`));
console.log(`Total model: ${pl.models.length}`);
pl.warnings.slice(0, 5).forEach((w) => console.log(`  ! ${w}`));

const live = parseForWebCsv(fs.readFileSync(LIVE, 'utf8'));
const liveByPn = new Map(live.rows.map((r) => [(r.partNumber || '').trim(), r]));

const compared = [];
for (const m of pl.models) {
  const expected = liveByPn.get(m.pn90);
  if (!expected) continue;
  const { row } = deriveForWebRow(m.record, { mktNameRule: legacy.mktNameRule });
  compared.push({ pn90: m.pn90, modelName: m.modelName, actual: row, expected });
}

console.log(`\nProduk yang ada di price list DAN sedang tayang: ${compared.length}\n`);

// Tagline/KSP/Link tidak berasal dari price list, jadi tidak dibandingkan di sini.
const SKIP = new Set(['tagline', 'ksp', 'link']);
const cols = FORWEB_COLUMNS.filter((c) => !SKIP.has(c.key));

/**
 * Excel error text that leaked into the published CSV.
 *
 * 73 of the 392 live products carry one. Counting those as "our derivation is wrong" measures the
 * wrong thing entirely — the live value is the broken one — so they are reported separately and
 * kept out of the accuracy figure.
 */
const EXCEL_ERROR = /#(VALUE|REF|N\/A|NAME|DIV\/0|NUM|NULL)[!?]?/;

/**
 * Line endings are not a difference worth counting. The published CSV mixes CRLF and LF inside
 * cells; the page splits multi-line values with /\r?\n/, so both render identically. Comparing
 * them raw put the Office column at 18.6% when every one of those values was in fact correct.
 */
const sameText = (a, b) => a.replace(/\r\n/g, '\n') === b.replace(/\r\n/g, '\n');

const stats = new Map(cols.map((c) => [c.key, { match: 0, diff: 0, broken: 0, samples: [] }]));

for (const c of compared) {
  for (const col of cols) {
    const a = String(c.actual[col.key] ?? '').trim();
    const e = String(c.expected[col.key] ?? '').trim();
    const s = stats.get(col.key);
    if (EXCEL_ERROR.test(e)) { s.broken += 1; continue; }
    if (sameText(a, e)) s.match += 1;
    else {
      s.diff += 1;
      if (s.samples.length < 2) s.samples.push({ pn: c.pn90, model: c.modelName, e, a });
    }
  }
}

console.log('Kolom                    cocok   beda  live rusak   akurasi');
console.log('----------------------------------------------------------');
let totalMatch = 0;
let totalAll = 0;
let totalBroken = 0;
for (const col of cols) {
  const s = stats.get(col.key);
  const tot = s.match + s.diff;
  totalMatch += s.match;
  totalAll += tot;
  totalBroken += s.broken;
  const pct = tot ? ((s.match / tot) * 100).toFixed(1) : '0.0';
  const flag = s.diff === 0 ? '' : '  <--';
  console.log(
    `${col.header.padEnd(24)} ${String(s.match).padStart(5)}  ${String(s.diff).padStart(5)}  ${String(s.broken).padStart(10)}   ${pct.padStart(6)}%${flag}`,
  );
}
console.log('----------------------------------------------------------');
console.log(
  `${'TOTAL'.padEnd(24)} ${String(totalMatch).padStart(5)}  ${String(totalAll - totalMatch).padStart(5)}  ${String(totalBroken).padStart(10)}   ${((totalMatch / totalAll) * 100).toFixed(1).padStart(6)}%`,
);
console.log(`\n${totalBroken} sel dikeluarkan dari hitungan karena nilai yang TAYANG berupa error Excel.`);

// Sebaran di bawah ini yang menjawab pertanyaan sebenarnya: kalau selisihnya menumpuk di
// segelintir produk, berarti data price list Januari 2026 memang lebih baru daripada baris yang
// tayang. Kalau merata di semua produk, berarti ada aturan derivasi yang salah.
const perLine = {};
const dist = {};
for (const c of compared) {
  let d = 0;
  for (const col of cols) {
    const av = String(c.actual[col.key] ?? '').trim();
    const ev = String(c.expected[col.key] ?? '').trim();
    if (!EXCEL_ERROR.test(ev) && !sameText(av, ev)) d += 1;
  }
  const ln = c.actual.partNumber.slice(2, 4).toUpperCase();
  perLine[ln] = perLine[ln] ?? { n: 0, diffs: 0, perfect: 0 };
  perLine[ln].n += 1;
  perLine[ln].diffs += d;
  if (d === 0) perLine[ln].perfect += 1;
  dist[d] = (dist[d] ?? 0) + 1;
}

console.log('\n=== Per product line ===');
console.log('line   produk   rata2 kolom beda   cocok sempurna');
for (const [ln, s] of Object.entries(perLine)) {
  console.log(
    `  ${ln}   ${String(s.n).padStart(5)}   ${(s.diffs / s.n).toFixed(1).padStart(12)}   ${String(s.perfect).padStart(13)}`,
  );
}

console.log(`\n=== Sebaran jumlah kolom berbeda per produk (dari ${cols.length} kolom) ===`);
Object.keys(dist)
  .map(Number)
  .sort((a, b) => a - b)
  .forEach((k) => console.log(`  ${String(k).padStart(2)} kolom beda : ${String(dist[k]).padStart(4)} produk`));

console.log('\n=== Contoh selisih per kolom ===');
for (const col of cols) {
  const s = stats.get(col.key);
  if (!s.diff) continue;
  console.log(`\n--- ${col.header} (${s.diff} beda)`);
  for (const ex of s.samples) {
    console.log(`  ${ex.pn} ${ex.model}`);
    console.log(`    tayang : ${JSON.stringify(ex.e.slice(0, 110))}`);
    console.log(`    derivasi: ${JSON.stringify(ex.a.slice(0, 110))}`);
  }
}
