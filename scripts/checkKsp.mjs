/**
 * Memeriksa mutu draf KSP terhadap seluruh katalog sungguhan.
 *
 * Draf KSP ditulis ke kolom yang tayang di website, jadi isinya tidak boleh sekadar "ada".
 * Yang dijaga di sini adalah hal-hal yang pernah benar-benar terjadi pada data ASUS: teks error
 * Excel (#VALUE!) yang ikut ter-export sebagai nilai biasa, dan nilai yang nyasar ke kolom lain
 * (teks memori di kolom GPU). Keduanya lolos kalau tidak diperiksa, dan hasilnya jadi kalimat
 * jualan yang salah.
 */
import fs from 'node:fs';
import path from 'node:path';
import { ROOT, loadModule } from './_bundle.mjs';
import { newestPriceList } from './_files.mjs';

const ASSETS = path.join(ROOT, 'assets');

const { readPriceList } = await loadModule('src/lib/priceListImport.ts');
const { readMasterWorkbook, MASTER_FILENAME } = await loadModule('src/lib/masterWorkbook.ts');
const { deriveForWebRow } = await loadModule('src/lib/deriveForWeb.ts');
const { buildCatalogFromRows, mergePriceList, catalogToRows } = await loadModule('src/lib/catalogStore.ts');
const { draftKsp, draftableRows, KSP_BULLET_TARGET } = await loadModule('src/lib/buildKsp.ts');
const { normalizeLabel } = await loadModule('src/vendor/pmCell.ts');

const masterPath = path.join(ROOT, 'data', MASTER_FILENAME);
if (!fs.existsSync(masterPath)) {
  console.error(`${MASTER_FILENAME} belum ada. Jalankan: npm run build:master`);
  process.exit(1);
}

const master = await readMasterWorkbook(fs.readFileSync(masterPath), MASTER_FILENAME);
let { catalog } = buildCatalogFromRows(master.rows, { origin: 'manual', sourceFile: MASTER_FILENAME });

const plName = newestPriceList(ASSETS);
if (plName) {
  const pl = await readPriceList(fs.readFileSync(path.join(ASSETS, plName)), plName);
  const series = new Map(master.series.map((s) => [normalizeLabel(s.code), s.marketingName]));
  const incoming = [];
  for (const m of pl.models) {
    if (!m.line) continue;
    const { row } = deriveForWebRow(m.record, { mktNameRule: series });
    incoming.push({
      pn90: m.pn90,
      row,
      provenance: { origin: 'priceList', sourceFile: plName, sheetName: m.sheetName, group: m.group },
    });
  }
  catalog = mergePriceList(catalog, incoming).catalog;
}

const rows = catalogToRows(catalog);
const targets = draftableRows(rows);
console.log(`Katalog     : ${rows.length} produk (${plName ?? 'tanpa price list'})`);
console.log(`KSP kosong & bisa didraf : ${targets.length}\n`);

const BAD = [
  ['teks error Excel', /#(VALUE|REF|NAME|NUM|NULL|DIV\/0|N\/A)/i],
  ['penanda kosong "Not Found"', /not\s*found/i],
  ['teks memori di posisi GPU', /with \d+\s*GB (DDR|LPDDR)/i],
  ['kurung kosong', /\(\s*\)/],
  ['bullet kosong', /^- *$/m],
];

const problems = [];
const lengths = [];
let threeBullets = 0;

for (const row of targets) {
  const ksp = draftKsp(row);
  const bullets = ksp.split('\n');
  if (bullets.length === KSP_BULLET_TARGET) threeBullets += 1;
  for (const b of bullets) lengths.push(b.length);
  for (const [label, re] of BAD) {
    if (re.test(ksp)) problems.push({ pn: row.partNumber, label, ksp });
  }
}

lengths.sort((a, b) => a - b);
const at = (p) => lengths[Math.min(lengths.length - 1, Math.floor(lengths.length * p))] ?? 0;
console.log(`Panjang bullet : median ${at(0.5)} / p90 ${at(0.9)} / maks ${lengths[lengths.length - 1] ?? 0} karakter`);
console.log(`Bullet lengkap : ${threeBullets} dari ${targets.length} dapat ${KSP_BULLET_TARGET} bullet\n`);

if (problems.length) {
  console.log(`GAGAL — ${problems.length} draf memuat nilai yang tidak layak tayang:\n`);
  for (const p of problems.slice(0, 10)) {
    console.log(`  ${p.pn}  (${p.label})`);
    console.log(`${p.ksp.split('\n').map((l) => '      ' + l).join('\n')}\n`);
  }
  if (problems.length > 10) console.log(`  ... dan ${problems.length - 10} lainnya`);
  process.exit(1);
}

const sample = targets[0];
if (sample) {
  console.log('Contoh draf:');
  console.log(`  ${sample.partNumber}  ${sample.marketingName}`);
  console.log(draftKsp(sample).split('\n').map((l) => '    ' + l).join('\n'));
  console.log();
}
console.log('LULUS — tidak ada draf KSP yang memuat teks error atau nilai nyasar.');
