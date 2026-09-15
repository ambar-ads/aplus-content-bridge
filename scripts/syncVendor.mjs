/**
 * Menyalin potongan "parsing contract" dari repo APLUS utama ke proyek ini.
 *
 * Proyek ini sengaja berdiri sendiri, jadi kode dari repo utama disalin, bukan di-import.
 * Konsekuensinya salinan itu bisa basi. Script ini merekam sha256 setiap file sumber ke
 * vendor-upstream/manifest.json; `npm run check:vendor` membandingkan ulang dan memberi
 * peringatan begitu repo utama berubah, supaya penyimpangan ketahuan, bukan terjadi diam-diam.
 *
 * Repo utama hanya DIBACA. Tidak ada satu pun file di sana yang disentuh.
 */
import { createHash } from 'node:crypto';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MAIN_REPO = process.env.APLUS_MAIN_REPO
  ? path.resolve(process.env.APLUS_MAIN_REPO)
  : path.resolve(ROOT, '..', 'asus-aplus-2-0');

const UPSTREAM_DIR = path.join(ROOT, 'vendor-upstream');
const MANIFEST = path.join(UPSTREAM_DIR, 'manifest.json');

/** File yang disalin verbatim untuk keperluan diff & hashing. */
const FILES = [
  'src/config/pmDatasheetLayout.ts',
  'src/config/specSchema.ts',
  'src/utils/specSource.ts',
  'src/utils/pmDatasheetImport.ts',
];

const sha256 = (buf) => createHash('sha256').update(buf).digest('hex');
const checkOnly = process.argv.includes('--check');

if (!fs.existsSync(MAIN_REPO)) {
  console.error(`Repo utama tidak ditemukan: ${MAIN_REPO}`);
  console.error('Set APLUS_MAIN_REPO kalau lokasinya berbeda.');
  process.exit(1);
}

const current = {};
const missing = [];
for (const rel of FILES) {
  const abs = path.join(MAIN_REPO, rel);
  if (!fs.existsSync(abs)) { missing.push(rel); continue; }
  current[rel] = sha256(fs.readFileSync(abs));
}

if (missing.length) {
  console.error('File sumber hilang di repo utama:');
  missing.forEach((m) => console.error(`  ${m}`));
  process.exit(1);
}

if (checkOnly) {
  if (!fs.existsSync(MANIFEST)) {
    console.error('Belum ada vendor-upstream/manifest.json. Jalankan: npm run sync:vendor');
    process.exit(1);
  }
  const prev = JSON.parse(fs.readFileSync(MANIFEST, 'utf8'));
  const drifted = FILES.filter((rel) => prev.files[rel] !== current[rel]);
  if (drifted.length) {
    console.error('Vendor sudah menyimpang dari repo utama:');
    drifted.forEach((d) => console.error(`  ${d}`));
    console.error('\nTinjau perubahannya, lalu jalankan: npm run sync:vendor');
    console.error(`Diff: diff vendor-upstream/${path.basename(drifted[0])} "${path.join(MAIN_REPO, drifted[0])}"`);
    process.exit(1);
  }
  console.log(`Vendor sinkron dengan repo utama (${FILES.length} file, disalin ${prev.syncedAt}).`);
  process.exit(0);
}

fs.mkdirSync(UPSTREAM_DIR, { recursive: true });
for (const rel of FILES) {
  fs.copyFileSync(path.join(MAIN_REPO, rel), path.join(UPSTREAM_DIR, path.basename(rel)));
}

// pmDatasheetLayout adalah satu-satunya yang ikut dikompilasi. Satu-satunya yang perlu diubah
// adalah import tipe lintas-repo; sisanya disalin apa adanya supaya diff-nya tetap bermakna.
const layoutSrc = fs.readFileSync(path.join(MAIN_REPO, 'src/config/pmDatasheetLayout.ts'), 'utf8');
const TYPE_IMPORT = "import type { ProductLineCode } from '@/types/specForm';";
if (!layoutSrc.includes(TYPE_IMPORT)) {
  console.error('Bentuk import di pmDatasheetLayout.ts berubah — sesuaikan syncVendor.mjs.');
  process.exit(1);
}
const adapted = layoutSrc.replace(
  TYPE_IMPORT,
  "import type { ProductLineCode } from './types';",
);
fs.mkdirSync(path.join(ROOT, 'src/vendor'), { recursive: true });
fs.writeFileSync(
  path.join(ROOT, 'src/vendor/pmLayout.ts'),
  `// DISALIN dari repo utama oleh scripts/syncVendor.mjs — jangan diedit langsung.\n` +
    `// Sumber: ${FILES[0]}\n` +
    `// Satu-satunya perubahan: import tipe diarahkan ke ./types.\n` +
    adapted,
);

fs.writeFileSync(
  MANIFEST,
  JSON.stringify({ mainRepo: MAIN_REPO, syncedAt: new Date().toISOString(), files: current }, null, 2) + '\n',
);

console.log(`Vendor tersinkron dari ${MAIN_REPO}`);
FILES.forEach((rel) => console.log(`  ${rel}  ${current[rel].slice(0, 12)}`));
console.log('\nDihasilkan: src/vendor/pmLayout.ts');
