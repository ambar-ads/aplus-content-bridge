/**
 * Menerapkan perubahan pada salinan `index.html` situs lama.
 *
 * Sengaja dijadikan script, bukan hasil edit tangan, supaya perubahannya bisa dijelaskan,
 * diulang, dan dibandingkan dengan berkas asli di `legacy-site/.original/index.html`.
 *
 * Dijalankan dari versi ASLI setiap kali, jadi menjalankannya dua kali tidak menumpuk perubahan.
 */
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const ORIGINAL = path.join(ROOT, 'legacy-site/.original/index.html');
const TARGET = path.join(ROOT, 'legacy-site/index.html');

let html = fs.readFileSync(ORIGINAL, 'utf8');
const patches = [];

function replace(label, from, to) {
  if (!html.includes(from)) {
    console.error(`GAGAL — potongan untuk "${label}" tidak ditemukan di berkas asli.`);
    process.exit(1);
  }
  html = html.replace(from, to);
  patches.push(label);
}

// ── 1. Gaya untuk baris petunjuk jumlah hasil ────────────────────────
replace(
  'gaya petunjuk hasil',
  `    .toast { position:fixed; bottom:1rem; left:50%;`,
  `    .match-hint { text-align:center; color:#8a6d00; background:#fff8e5; border:1px solid #f0d98c; border-radius:4px; padding:.5rem .75rem; margin-bottom:1rem; font-size:.875rem; display:none; }
    .toast { position:fixed; bottom:1rem; left:50%;`,
);

// ── 2. Tempat menampilkan petunjuk itu ───────────────────────────────
replace(
  'elemen petunjuk hasil',
  `    <div id="results-wrapper" style="display:none;">
      <!-- Name Section -->`,
  `    <div id="results-wrapper" style="display:none;">
      <div class="match-hint" id="match-hint"></div>
      <!-- Name Section -->`,
);

// ── 3. Contoh pada kotak pencarian ───────────────────────────────────
replace(
  'placeholder kotak pencarian',
  `<input id="search" placeholder="e.g. B1402CVA-NK7110X">`,
  `<input id="search" placeholder="e.g. B1402CVA-NK7110X, 90NX04U1-M004P0, ExpertBook B1">`,
);

replace(
  'teks bantuan',
  `<div class="message" id="message">Type a name to search (e.g. B1402CVA-NK7110X)</div>`,
  `<div class="message" id="message">Search by model name, 90PN, or marketing name</div>`,
);

// ── 4. Ambil referensi elemen petunjuk ───────────────────────────────
replace(
  'referensi elemen',
  `    const toast = document.getElementById('toast');`,
  `    const toast = document.getElementById('toast');
    const matchHint = document.getElementById('match-hint');`,
);

// ── 5. Jangan hapus Part Number: sekarang ikut dicari ────────────────
replace(
  'pertahankan Part Number',
  `        data = d.map(row => { delete row['Part Number']; return row; });`,
  `        // Part Number dulu dibuang di sini. Sekarang dipertahankan karena 90PN ikut dicari;
        // kolom ini tetap tidak pernah ditampilkan di halaman.
        data = d;`,
);

// ── 6. Cari di tiga kolom, bukan satu ────────────────────────────────
replace(
  'pencarian tiga kolom',
  `      const row = data.find(r => (r['Product Name'] || '').toLowerCase().includes(term));
      if (!row) {
        messageEl.textContent = 'No product found';
        results.style.display = 'none';
        retryBtn.style.display = 'block';
        return;
      }`,
  `      // Dulu hanya cocok dengan Product Name. 90PN adalah kunci resmi produk dan sering itu
      // yang dipegang sales, sedangkan marketing name adalah yang diingat orang.
      const hits = data.filter(r =>
        (r['Product Name'] || '').toLowerCase().includes(term) ||
        (r['Part Number'] || '').toLowerCase().includes(term) ||
        (r['Marketing name'] || '').toLowerCase().includes(term)
      );
      const row = hits[0];
      if (!row) {
        messageEl.textContent = 'No product found';
        messageEl.style.display = 'block';
        results.style.display = 'none';
        retryBtn.style.display = 'block';
        return;
      }
      // Satu marketing name dipakai puluhan model, jadi tanpa keterangan ini pencarian seperti
      // "ExpertBook B1" akan terlihat memilih satu produk secara acak.
      if (hits.length > 1) {
        matchHint.textContent = hits.length + ' products match — showing the first. Type something more specific to narrow it down.';
        matchHint.style.display = 'block';
      } else {
        matchHint.style.display = 'none';
      }`,
);

// ── 7. Sembunyikan petunjuk saat direset ─────────────────────────────
replace(
  'reset petunjuk',
  `    function reset() {
      input.value = '';`,
  `    function reset() {
      input.value = '';
      matchHint.style.display = 'none';`,
);

replace(
  'teks bantuan saat reset',
  `      messageEl.textContent = 'Type a name to search (e.g. B1402CVA-NK7110X)';
      messageEl.style.display = 'block';`,
  `      messageEl.textContent = 'Search by model name, 90PN, or marketing name';
      messageEl.style.display = 'block';`,
);

// ── 8. Penanganan gagal memuat berkas ────────────────────────────────
replace(
  'penanganan gagal memuat',
  `    Papa.parse('A+ Content Source(ForWeb).csv', {
      download: true,
      header: true,
      skipEmptyLines: true,
      complete: ({ data: d }) => {`,
  `    function showLoadError(detail) {
      input.disabled = true;
      messageEl.style.display = 'block';
      messageEl.textContent =
        'Could not load the product data. Check that "A+ Content Source(ForWeb).csv" exists in the ' +
        'repository and that the filename matches exactly.' + (detail ? ' (' + detail + ')' : '');
    }

    Papa.parse('A+ Content Source(ForWeb).csv', {
      download: true,
      header: true,
      skipEmptyLines: true,
      error: err => showLoadError(err && err.message ? err.message : 'file could not be read'),
      complete: ({ data: d, meta }) => {
        // Respons 200 yang ternyata bukan CSV (halaman 404, atau .xlsx yang salah terunggah)
        // dulu ter-parse jadi data sampah dan halaman diam tanpa isi.
        if (!meta || !meta.fields || meta.fields.indexOf('Product Name') === -1) {
          return showLoadError('unrecognised file format');
        }`,
);

replace(
  'teks bantuan setelah termuat',
  `        input.disabled = false;
        messageEl.textContent = 'Type a name to search (e.g. B1402CVA-NK7110X)';`,
  `        input.disabled = false;
        messageEl.textContent = 'Search by model name, 90PN, or marketing name';`,
);

fs.writeFileSync(TARGET, html);

console.log(`Ditulis: ${path.relative(ROOT, TARGET)}`);
console.log(`${patches.length} perubahan diterapkan pada berkas asli:`);
patches.forEach((p, i) => console.log(`  ${i + 1}. ${p}`));
console.log(`\nBandingkan: diff legacy-site/.original/index.html legacy-site/index.html`);
