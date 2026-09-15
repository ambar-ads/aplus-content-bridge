/**
 * Baca & tulis CSV "A+ Content Source(ForWeb)" persis seperti yang dihasilkan Excel.
 *
 * Kontraknya bukan sekadar "CSV": halaman lama sudah berjalan bertahun-tahun di atas bentuk byte
 * tertentu, dan yang membuatnya cocok adalah detail-detail berikut.
 *
 *   - UTF-8 **dengan BOM** — hasil "CSV UTF-8" dari Excel. Efek sampingnya: saat PapaParse
 *     memakai `header: true`, nama kolom pertama menjadi "\uFEFFPart Number", sehingga
 *     `delete row['Part Number']` di halaman lama justru menghapus kolom ke-7. Itu sudah jadi
 *     perilaku yang berjalan; jangan "diperbaiki" dengan membuang BOM.
 *   - Antar-record **CRLF**, tetapi newline **di dalam** sel adalah **LF telanjang**.
 *   - Quoting minimal: hanya kalau ada koma, kutip ganda, atau newline.
 *   - Header `Part Number` sengaja muncul dua kali.
 *
 * Parser di sini sengaja membaca **berdasarkan posisi**, bukan nama kolom, justru karena dua hal
 * di atas (BOM + header kembar) membuat pembacaan berbasis nama ambigu.
 */
import { FORWEB_COLUMNS, EMPTY_FORWEB_ROW } from './config';
import type { ForWebRow } from './types';

const BOM = '\uFEFF';

/** Pecah teks CSV jadi array record, menghormati kutip dan newline di dalam sel. */
export function parseCsvGrid(text: string): string[][] {
  const src = text.startsWith(BOM) ? text.slice(1) : text;
  const rows: string[][] = [];
  let row: string[] = [];
  let field = '';
  let quoted = false;

  for (let i = 0; i < src.length; i += 1) {
    const ch = src[i];

    if (quoted) {
      if (ch === '"') {
        if (src[i + 1] === '"') { field += '"'; i += 1; }
        else quoted = false;
      } else {
        field += ch;
      }
      continue;
    }

    if (ch === '"') { quoted = true; continue; }
    if (ch === ',') { row.push(field); field = ''; continue; }
    if (ch === '\r' || ch === '\n') {
      if (ch === '\r' && src[i + 1] === '\n') i += 1;
      row.push(field);
      rows.push(row);
      row = [];
      field = '';
      continue;
    }
    field += ch;
  }

  if (field !== '' || row.length) { row.push(field); rows.push(row); }
  return rows;
}

/** Nilai yang dipakai workbook lama sebagai penanda "kosong". */
export function isBlankValue(v: string): boolean {
  const t = v.trim();
  return t === '' || t === '0';
}

/**
 * Baca CSV ForWeb jadi baris terstruktur.
 *
 * Baris hantu dibuang: file yang tayang berisi 2003 record padahal hanya 392 produk sungguhan —
 * sisanya adalah spill range kosong dari Excel, sekitar 80% isi file.
 */
export function parseForWebCsv(text: string): { rows: ForWebRow[]; ghostRows: number; headerMatches: boolean } {
  const grid = parseCsvGrid(text);
  if (!grid.length) return { rows: [], ghostRows: 0, headerMatches: false };

  const header = grid[0].map((h) => h.replace(/^\uFEFF/, '').trim());
  const headerMatches =
    header.length === FORWEB_COLUMNS.length &&
    FORWEB_COLUMNS.every((c, i) => header[i] === c.header);

  const rows: ForWebRow[] = [];
  let ghostRows = 0;

  for (let r = 1; r < grid.length; r += 1) {
    const cells = grid[r];
    const productName = (cells[1] ?? '').trim();
    if (isBlankValue(productName)) { ghostRows += 1; continue; }

    // Nilai disimpan APA ADANYA, termasuk sentinel "0". Menormalkannya di sini membuat katalog
    // tidak bisa menghasilkan ulang file yang sekarang tayang byte-per-byte, padahal kesetiaan
    // itulah yang dipakai `check:forweb` untuk membuktikan tidak ada teks terbit yang berubah.
    // Pertanyaan "isinya kosong atau tidak" dijawab lewat `isBlankValue`, bukan bentuk simpanannya.
    const row = { ...EMPTY_FORWEB_ROW } as ForWebRow;
    FORWEB_COLUMNS.forEach((col, i) => {
      row[col.key] = cells[i] ?? '';
    });
    rows.push(row);
  }

  return { rows, ghostRows, headerMatches };
}

function quote(value: string): string {
  if (value === '') return '';
  return /[",\n\r]/.test(value) ? `"${value.replace(/"/g, '""')}"` : value;
}

/** Tulis CSV ForWeb dalam bentuk byte yang sama dengan export Excel. */
export function serializeForWebCsv(rows: ForWebRow[]): string {
  const lines: string[] = [];
  lines.push(FORWEB_COLUMNS.map((c) => quote(c.header)).join(','));

  for (const row of rows) {
    const cells = FORWEB_COLUMNS.map((col) => {
      // Newline di dalam sel ditulis apa adanya. File yang tayang campur: sebagian besar LF,
      // tetapi 123 sel memakai CRLF. Menormalkannya membuat file tidak lagi identik dengan yang
      // sudah terbit tanpa alasan yang jelas. Nilai baru hasil derivasi sudah lewat `sanitize()`
      // di sisi import, jadi memang sudah LF.
      const raw = row[col.key] ?? '';
      return quote(raw.trim() === '' ? col.emptyAs : raw);
    });
    lines.push(cells.join(','));
  }

  // CRLF antar-record; LF di dalam sel sudah dipertahankan apa adanya oleh `quote`.
  return BOM + lines.join('\r\n') + '\r\n';
}
