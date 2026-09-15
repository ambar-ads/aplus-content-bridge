/**
 * File Excel baru yang menggantikan `A+ Content Source (1).xlsx`.
 *
 * Workbook lama punya rantai dependensi berlapis: sheet `ForWeb` bergantung pada dua rumus spill
 * yang digabung berdasarkan urutan sort, yang bergantung pada `RAW.KSP.LInk` dan tiga sheet
 * `Auto-*`, yang masing-masing menunjuk blok raw datasheet lewat huruf kolom dan XLOOKUP ke
 * `MKT name rule`. Satu baris bergeser di datasheet PM merusak seluruh rantai itu, diam-diam.
 *
 * File ini dirancang dengan aturan sebaliknya:
 *
 *   - **Tanpa formula.** Tidak ada satu pun sel berisi rumus. Semua nilai literal.
 *   - **Satu baris, satu produk.** Tidak ada penggabungan antar-sheet.
 *   - **Dibaca lewat judul kolom**, bukan posisi. Menambah atau menggeser kolom tidak merusak apa pun.
 *   - **Bolak-balik.** Alat yang mengisinya dari price list, tim marketing yang melengkapi copy,
 *     lalu alat membacanya kembali. Filenya selalu lengkap dan bisa dibaca sendiri tanpa alat.
 */
import ExcelJS from 'exceljs';
import { openWorkbook } from './openWorkbook';
import { sanitize, normalizeLabel } from '@/vendor/pmCell';
import { EMPTY_FORWEB_ROW } from './config';
import { isBlankValue } from './legacyCsv';
import type { ForWebKey, ForWebRow } from './types';

export const MASTER_FILENAME = 'A+ Content Master.xlsx';

const SHEET_PRODUCTS = 'Products';
const SHEET_SERIES = 'Series';
const SHEET_GUIDE = 'Guide';

/** Earlier builds used Indonesian tab names; files already in circulation still open cleanly. */
const SERIES_SHEET_ALIASES = [SHEET_SERIES, 'Seri'];

/** Judul kolom ada di baris 1; data mulai baris 2. Sengaja sesederhana mungkin. */
const HEADER_ROW = 1;
const FIRST_DATA_ROW = 2;

type Owner = 'auto' | 'PM' | 'MKT';

export interface MasterColumn {
  key: ForWebKey;
  header: string;
  width: number;
  wrap?: boolean;
  owner: Owner;
}

/**
 * Kolom file master, berpadanan satu-satu dengan CSV ForWeb.
 *
 * `Part Number` di CSV muncul dua kali; di sini cukup satu kolom `90PN`, dan saat menulis CSV
 * nilainya dipakai untuk kedua kolom itu. Pada 392 produk yang tayang, keduanya memang selalu
 * berisi nilai yang sama.
 */
export const MASTER_COLUMNS: MasterColumn[] = [
  { key: 'partNumber', header: '90PN', width: 19, owner: 'auto' },
  { key: 'productName', header: 'Product Name', width: 24, owner: 'auto' },
  { key: 'marketingName', header: 'Marketing Name', width: 27, owner: 'PM' },
  { key: 'tagline', header: 'Tagline', width: 38, wrap: true, owner: 'PM' },
  { key: 'ksp', header: 'KSP', width: 52, wrap: true, owner: 'PM' },
  { key: 'link', header: 'Link', width: 34, owner: 'MKT' },
  { key: 'title', header: 'Title', width: 46, wrap: true, owner: 'auto' },
  { key: 'cpu', header: 'CPU', width: 30, wrap: true, owner: 'auto' },
  { key: 'gpu', header: 'GPU', width: 26, wrap: true, owner: 'auto' },
  { key: 'os', header: 'OS', width: 16, owner: 'auto' },
  { key: 'office', header: 'Office', width: 24, wrap: true, owner: 'auto' },
  { key: 'panelSize', header: 'Panel Size', width: 11, owner: 'auto' },
  { key: 'resolution', header: 'Resolution', width: 24, owner: 'auto' },
  { key: 'brightness', header: 'Brightness', width: 14, owner: 'auto' },
  { key: 'ram', header: 'RAM', width: 20, wrap: true, owner: 'auto' },
  { key: 'ssd', header: 'SSD', width: 28, wrap: true, owner: 'auto' },
  { key: 'ioPort', header: 'IO Port', width: 38, wrap: true, owner: 'auto' },
  { key: 'weight', header: 'Weight (with Battery)', width: 20, owner: 'auto' },
  { key: 'size', header: 'Size', width: 26, owner: 'auto' },
  { key: 'expansionSlot', header: 'Expansion Slot', width: 30, wrap: true, owner: 'auto' },
  { key: 'includeInTheBox', header: 'Include in the box', width: 28, wrap: true, owner: 'auto' },
  { key: 'battery', header: 'Battery', width: 13, owner: 'auto' },
  { key: 'powerSupply', header: 'Power Supply', width: 22, owner: 'auto' },
  { key: 'security', header: 'Security', width: 30, wrap: true, owner: 'auto' },
];

/** Kolom yang kalau kosong membuat produk tayang tanpa isi — ditandai kuning saat dibangkitkan. */
const COPY_KEYS: ForWebKey[] = ['tagline', 'ksp', 'link'];

const HEAD_FILL = 'FF1F3A52';
const HEAD_PM = 'FF2F5D3F';
const HEAD_MKT = 'FF5C4326';
const NEEDS_FILL = 'FFFFF3CD';
const NEW_FILL = 'FFE7F1FA';
const BORDER = 'FFD5DCE4';

export interface SeriesEntry {
  code: string;
  marketingName: string;
  tagline?: string;
}

export interface BuildMasterOptions {
  rows: ForWebRow[];
  series?: SeriesEntry[];
  /** 90PN produk yang baru masuk bulan ini, ditandai biru agar mudah ditemukan. */
  newPn90?: Set<string>;
  generatedAt?: Date;
}

function headerColor(owner: Owner): string {
  if (owner === 'PM') return HEAD_PM;
  if (owner === 'MKT') return HEAD_MKT;
  return HEAD_FILL;
}

export async function buildMasterWorkbook(options: BuildMasterOptions): Promise<ArrayBuffer> {
  const { rows, series = [], newPn90 = new Set<string>() } = options;
  const generatedAt = options.generatedAt ?? new Date();

  const wb = new ExcelJS.Workbook();
  wb.creator = 'A+ Content Bridge';
  wb.created = generatedAt;

  // ── Petunjuk ──────────────────────────────────────────────────────
  const guide = wb.addWorksheet(SHEET_GUIDE, {
    properties: { tabColor: { argb: HEAD_FILL } },
  });
  guide.columns = [{ width: 4 }, { width: 104 }];
  const guideLines: [string, boolean][] = [
    ['A+ Content Master', true],
    ['', false],
    [`Generated ${generatedAt.toLocaleDateString('en-GB', { day: 'numeric', month: 'long', year: 'numeric' })} — ${rows.length} products.`, false],
    ['', false],
    ['This file replaces A+ Content Source.xlsx. There are no formulas in it, and no sheet', false],
    ['depends on any other sheet. One row is one product.', false],
    ['', false],
    ['What you fill in on the Products sheet:', true],
    ['', false],
    ['  Marketing Name, Tagline, KSP   — filled by PM (green column headings)', false],
    ['  Link                            — filled by MKT (brown column heading)', false],
    ['', false],
    ['Every other column is filled automatically from the ABP Price List. You may overwrite them', false],
    ['where needed; anything you type is never written back over.', false],
    ['', false],
    ['Cell colours:', true],
    ['', false],
    ['  Yellow   — Tagline, KSP or Link is still empty. The product still appears on the site,', false],
    ['             but shows "No Available" and the Materials button says "Coming Soon".', false],
    ['  Blue     — product added this month.', false],
    ['', false],
    ['Rules:', true],
    ['', false],
    ['  Do not change the 90PN column. It is the key used to match products.', false],
    ['  Do not delete rows for older products, even if they are no longer sold.', false],
    ['  Do not change the column headings in row 1.', false],
    ['  Adding new columns on the right is fine — they are ignored.', false],
    ['', false],
    ['Write KSP one point per line. Press Alt+Enter for a new line inside a cell.', false],
  ];
  guideLines.forEach(([text, bold], i) => {
    const cell = guide.getCell(i + 2, 2);
    cell.value = text;
    cell.font = { name: 'Calibri', size: bold ? 12 : 11, bold };
    if (i === 0) cell.font = { name: 'Calibri', size: 16, bold: true, color: { argb: HEAD_FILL } };
  });

  // ── Products ──────────────────────────────────────────────────────
  const ws = wb.addWorksheet(SHEET_PRODUCTS, {
    views: [{ state: 'frozen', xSplit: 2, ySplit: HEADER_ROW }],
    properties: { tabColor: { argb: HEAD_PM } },
  });
  ws.columns = MASTER_COLUMNS.map((c) => ({ width: c.width }));

  const header = ws.getRow(HEADER_ROW);
  MASTER_COLUMNS.forEach((col, i) => {
    const cell = header.getCell(i + 1);
    cell.value = col.header;
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: headerColor(col.owner) } };
    cell.alignment = { vertical: 'middle', horizontal: 'left' };
  });
  header.height = 24;
  ws.autoFilter = {
    from: { row: HEADER_ROW, column: 1 },
    to: { row: HEADER_ROW, column: MASTER_COLUMNS.length },
  };

  rows.forEach((row, r) => {
    const excelRow = ws.getRow(FIRST_DATA_ROW + r);
    const needsCopy = COPY_KEYS.some((k) => isBlankValue(row[k] ?? ''));
    const isNew = newPn90.has((row.partNumber ?? '').trim());

    MASTER_COLUMNS.forEach((col, i) => {
      const cell = excelRow.getCell(i + 1);
      // Sentinel "0" warisan workbook lama tidak dibawa ke file baru — di sini kosong berarti kosong.
      const value = row[col.key] ?? '';
      cell.value = isBlankValue(value) ? null : value;
      cell.alignment = { vertical: 'top', wrapText: col.wrap ?? false };
      cell.font = { name: 'Calibri', size: 11 };
      cell.border = {
        bottom: { style: 'hair', color: { argb: BORDER } },
        right: { style: 'hair', color: { argb: BORDER } },
      };
      if (needsCopy && COPY_KEYS.includes(col.key)) {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NEEDS_FILL } };
      } else if (isNew && col.key === 'partNumber') {
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NEW_FILL } };
      }
    });
  });

  // ── Seri ──────────────────────────────────────────────────────────
  const seriesSheet = wb.addWorksheet(SHEET_SERIES, {
    views: [{ state: 'frozen', ySplit: HEADER_ROW }],
  });
  seriesSheet.columns = [{ width: 16 }, { width: 32 }, { width: 52 }];
  ['Series Code', 'Marketing Name', 'Default Tagline'].forEach((h, i) => {
    const cell = seriesSheet.getCell(HEADER_ROW, i + 1);
    cell.value = h;
    cell.font = { name: 'Calibri', size: 11, bold: true, color: { argb: 'FFFFFFFF' } };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: HEAD_FILL } };
  });
  seriesSheet.getCell(HEADER_ROW, 4).value =
    'Used only to name products that arrive for the first time. If a series code is missing here, the new product is left without a Marketing Name for you to fill in.';
  seriesSheet.getCell(HEADER_ROW, 4).font = { name: 'Calibri', size: 10, italic: true };
  series.forEach((s, i) => {
    const row = seriesSheet.getRow(FIRST_DATA_ROW + i);
    row.getCell(1).value = s.code;
    row.getCell(2).value = s.marketingName;
    row.getCell(3).value = s.tagline ?? null;
    row.getCell(3).alignment = { wrapText: true, vertical: 'top' };
  });

  return (await wb.xlsx.writeBuffer()) as ArrayBuffer;
}

// ─── Membaca kembali ────────────────────────────────────────────────

export interface MasterReadResult {
  rows: ForWebRow[];
  series: SeriesEntry[];
  /** Judul kolom yang ada di file tapi tidak dikenali — diabaikan, bukan error. */
  unknownColumns: string[];
  missingColumns: string[];
  warnings: string[];
}

export async function readMasterWorkbook(
  input: ArrayBuffer | Uint8Array,
  fileName?: string,
): Promise<MasterReadResult> {
  const opened = await openWorkbook(input, fileName);
  const wb = opened.workbook;

  const warnings: string[] = [...opened.warnings];
  const ws = wb.getWorksheet(SHEET_PRODUCTS);
  if (!ws) {
    throw new Error(
      `This file has no "${SHEET_PRODUCTS}" sheet. Check that you picked ${MASTER_FILENAME} and not the old workbook.`,
    );
  }

  // Pemetaan berdasarkan JUDUL kolom, bukan posisi. Inilah yang membuat file ini tahan terhadap
  // penambahan atau penggeseran kolom — kelemahan utama sheet Auto-* di workbook lama.
  const byHeader = new Map<string, number>();
  const unknownColumns: string[] = [];
  const known = new Map(MASTER_COLUMNS.map((c) => [normalizeLabel(c.header), c.key]));

  const headerRow = ws.getRow(HEADER_ROW);
  for (let c = 1; c <= ws.columnCount; c += 1) {
    const label = sanitize(headerRow.getCell(c).value);
    if (!label) continue;
    const norm = normalizeLabel(label);
    if (known.has(norm)) {
      if (!byHeader.has(norm)) byHeader.set(norm, c);
    } else {
      unknownColumns.push(label);
    }
  }

  const missingColumns = MASTER_COLUMNS.filter(
    (c) => !byHeader.has(normalizeLabel(c.header)),
  ).map((c) => c.header);

  const pnCol = byHeader.get(normalizeLabel('90PN'));
  if (!pnCol) {
    throw new Error('No "90PN" column in the Products sheet. The column headings in row 1 must not be changed.');
  }

  const rows: ForWebRow[] = [];
  const seen = new Set<string>();

  for (let r = FIRST_DATA_ROW; r <= ws.rowCount; r += 1) {
    const excelRow = ws.getRow(r);
    const pn90 = sanitize(excelRow.getCell(pnCol).value).trim();
    if (!pn90) continue;

    if (seen.has(pn90)) {
      warnings.push(`Duplicate 90PN "${pn90}" in row ${r} — the first one is used`);
      continue;
    }
    seen.add(pn90);

    const row = { ...EMPTY_FORWEB_ROW } as ForWebRow;
    for (const col of MASTER_COLUMNS) {
      const c = byHeader.get(normalizeLabel(col.header));
      row[col.key] = c ? sanitize(excelRow.getCell(c).value) : '';
    }
    // CSV ForWeb memuat Part Number dua kali; keduanya selalu bernilai sama.
    row.partNumber = pn90;
    row.partNumberAuto = pn90;
    rows.push(row);
  }

  const series: SeriesEntry[] = [];
  const seriesSheet = SERIES_SHEET_ALIASES.map((n) => wb.getWorksheet(n)).find(Boolean);
  if (seriesSheet) {
    for (let r = FIRST_DATA_ROW; r <= seriesSheet.rowCount; r += 1) {
      const code = sanitize(seriesSheet.getRow(r).getCell(1).value).trim();
      const marketingName = sanitize(seriesSheet.getRow(r).getCell(2).value).trim();
      if (!code || !marketingName) continue;
      series.push({
        code,
        marketingName,
        tagline: sanitize(seriesSheet.getRow(r).getCell(3).value).trim() || undefined,
      });
    }
  }

  if (unknownColumns.length) {
    warnings.push(`Unrecognised columns, ignored: ${unknownColumns.join(', ')}`);
  }
  if (missingColumns.length) {
    warnings.push(`Missing columns, treated as empty: ${missingColumns.join(', ')}`);
  }

  return { rows, series, unknownColumns, missingColumns, warnings };
}
