/**
 * Membaca workbook lama `A+ Content Source*.xlsx`.
 *
 * Dua hal diambil dari sini, dan keduanya tidak ada di file Excel PM mana pun:
 *
 *   1. Tagline / KSP / Link per produk (sheet `RAW.KSP.LInk` kolom D / E / F) — copy marketing
 *      yang diketik manual oleh PM dan MKT. Inilah isi yang paling dipakai marketing di halaman
 *      lama, dan alasan kenapa "arahkan saja ke Excel baru" tidak bisa menggantikan workbook ini.
 *   2. Tabel `MKT name rule` — pemetaan kode seri ke marketing name (dan tagline default seri).
 */
import { openWorkbook } from './openWorkbook';
import { sanitize, normalizeLabel } from '@/vendor/pmCell';
import { isBlankValue } from './legacyCsv';
import { isPn90 } from './deriveForWeb';
import type { CopyRecord } from './catalogStore';

const RAW_SHEET = 'RAW.KSP.LInk';
const MKT_SHEET = 'MKT name rule';

/** Header sheet RAW ada di baris 2; baris 1 hanya mencatat siapa yang mengisi tiap kolom. */
const RAW_HEADER_ROW = 2;
const RAW_FIRST_DATA_ROW = 3;

/** Kolom pada sheet RAW.KSP.LInk. Kolom H adalah nilai hasil paste, bukan formula. */
const RAW_COL = {
  sortSource: 1, // A — formula, berisi Part No
  modelName: 2, // B — formula
  marketingName: 3, // C — formula
  tagline: 4, // D — diketik PM
  ksp: 5, // E — diketik PM
  link: 6, // F — diisi MKT
  partNo: 8, // H — awal blok raw datasheet PM
} as const;

const MKT_HEADER_ROW = 2;
const MKT_COL = { series: 1, codeRoot: 2, marketingName: 3, tagline: 4 } as const;

export interface LegacyWorkbookResult {
  /** Copy marketing, berkunci 90PN. */
  copyByPn90: Map<string, CopyRecord>;
  /** Kode seri ternormalisasi -> marketing name. */
  mktNameRule: Map<string, string>;
  /** Kode seri ternormalisasi -> tagline default seri, untuk mengisi produk baru. */
  taglineBySeries: Map<string, string>;
  stats: { rawRows: number; withCopy: number; mktRules: number };
  warnings: string[];
}

export async function readLegacyWorkbook(
  input: ArrayBuffer | Uint8Array,
  fileName?: string,
): Promise<LegacyWorkbookResult> {
  const opened = await openWorkbook(input, fileName);
  const wb = opened.workbook;
  const warnings: string[] = [...opened.warnings];

  const copyByPn90 = new Map<string, CopyRecord>();
  const mktNameRule = new Map<string, string>();
  const taglineBySeries = new Map<string, string>();
  let rawRows = 0;
  let withCopy = 0;

  const raw = wb.getWorksheet(RAW_SHEET);
  if (!raw) {
    warnings.push(`Sheet "${RAW_SHEET}" is missing — Tagline/KSP/Link cannot be read`);
  } else {
    const header = sanitize(raw.getRow(RAW_HEADER_ROW).getCell(RAW_COL.partNo).value);
    if (!/part\s*no/i.test(header)) {
      warnings.push(
        `Column H of "${RAW_SHEET}" is headed "${header}", not "Part No" — the layout may have changed`,
      );
    }

    for (let r = RAW_FIRST_DATA_ROW; r <= raw.rowCount; r += 1) {
      const row = raw.getRow(r);
      // Kolom H adalah hasil paste literal; kolom A formula. Kalau PM berhenti mem-paste raw
      // data, A ikut kosong, jadi H yang dijadikan kunci utama.
      const pn90 =
        sanitize(row.getCell(RAW_COL.partNo).value).trim() ||
        sanitize(row.getCell(RAW_COL.sortSource).value).trim();
      // Baris sisa di bawah data menghasilkan error formula dan objek ExcelJS di kolom ini.
      if (!pn90 || isBlankValue(pn90) || !isPn90(pn90)) continue;
      rawRows += 1;

      const tagline = sanitize(row.getCell(RAW_COL.tagline).value);
      const ksp = sanitize(row.getCell(RAW_COL.ksp).value);
      const link = sanitize(row.getCell(RAW_COL.link).value);
      const modelName = sanitize(row.getCell(RAW_COL.modelName).value);

      if (copyByPn90.has(pn90)) {
        warnings.push(`Duplicate 90PN "${pn90}" in ${RAW_SHEET} row ${r} — the first one is used`);
        continue;
      }
      if (!isBlankValue(tagline) || !isBlankValue(ksp) || !isBlankValue(link)) withCopy += 1;

      copyByPn90.set(pn90, { tagline, ksp, link, modelName });
    }
  }

  const mkt = wb.getWorksheet(MKT_SHEET);
  if (!mkt) {
    warnings.push(`Sheet "${MKT_SHEET}" is missing — new products will be named "Not found"`);
  } else {
    for (let r = MKT_HEADER_ROW + 1; r <= mkt.rowCount; r += 1) {
      const row = mkt.getRow(r);
      const code = sanitize(row.getCell(MKT_COL.codeRoot).value).trim();
      const name = sanitize(row.getCell(MKT_COL.marketingName).value).trim();
      if (!code || !name) continue;
      const key = normalizeLabel(code);
      if (!mktNameRule.has(key)) {
        mktNameRule.set(key, name);
        const tagline = sanitize(row.getCell(MKT_COL.tagline).value).trim();
        if (tagline) taglineBySeries.set(key, tagline);
      }
    }
  }

  return {
    copyByPn90,
    mktNameRule,
    taglineBySeries,
    stats: { rawRows, withCopy, mktRules: mktNameRule.size },
    warnings,
  };
}
