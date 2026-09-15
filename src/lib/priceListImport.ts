/**
 * Reads the datasheet sheets out of an ASUS price list.
 *
 * Sheets are recognised by structure, never by name: column A must carry a "Part No" row followed
 * by a "Sales Model Name" row. ASUS renames tabs between files, so the name is used only to guess
 * the product line and to mark the separate TKDN catalogue.
 *
 * The layout is field-per-row, model-per-column, but WHICH rows hold the anchors varies between
 * ASUS's own files — see `findDatasheetLayout`. One workbook can even mix three layouts.
 *
 * A monthly workbook often carries two vintages of the same datasheet (last week's and this
 * week's). Both are read; the first occurrence of a 90PN wins and the rest are reported.
 */
import type ExcelJS from 'exceljs';
import { openWorkbook } from './openWorkbook';
import { sanitize, normalizeLabel, productLineFromSheetName, sheetGroupFromName } from '@/vendor/pmCell';
import { PM_FIRST_MODEL_COL } from '@/vendor/pmLayout';
import { makePmRecord, productLineFromPn90, isPn90 } from './deriveForWeb';
import type { PmLine, PmRecord } from './deriveForWeb';
import type { ProductLineCode } from '@/vendor/types';

export interface PriceListModel {
  pn90: string;
  modelName: string;
  sheetName: string;
  /** Product line menurut 90PN — inilah yang dipakai untuk derivasi. */
  line: PmLine | null;
  /** Product line menurut nama sheet, sebagai pembanding. */
  sheetLine: ProductLineCode | null;
  group: 'main' | 'tkdn';
  record: PmRecord;
}

export interface PriceListSheet {
  name: string;
  /** Which rows this sheet keeps its anchors on — they differ between ASUS's own files. */
  partNoRow: number;
  modelNameRow: number;
  line: ProductLineCode | null;
  group: 'main' | 'tkdn';
  models: number;
  labels: number;
}

export interface PriceListResult {
  sheets: PriceListSheet[];
  models: PriceListModel[];
  warnings: string[];
}

interface SheetLayout {
  partNoRow: number;
  modelNameRow: number;
}

/** How far down column A the two anchor labels are worth looking for. */
const LAYOUT_SEARCH_ROWS = 20;

/**
 * Find where a datasheet keeps its two anchor rows.
 *
 * The layout is not fixed across ASUS's own files. The ABP price list puts Part No on row 2 and
 * Sales Model Name on row 9, with rows 3-8 holding embedded photos. The ACID monthly price list
 * uses three different layouts in one workbook: rows 1/2 for Desktop and All-in-One, rows 1/8 for
 * Commercial Notebook, rows 2/9 for the TKDN sheets.
 *
 * So the rows are searched for rather than assumed. Hard-coding them is what made only the three
 * TKDN sheets import from the ACID file while six others were silently skipped.
 */
function findDatasheetLayout(ws: ExcelJS.Worksheet): SheetLayout | null {
  let partNoRow = 0;
  let modelNameRow = 0;
  const last = Math.min(LAYOUT_SEARCH_ROWS, ws.rowCount);

  for (let r = 1; r <= last; r += 1) {
    const label = normalizeLabel(sanitize(ws.getCell(r, 1).value));
    if (!partNoRow && label === 'partno') partNoRow = r;
    else if (partNoRow && !modelNameRow && label === 'salesmodelname') modelNameRow = r;
  }

  if (!partNoRow || !modelNameRow || modelNameRow <= partNoRow) return null;
  return { partNoRow, modelNameRow };
}

export async function readPriceList(
  input: ArrayBuffer | Uint8Array,
  fileName?: string,
): Promise<PriceListResult> {
  const opened = await openWorkbook(input, fileName);
  const wb = opened.workbook;

  const sheets: PriceListSheet[] = [];
  const models: PriceListModel[] = [];
  const warnings: string[] = [...opened.warnings];
  const seen = new Map<string, string>();
  const duplicatePairs = new Map<string, number>();

  wb.eachSheet((ws) => {
    const layout = findDatasheetLayout(ws);
    if (!layout) return;
    const { partNoRow, modelNameRow } = layout;

    const sheetLine = productLineFromSheetName(ws.name);
    const group = sheetGroupFromName(ws.name);

    // Kumpulkan label kolom A sekali saja, lalu pakai ulang untuk setiap kolom model.
    const labels: { row: number; label: string }[] = [];
    for (let r = partNoRow; r <= ws.rowCount; r += 1) {
      const label = sanitize(ws.getCell(r, 1).value);
      // Baris 1 berisi tautan navigasi "Click to Back to Main Menu".
      if (!label || /^click to/i.test(label)) continue;
      labels.push({ row: r, label });
    }

    // Jangan berhenti di kolom kosong pertama: TKDN_Notebook punya celah di tengah, dan semua
    // model setelahnya akan hilang kalau loop-nya diputus.
    const columns: number[] = [];
    for (let c = PM_FIRST_MODEL_COL; c <= ws.columnCount; c += 1) {
      const pn = sanitize(ws.getCell(partNoRow, c).value);
      const mn = sanitize(ws.getCell(modelNameRow, c).value);
      if (pn || mn) columns.push(c);
    }

    for (const c of columns) {
      const pn90 = sanitize(ws.getCell(partNoRow, c).value).trim();
      const modelName = sanitize(ws.getCell(modelNameRow, c).value).trim();
      if (!pn90) {
        warnings.push(`${ws.name}: model column "${modelName || '(unnamed)'}" has no Part No — skipped`);
        continue;
      }
      // Sebagian sheet mengulang label kolom A di kolom B, sehingga terbaca sebagai model palsu
      // bernama "Sales Model Name" dengan part number "Part No".
      if (!isPn90(pn90)) continue;
      if (seen.has(pn90)) {
        // A monthly workbook repeats whole datasheets, so this fires hundreds of times. Counted
        // per sheet pair and reported once, or the real warnings drown in it.
        const pair = `${seen.get(pn90)} → ${ws.name}`;
        duplicatePairs.set(pair, (duplicatePairs.get(pair) ?? 0) + 1);
        continue;
      }
      seen.set(pn90, ws.name);

      const pairs: [string, string][] = labels.map(({ row, label }) => [
        label,
        sanitize(ws.getCell(row, c).value),
      ]);

      const line = productLineFromPn90(pn90);
      if (!line) {
        warnings.push(`${modelName || pn90}: 90PN prefix "${pn90.slice(0, 4)}" is not recognised (not 90NX/90PT/90PF)`);
      }

      models.push({
        pn90,
        modelName,
        sheetName: ws.name,
        line,
        sheetLine,
        group,
        record: makePmRecord(pairs),
      });
    }

    sheets.push({
      name: ws.name,
      line: sheetLine,
      group,
      models: columns.length,
      labels: labels.length,
      partNoRow,
      modelNameRow,
    });
  });

  for (const [pair, count] of duplicatePairs) {
    warnings.push(
      `${count} model(s) appear in both "${pair.split(' → ')[0]}" and "${pair.split(' → ')[1]}" — ` +
        'the first sheet was used. That is normal when a workbook carries more than one week of data.',
    );
  }

  if (!sheets.length) {
    warnings.push(
      'No datasheet sheets were found. A datasheet is recognised by a "Part No" row followed by a "Sales Model Name" row in column A — check that this really is a price list.',
    );
  }

  return { sheets, models, warnings };
}
