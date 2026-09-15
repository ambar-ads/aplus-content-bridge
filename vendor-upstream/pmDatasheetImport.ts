import ExcelJS from 'exceljs';
import type { ProductLineCode } from '@/types/specForm';
import {
  PM_FIRST_MODEL_COL,
  PM_MODEL_NAME_ROW,
  PM_PART_NO_ROW,
  PM_DATASHEETS,
  PM_ROW_ORDER,
} from '@/config/pmDatasheetLayout';
import { normalizeLabel } from '@/config/specSchema';
import {
  buildFieldLookup,
  extractInlineDisclaimers,
  resolveLabel,
  sanitize,
  type FootnoteSuggestion,
  type ParsedProduct,
  type ParsedTemplate,
} from './specSource';

/**
 * Read an ASUS PM price list the way ASUS ships it.
 *
 * The point of this module is that marketing should not have to convert anything. The PM
 * workbook is already field-per-row / model-per-column, and the catalog already knows ASUS's
 * column spellings, so the conversion step that used to mean retyping ~75 values per model
 * into a differently-shaped workbook is just a file upload.
 *
 * Nothing here assumes sheet names. ASUS renames tabs between releases and ships TKDN variants
 * with their own naming, so a sheet is recognised by its shape — "Part No" in A2 and "Sales
 * Model Name" in A9 — and mapped to a product line by name only afterwards.
 */

export interface PmSheetResult {
  sheetName: string;
  line: ProductLineCode;
  /** TKDN models are a separate catalogue with non-overlapping part numbers. */
  group: 'main' | 'tkdn';
  parsed: ParsedTemplate;
  /** True when the line came from the sheet name rather than the recorded layout. */
  lineWasGuessed: boolean;
}

export interface PmParseResult {
  sheets: PmSheetResult[];
  errors: string[];
  warnings: string[];
}

const KNOWN_SHEETS = new Map(PM_DATASHEETS.map((s) => [s.sheetName.toLowerCase(), s]));

/**
 * Map a sheet name to a product line.
 *
 * Underscore is a word character, so `\b` never fires inside "TKDN_AiO_Datasheet" — which is
 * exactly how two real sheets went unrecognised. Split on any run of non-alphanumerics and
 * match whole words instead.
 */
export function productLineFromSheetName(sheetName: string): ProductLineCode | null {
  const known = KNOWN_SHEETS.get(sheetName.toLowerCase());
  if (known) return known.line;

  const name = ` ${sheetName.toLowerCase().replace(/[^a-z0-9]+/g, ' ')} `;
  if (/chromebook| cb /.test(name)) return 'CB';
  if (/notebook| nb /.test(name)) return 'NX';
  if (/all in one| aio /.test(name)) return 'PT';
  if (/desktop| dt /.test(name)) return 'PF';
  return null;
}

/** A sheet is a datasheet when A2 is "Part No" and A9 is "Sales Model Name". */
function isDatasheet(ws: ExcelJS.Worksheet): boolean {
  if (ws.rowCount < PM_MODEL_NAME_ROW) return false;
  return (
    normalizeLabel(sanitize(ws.getCell(PM_PART_NO_ROW, 1).value)) === 'partno' &&
    normalizeLabel(sanitize(ws.getCell(PM_MODEL_NAME_ROW, 1).value)) === 'salesmodelname'
  );
}

function readSheet(ws: ExcelJS.Worksheet, line: ProductLineCode): ParsedTemplate {
  const lookup = buildFieldLookup(line);
  const warnings: string[] = [];

  // A column counts when it has a Part No or a Model Name. Do not stop at the first empty
  // column: TKDN_Notebook has a gap in the middle and everything past it would be lost.
  const columns: number[] = [];
  for (let c = PM_FIRST_MODEL_COL; c <= ws.columnCount; c += 1) {
    if (sanitize(ws.getCell(PM_MODEL_NAME_ROW, c).value) || sanitize(ws.getCell(PM_PART_NO_ROW, c).value)) {
      columns.push(c);
    }
  }

  const products: ParsedProduct[] = columns.map((c) => ({
    columnLabel: sanitize(ws.getCell(PM_MODEL_NAME_ROW, c).value) || `Column ${c}`,
    modelName: '',
    specs: {},
    productInfo: {},
    unmapped: [],
  }));

  const suggestions: FootnoteSuggestion[] = [];
  const seen = new Set<string>();
  const recorded = new Set(
    (PM_ROW_ORDER[line] ?? []).filter((l): l is string => !!l).map(normalizeLabel),
  );
  const newLabels: string[] = [];

  for (let r = PM_PART_NO_ROW; r <= ws.rowCount; r += 1) {
    const label = sanitize(ws.getCell(r, 1).value);
    // Row 1 is a navigation link back to the workbook's menu, not data.
    if (!label || /^click to/i.test(label)) continue;

    const keys = resolveLabel(lookup, label);
    if (recorded.size > 0 && !recorded.has(normalizeLabel(label))) newLabels.push(label);

    columns.forEach((c, index) => {
      const raw = sanitize(ws.getCell(r, c).value);
      if (!raw) return;

      const { cleaned, disclaimers } = extractInlineDisclaimers(raw);
      const value = cleaned || raw;
      const product = products[index];

      if (keys.length === 0) {
        product.unmapped.push({ label, value });
      } else {
        keys.forEach((key) => {
          if (key.startsWith('info.')) {
            const bare = key.slice(5);
            product.productInfo[bare] = value;
            if (bare === 'modelName') product.modelName = value;
          } else {
            product.specs[key] = value;
          }
        });
      }

      disclaimers.forEach((text) => {
        const anchor = keys[0] ?? label;
        const dedupe = `${anchor}|${text}`;
        if (seen.has(dedupe)) return;
        seen.add(dedupe);
        suggestions.push({
          fieldKey: anchor,
          fieldLabel: lookup.labelOf.get(anchor) ?? label,
          text,
          cleanedValue: value,
        });
      });
    });
  }

  if (newLabels.length > 0) {
    warnings.push(
      `${newLabels.length} row(s) are new since this sheet's layout was recorded ` +
        `(${newLabels.slice(0, 4).join(', ')}${newLabels.length > 4 ? '…' : ''}). ` +
        'They still import as additional specifications — re-run scripts/derivePmLayout.mjs to ' +
        'keep the Excel template paste-aligned.',
    );
  }

  const nameless = products.filter((p) => !p.modelName).length;
  if (nameless > 0) {
    warnings.push(`${nameless} column(s) have no Sales Model Name and cannot be published as-is.`);
  }

  return {
    productLine: line,
    schemaVersion: null,
    products,
    footnotes: [],
    suggestions,
    errors: [],
    warnings,
  };
}

/**
 * Parse every datasheet in an ASUS PM workbook.
 *
 * Returns one entry per sheet rather than a merged list: a batch has to be one product line,
 * and one sheet is exactly one line.
 */
export async function parsePmWorkbook(
  file: File | Blob | ArrayBuffer | Uint8Array,
): Promise<PmParseResult> {
  const wb = new ExcelJS.Workbook();
  const buffer =
    file instanceof ArrayBuffer
      ? file
      : ArrayBuffer.isView(file)
        ? (file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer)
        : await file.arrayBuffer();
  await wb.xlsx.load(buffer);

  const sheets: PmSheetResult[] = [];
  const errors: string[] = [];
  const warnings: string[] = [];

  wb.eachSheet((ws) => {
    if (!isDatasheet(ws)) return;
    const line = productLineFromSheetName(ws.name);
    if (!line) {
      warnings.push(
        `Sheet "${ws.name}" looks like a datasheet but does not name a product line. ` +
          'Import it through the Excel template instead, or rename the sheet.',
      );
      return;
    }
    sheets.push({
      sheetName: ws.name,
      line,
      group: /^tkdn/i.test(ws.name) ? 'tkdn' : 'main',
      parsed: readSheet(ws, line),
      lineWasGuessed: !KNOWN_SHEETS.has(ws.name.toLowerCase()),
    });
  });

  if (sheets.length === 0) {
    errors.push(
      'No ASUS datasheet found in this workbook. A datasheet sheet has "Part No" in cell A2 ' +
        'and "Sales Model Name" in cell A9. Upload the ASUS spec template instead if this is ' +
        'a filled template.',
    );
  }

  return { sheets, errors, warnings };
}

/** Does this workbook look like a PM price list rather than a filled A+ template? */
export async function looksLikePmWorkbook(
  file: File | Blob | ArrayBuffer | Uint8Array,
): Promise<boolean> {
  const wb = new ExcelJS.Workbook();
  const buffer =
    file instanceof ArrayBuffer
      ? file
      : ArrayBuffer.isView(file)
        ? (file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer)
        : await file.arrayBuffer();
  await wb.xlsx.load(buffer);
  if (wb.getWorksheet('Product Data')) return false;
  let found = false;
  wb.eachSheet((ws) => {
    if (!found && isDatasheet(ws)) found = true;
  });
  return found;
}

// ─── Paste ──────────────────────────────────────────────────────────

/**
 * Split clipboard TSV into rows and cells, honouring Excel's quoting.
 *
 * `split('\t')` per line is wrong for this data and always was: Excel wraps any cell holding a
 * newline or a tab in double quotes and writes the newline literally, and ASUS values are full
 * of them — I/O ports, Storage and Expansion Slot are multi-line in every single model. A naive
 * split turns one model into a dozen half-rows.
 */
export function parseClipboardTable(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];

    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i += 1;
        } else quoted = false;
      } else cell += ch;
      continue;
    }

    if (ch === '"' && cell === '') {
      quoted = true;
    } else if (ch === '\t') {
      row.push(cell);
      cell = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += ch;
    }
  }
  if (cell !== '' || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }
  return rows.filter((r) => r.some((c) => c.trim() !== ''));
}

/**
 * Read a block copied straight out of a PM datasheet: a label column followed by one column
 * per model.
 *
 * The old paste path generated an entire template workbook in memory and re-parsed it just to
 * reuse the alias resolver, which was slow, coupled paste to the template's layout, and
 * silently dropped any value whose label happened to sit on a merged banner row. The resolver
 * is shared directly now, and a two-column paste is just the N=1 case of this.
 */
export function parsePastedBlock(text: string, line: ProductLineCode): ParsedTemplate {
  const rows = parseClipboardTable(text);
  const warnings: string[] = [];
  const errors: string[] = [];

  if (rows.length === 0) {
    return {
      productLine: line,
      schemaVersion: null,
      products: [],
      footnotes: [],
      suggestions: [],
      errors: ['Nothing to read. Copy the field-name column plus at least one model column.'],
      warnings,
    };
  }

  const modelCount = Math.max(...rows.map((r) => r.length)) - 1;
  if (modelCount < 1) {
    return {
      productLine: line,
      schemaVersion: null,
      products: [],
      footnotes: [],
      suggestions: [],
      errors: ['Only one column found. Copy the field names *and* the model column(s) beside them.'],
      warnings,
    };
  }

  const lookup = buildFieldLookup(line);
  // Column headings come from the Sales Model Name row when the paste includes it.
  const nameRow = rows.find((r) => resolveLabel(lookup, r[0] ?? '').includes('info.modelName'));

  const products: ParsedProduct[] = Array.from({ length: modelCount }, (_, i) => ({
    columnLabel: nameRow?.[i + 1]?.trim() || `Model ${i + 1}`,
    modelName: '',
    specs: {},
    productInfo: {},
    unmapped: [],
  }));

  const suggestions: FootnoteSuggestion[] = [];
  const seen = new Set<string>();

  rows.forEach((cells) => {
    const label = (cells[0] ?? '').replace(/_x000D_/g, '').trim();
    if (!label || /^click to/i.test(label)) return;
    const keys = resolveLabel(lookup, label);

    for (let i = 0; i < modelCount; i += 1) {
      const raw = (cells[i + 1] ?? '').replace(/_x000D_/g, '').replace(/\r\n?/g, '\n').trim();
      if (!raw) continue;

      const { cleaned, disclaimers } = extractInlineDisclaimers(raw);
      const value = cleaned || raw;
      const product = products[i];

      if (keys.length === 0) {
        product.unmapped.push({ label, value });
      } else {
        keys.forEach((key) => {
          if (key.startsWith('info.')) {
            const bare = key.slice(5);
            product.productInfo[bare] = value;
            if (bare === 'modelName') product.modelName = value;
          } else {
            product.specs[key] = value;
          }
        });
      }

      disclaimers.forEach((text2) => {
        const anchor = keys[0] ?? label;
        const dedupe = `${anchor}|${text2}`;
        if (seen.has(dedupe)) return;
        seen.add(dedupe);
        suggestions.push({
          fieldKey: anchor,
          fieldLabel: lookup.labelOf.get(anchor) ?? label,
          text: text2,
          cleanedValue: value,
        });
      });
    }
  });

  const nameless = products.filter((p) => !p.modelName).length;
  if (nameless > 0) {
    warnings.push(
      `${nameless} column(s) have no Model Name. Include the "Sales Model Name" row in the ` +
        'paste, or fill it in on the next step.',
    );
  }

  return {
    productLine: line,
    schemaVersion: null,
    products,
    footnotes: [],
    suggestions,
    errors,
    warnings,
  };
}
