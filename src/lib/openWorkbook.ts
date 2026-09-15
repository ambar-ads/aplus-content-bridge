/**
 * One place where every .xlsx this tool reads is opened.
 *
 * It exists because the files ASUS sends are not always well formed, and each reader used to
 * repeat its own `new Workbook(); load(...)`. Centralising it means a file that needs recovery is
 * recovered the same way whichever slot it is dropped into, and the recovery is reported rather
 * than happening silently.
 */
import ExcelJS from 'exceljs';
import { hasEndOfCentralDirectory, rebuildZip } from './zipRepair';

export interface OpenedWorkbook {
  workbook: ExcelJS.Workbook;
  /** True when the file had to be rebuilt before it could be read. */
  repaired: boolean;
  warnings: string[];
}

export async function openWorkbook(
  input: ArrayBuffer | Uint8Array,
  fileName = 'this file',
): Promise<OpenedWorkbook> {
  const bytes = input instanceof Uint8Array ? input : new Uint8Array(input);
  const warnings: string[] = [];
  let usable = bytes;
  let repaired = false;

  if (!hasEndOfCentralDirectory(bytes)) {
    // Truncated download or copy. Excel recovers these silently; do the same rather than making
    // someone re-save the file by hand every month.
    const result = await rebuildZip(bytes);
    usable = result.buffer;
    repaired = true;
    warnings.push(
      `"${fileName}" was cut short and has been recovered: ${result.kept} parts read` +
        (result.dropped.length ? `, ${result.dropped.length} incomplete part(s) left out` : '') +
        '. Cell values are unaffected, but ask the sender for a complete copy when you can.',
    );
    if (result.cleaned.length) {
      warnings.push(
        `${result.cleaned.length} sheet(s) lost their link definitions in the truncation; the ` +
          'links were removed so the data could be read. No cell values were changed.',
      );
    }
  }

  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(usable as unknown as ArrayBuffer);
  return { workbook, repaired, warnings };
}
