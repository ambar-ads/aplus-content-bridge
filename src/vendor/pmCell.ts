/**
 * Fungsi kecil yang disalin-ulang dari repo APLUS utama agar proyek ini berdiri sendiri.
 *
 * Berbeda dengan `pmLayout.ts` yang disalin verbatim oleh syncVendor, file ini ditulis ulang
 * dengan tangan karena fungsi-fungsi aslinya terikat pada dependency besar (specFieldConfig,
 * ~1670 baris) yang tidak dibutuhkan di sini — target kita CSV 25 kolom, bukan model spec baru.
 *
 * Perilakunya harus identik dengan sumbernya. `npm run check:vendor` akan memberi peringatan
 * kalau file sumber di repo utama berubah, supaya file ini ikut ditinjau ulang.
 *
 * Sumber:
 *   sanitize                 -> src/utils/specSource.ts
 *   normalizeLabel           -> src/config/specSchema.ts
 *   productLineFromSheetName -> src/utils/pmDatasheetImport.ts
 */
import type { ProductLineCode } from './types';
import { PM_DATASHEETS } from './pmLayout';

/** Excel leaks `_x000D_` for embedded carriage returns; strip it and normalize newlines. */
export function sanitize(value: unknown): string {
  if (value === null || value === undefined) return '';
  const raw =
    typeof value === 'object' && value !== null && 'richText' in (value as object)
      ? (value as { richText: { text: string }[] }).richText.map((t) => t.text).join('')
      : typeof value === 'object' && value !== null && 'text' in (value as object)
        ? String((value as { text: unknown }).text)
        : typeof value === 'object' && value !== null && 'result' in (value as object)
          ? // A formula cell: PM datasheets pull values across sheets, so take the cached result.
            String((value as { result: unknown }).result ?? '')
          : String(value);
  return raw
    .replace(/_x000D_/g, '')
    .replace(/\r\n?/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

export function normalizeLabel(label: string): string {
  return label.toLowerCase().replace(/[^a-z0-9]/g, '');
}

const KNOWN_SHEETS = new Map(PM_DATASHEETS.map((s) => [s.sheetName.toLowerCase(), s]));

/**
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

/** TKDN adalah katalog terpisah dengan part number yang tidak beririsan — jangan digabung diam-diam. */
export function sheetGroupFromName(sheetName: string): 'main' | 'tkdn' {
  const known = KNOWN_SHEETS.get(sheetName.toLowerCase());
  if (known) return known.group;
  return /tkdn/i.test(sheetName) ? 'tkdn' : 'main';
}
