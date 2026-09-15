/**
 * Memastikan sebuah buffer benar-benar file .xlsx sebelum diserahkan ke ExcelJS.
 *
 * Tanpa ini, memilih file yang salah menghasilkan pesan zip mentah yang tidak berarti apa-apa
 * bagi marketing. Bukan kasus teoretis: riwayat commit repo lama menunjukkan file salah pernah
 * beberapa kali terunggah. Pemotongan berkas ditangani terpisah oleh `openWorkbook`, yang
 * memulihkannya alih-alih menolaknya.
 */

/** Semua file Office modern adalah arsip ZIP, jadi diawali "PK\x03\x04". */
const ZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04];
/** File .xls lama (BIFF/OLE) diawali D0 CF 11 E0 — ExcelJS tidak bisa membacanya. */
const OLE_MAGIC = [0xd0, 0xcf, 0x11, 0xe0];

function startsWith(bytes: Uint8Array, magic: number[]): boolean {
  return magic.every((b, i) => bytes[i] === b);
}

export class XlsxError extends Error {}

export function assertXlsx(buffer: ArrayBuffer, fileName: string): void {
  const head = new Uint8Array(buffer.slice(0, 4));

  if (startsWith(head, OLE_MAGIC)) {
    throw new XlsxError(
      `"${fileName}" is still in the old .xls format. Open it in Excel and save it again as .xlsx.`,
    );
  }

  if (!startsWith(head, ZIP_MAGIC)) {
    throw new XlsxError(
      `"${fileName}" is not a valid .xlsx file (${buffer.byteLength.toLocaleString('en')} bytes). ` +
        'Check that you picked the Excel file, not a CSV or a saved web page.',
    );
  }

  // A missing End of Central Directory used to be rejected here. It is no longer fatal:
  // `openWorkbook` rebuilds truncated archives the same way Excel recovers them, so a file cut
  // short in transit still gets read instead of being handed back to the user as unusable.
}
