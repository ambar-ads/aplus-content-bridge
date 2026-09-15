/**
 * Rebuilds a truncated .xlsx so it can be read.
 *
 * `ACID Monthly Pricelist 3 August - 30 August 2026 - ABP.xlsx` arrived from ASUS missing its last
 * ~705 KB. Two things were lost with it: the zip central directory, and every
 * `xl/worksheets/_rels/*.xml.rels`. Excel recovers such a file by walking the local file headers
 * from the start and shrugging off the dangling references, which is why it opens on screen.
 * ExcelJS goes through JSZip, which insists on the central directory, and then trips over the
 * missing relationships when resolving hyperlinks.
 *
 * Everything that matters here survived the cut: all 22 worksheets, the workbook, the shared
 * strings. So rather than asking anyone to re-save the file by hand every month, this walks the
 * same local headers Excel does, keeps whatever is complete, strips the references that now point
 * at nothing, and writes a fresh central directory around the result.
 *
 * Cell data is never re-encoded — each surviving entry's stored bytes, CRC and sizes are copied
 * verbatim. Only worksheet XML that had to be edited is rewritten, and it is written uncompressed
 * so no deflate implementation is needed.
 */

const LOCAL_HEADER_SIG = 0x04034b50;
const CENTRAL_HEADER_SIG = 0x02014b50;
const EOCD_SIG = 0x06054b50;
const EOCD_MIN_SIZE = 22;
/** A zip comment can push the end record back by at most 64 KB. */
const EOCD_SEARCH_WINDOW = 66_000;

const DEFLATED = 8;
const STORED = 0;

export interface ZipEntry {
  name: string;
  method: number;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  dataStart: number;
  complete: boolean;
}

export function hasEndOfCentralDirectory(buf: Uint8Array): boolean {
  const from = Math.max(0, buf.length - EOCD_SEARCH_WINDOW);
  for (let i = buf.length - EOCD_MIN_SIZE; i >= from; i -= 1) {
    if (buf[i] === 0x50 && buf[i + 1] === 0x4b && buf[i + 2] === 0x05 && buf[i + 3] === 0x06) {
      return true;
    }
  }
  return false;
}

/** Walk the local file headers, the way Excel's own recovery does. */
export function readLocalEntries(buf: Uint8Array): ZipEntry[] {
  const view = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const decoder = new TextDecoder('utf-8');
  const entries: ZipEntry[] = [];
  let off = 0;

  while (off + 30 <= buf.length) {
    if (view.getUint32(off, true) !== LOCAL_HEADER_SIG) break;

    const flags = view.getUint16(off + 6, true);
    // With a data descriptor the sizes live after the data, so the stream cannot be walked
    // forwards. OOXML writers do not use it; bail out rather than guess.
    if (flags & 0x08) break;

    const method = view.getUint16(off + 8, true);
    const crc = view.getUint32(off + 14, true);
    const compressedSize = view.getUint32(off + 18, true);
    const uncompressedSize = view.getUint32(off + 22, true);
    const nameLen = view.getUint16(off + 26, true);
    const extraLen = view.getUint16(off + 28, true);

    const name = decoder.decode(buf.subarray(off + 30, off + 30 + nameLen));
    const dataStart = off + 30 + nameLen + extraLen;

    entries.push({
      name,
      method,
      crc,
      compressedSize,
      uncompressedSize,
      dataStart,
      complete: dataStart + compressedSize <= buf.length,
    });

    off = dataStart + compressedSize;
  }

  return entries;
}

// ─── Small zip primitives ───────────────────────────────────────────

let crcTable: Uint32Array | null = null;

function crc32(data: Uint8Array): number {
  if (!crcTable) {
    crcTable = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let c = i;
      for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
      crcTable[i] = c >>> 0;
    }
  }
  let crc = 0xffffffff;
  for (let i = 0; i < data.length; i += 1) {
    crc = (crc >>> 8) ^ crcTable[(crc ^ data[i]) & 0xff];
  }
  return (crc ^ 0xffffffff) >>> 0;
}

/** Uses the platform's own inflate, so no compression library is pulled in for this one case. */
async function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
  const stream = new Blob([data as unknown as BlobPart])
    .stream()
    .pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

/**
 * Remove the parts of a worksheet that point at relationship files.
 *
 * The rels are gone, so hyperlinks, drawings and embedded objects cannot resolve. None of them
 * carry cell values, which is all this tool reads.
 */
function stripDanglingReferences(xml: string): { xml: string; changed: boolean } {
  const before = xml;
  const cleaned = xml
    .replace(/<hyperlinks[\s\S]*?<\/hyperlinks>/g, '')
    .replace(/<hyperlinks\s*\/>/g, '')
    .replace(/<drawing\b[^>]*\/>/g, '')
    .replace(/<legacyDrawing\b[^>]*\/>/g, '')
    .replace(/<legacyDrawingHF\b[^>]*\/>/g, '')
    .replace(/<picture\b[^>]*\/>/g, '')
    .replace(/<oleObjects[\s\S]*?<\/oleObjects>/g, '')
    .replace(/<controls[\s\S]*?<\/controls>/g, '')
    .replace(/<tableParts[\s\S]*?<\/tableParts>/g, '')
    .replace(/<tableParts\s*\/>/g, '');
  return { xml: cleaned, changed: cleaned !== before };
}

export interface RepairResult {
  buffer: Uint8Array;
  /** Entries left out because their data ran past the end of the file. */
  dropped: string[];
  /** Worksheets whose dangling hyperlink/drawing references had to be removed. */
  cleaned: string[];
  kept: number;
}

function writeUint32(out: Uint8Array, at: number, value: number) {
  out[at] = value & 0xff;
  out[at + 1] = (value >>> 8) & 0xff;
  out[at + 2] = (value >>> 16) & 0xff;
  out[at + 3] = (value >>> 24) & 0xff;
}

function writeUint16(out: Uint8Array, at: number, value: number) {
  out[at] = value & 0xff;
  out[at + 1] = (value >>> 8) & 0xff;
}

interface OutEntry {
  name: Uint8Array;
  method: number;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  data: Uint8Array;
}

/**
 * Write a fresh zip containing every complete entry.
 *
 * Extra fields are dropped: they hold timestamps and Zip64 hints that nothing here needs, and
 * leaving them out keeps the offset arithmetic simple.
 */
export async function rebuildZip(buf: Uint8Array): Promise<RepairResult> {
  const entries = readLocalEntries(buf);
  const complete = entries.filter((e) => e.complete);
  const dropped = entries.filter((e) => !e.complete).map((e) => e.name);
  const cleaned: string[] = [];

  if (!complete.length) {
    throw new Error('No readable entries were found in this file.');
  }

  const present = new Set(complete.map((e) => e.name));
  const encoder = new TextEncoder();
  const decoder = new TextDecoder('utf-8');
  const out: OutEntry[] = [];

  for (const entry of complete) {
    const raw = buf.subarray(entry.dataStart, entry.dataStart + entry.compressedSize);
    const isWorksheet = /^xl\/worksheets\/sheet\d+\.xml$/.test(entry.name);
    const relsName = entry.name.replace('xl/worksheets/', 'xl/worksheets/_rels/') + '.rels';

    // Only worksheets that lost their relationship file need rewriting; everything else is
    // copied through untouched.
    if (isWorksheet && !present.has(relsName)) {
      try {
        const xml =
          entry.method === DEFLATED ? decoder.decode(await inflateRaw(raw)) : decoder.decode(raw);
        const result = stripDanglingReferences(xml);
        if (result.changed) {
          const bytes = encoder.encode(result.xml);
          out.push({
            name: encoder.encode(entry.name),
            method: STORED,
            crc: crc32(bytes),
            compressedSize: bytes.length,
            uncompressedSize: bytes.length,
            data: bytes,
          });
          cleaned.push(entry.name);
          continue;
        }
      } catch {
        // Could not decode it — fall through and copy the original bytes.
      }
    }

    out.push({
      name: encoder.encode(entry.name),
      method: entry.method,
      crc: entry.crc,
      compressedSize: entry.compressedSize,
      uncompressedSize: entry.uncompressedSize,
      data: raw,
    });
  }

  const localSize = out.reduce((n, e) => n + 30 + e.name.length + e.data.length, 0);
  const centralSize = out.reduce((n, e) => n + 46 + e.name.length, 0);
  const zip = new Uint8Array(localSize + centralSize + EOCD_MIN_SIZE);

  const offsets: number[] = [];
  let at = 0;

  for (const entry of out) {
    offsets.push(at);
    writeUint32(zip, at, LOCAL_HEADER_SIG);
    writeUint16(zip, at + 4, 20);
    writeUint16(zip, at + 6, 0);
    writeUint16(zip, at + 8, entry.method);
    writeUint16(zip, at + 10, 0);
    writeUint16(zip, at + 12, 0);
    writeUint32(zip, at + 14, entry.crc);
    writeUint32(zip, at + 18, entry.compressedSize);
    writeUint32(zip, at + 22, entry.uncompressedSize);
    writeUint16(zip, at + 26, entry.name.length);
    writeUint16(zip, at + 28, 0);
    zip.set(entry.name, at + 30);
    zip.set(entry.data, at + 30 + entry.name.length);
    at += 30 + entry.name.length + entry.data.length;
  }

  const centralStart = at;

  out.forEach((entry, i) => {
    writeUint32(zip, at, CENTRAL_HEADER_SIG);
    writeUint16(zip, at + 4, 20);
    writeUint16(zip, at + 6, 20);
    writeUint16(zip, at + 8, 0);
    writeUint16(zip, at + 10, entry.method);
    writeUint16(zip, at + 12, 0);
    writeUint16(zip, at + 14, 0);
    writeUint32(zip, at + 16, entry.crc);
    writeUint32(zip, at + 20, entry.compressedSize);
    writeUint32(zip, at + 24, entry.uncompressedSize);
    writeUint16(zip, at + 28, entry.name.length);
    writeUint16(zip, at + 30, 0);
    writeUint16(zip, at + 32, 0);
    writeUint16(zip, at + 34, 0);
    writeUint16(zip, at + 36, 0);
    writeUint32(zip, at + 38, 0);
    writeUint32(zip, at + 42, offsets[i]);
    zip.set(entry.name, at + 46);
    at += 46 + entry.name.length;
  });

  writeUint32(zip, at, EOCD_SIG);
  writeUint16(zip, at + 4, 0);
  writeUint16(zip, at + 6, 0);
  writeUint16(zip, at + 8, out.length);
  writeUint16(zip, at + 10, out.length);
  writeUint32(zip, at + 12, at - centralStart);
  writeUint32(zip, at + 16, centralStart);
  writeUint16(zip, at + 20, 0);

  return { buffer: zip, dropped, cleaned, kept: out.length };
}
