/**
 * Menirukan formula sheet `Auto-NX` / `Auto-AIO` / `Auto-PF` dari workbook lama.
 *
 * Sheet-sheet itulah "backend" sistem lama yang sebenarnya: mereka mengubah datasheet PM mentah
 * menjadi teks siap tayang. Modul ini menuliskan ulang aturannya dalam TypeScript, dibaca
 * langsung dari formula di baris 2 tiap sheet.
 *
 * PENTING — modul ini hanya dipakai untuk produk BARU. 392 baris yang sudah terbit dipertahankan
 * apa adanya oleh `catalogStore`, supaya satu aturan yang keliru di sini tidak bisa mengubah teks
 * yang sudah dipakai marketing di listing e-commerce.
 *
 * Kejanggalan sistem lama sengaja DIPERTAHANKAN, bukan diperbaiki, supaya produk baru konsisten
 * dengan yang sudah tayang. Semuanya terdaftar di `LEGACY_QUIRKS` di bawah.
 */
import { normalizeLabel } from '@/vendor/pmCell';
import { EMPTY_FORWEB_ROW } from './config';
import type { ForWebRow } from './types';

export type PmLine = 'NX' | 'PT' | 'PF';

/**
 * Kejanggalan yang sengaja ditiru, bukan diperbaiki.
 *
 * Memperbaikinya di sini akan membuat produk baru tampil berbeda dari ratusan produk yang sudah
 * tayang, dan membuat `check:forweb` gagal. Perbaikannya harus jadi keputusan tersendiri yang
 * diterapkan ke seluruh katalog sekaligus.
 */
export const LEGACY_QUIRKS = [
  {
    id: 'pf-io-swapped',
    affects: 'Desktop (90PF)',
    detail:
      'The Auto-PF sheet reads "Rear I/O Port" into the column labelled "Front:" and ' +
      '"Front I/O Port" into the one labelled "Rear:", so desktop front and rear ports are ' +
      'shown the wrong way round on the site.',
  },
  {
    id: 'pf-panel-hardcoded',
    affects: 'Desktop (90PF)',
    detail:
      'Panel Size, Resolution and Brightness for desktops are hard-coded to 21.5", ' +
      'FHD (1920 x 1080) 16:9 and 250nits rather than read from the datasheet.',
  },
  {
    id: 'mkt-name-not-found',
    affects: 'Mostly Desktop (90PF)',
    detail:
      'When the marketing name lookup failed, the old system still wrote the literal text ' +
      '"Not found" and carried it into Title. 92 of the 392 live products are affected.',
  },
] as const;

/**
 * Tiga tempat di mana modul ini SENGAJA berbeda dari sheet Auto-*.
 *
 * Batasannya begini: kejanggalan kosmetik ditiru (agar produk baru seragam dengan yang tayang),
 * tetapi formula yang menghasilkan nilai jelas-jelas rusak tidak ditiru — menyalin kerusakan ke
 * produk baru tidak ada gunanya. Ketiganya hanya berlaku untuk produk baru; 392 baris yang sudah
 * terbit tidak tersentuh.
 */
export const DELIBERATE_FIXES = [
  {
    id: 'panel-size-not-found',
    detail:
      'Auto-NX ran INT() on the panel size. ASUS now writes it as text ("14.0-inch"), so INT() ' +
      'errors and produces "Not Found" — 67 of the 139 live notebooks look like that today. ' +
      'Here the number is read out instead, so both 14" and 15.6" come out right.',
  },
  {
    id: 'renamed-source-rows',
    detail:
      'ASUS renames datasheet rows between price list versions: GPU to Integrated GPU, On board ' +
      'memory to Total System Memory, Security Reader to FingerPrint, Weight (with Battery) to ' +
      'Starting Weight (with Battery). Every spelling is accepted, so a new price list cannot ' +
      'quietly produce an empty column.',
  },
  {
    id: 'row-layout-varies',
    detail:
      'Datasheet rows are located by searching column A for "Part No" and "Sales Model Name" ' +
      'rather than assuming fixed rows. One ACID workbook mixes three layouts — rows 1/2 for ' +
      'Desktop and All-in-One, 1/8 for Commercial Notebook, 2/9 for TKDN — and assuming the ABP ' +
      'positions silently skipped six of its nine datasheets.',
  },
] as const;

// ─── Padanan fungsi Excel ────────────────────────────────────────────

/**
 * TRIM ala Excel: rapatkan spasi ganda dan buang spasi tepi, TANPA menyentuh newline.
 *
 * Excel TRIM hanya mengurus spasi; CHAR(10) dibiarkan. Memakai `\s+` di sini akan meratakan
 * nilai multi-baris menjadi satu baris — persis yang sempat terjadi pada kolom CPU untuk
 * prosesor Ryzen AI yang memang ditulis dua baris.
 */
function xlTrim(s: string): string {
  return s.replace(/[^\S\n]+/g, ' ').replace(/^\s+|\s+$/g, '');
}

function textBefore(s: string, delim: string): string | null {
  const i = s.indexOf(delim);
  return i === -1 ? null : s.slice(0, i);
}

function textAfter(s: string, delim: string): string | null {
  const i = s.indexOf(delim);
  return i === -1 ? null : s.slice(i + delim.length);
}

// ─── Pipeline CPU ────────────────────────────────────────────────────

/**
 * `CPU Normal` — buang kurung pertama beserta isinya dan kata "Processor".
 * PF juga membuang newline dan teks "N/A" lebih dulu.
 */
export function cpuNormal(processorFull: string, line: PmLine): string {
  let v = processorFull;
  if (line === 'PF') v = v.replace(/\n/g, '');

  const open = v.indexOf('(');
  const close = v.indexOf(')');
  if (open !== -1 && close > open) v = v.slice(0, open) + v.slice(close + 1);

  v = v.replace(/Processor/g, '');
  if (line === 'PF') v = v.replace(/N\/A/g, '');
  return xlTrim(v);
}

/** `CPU WIP` — teks sebelum "Processor" ditambah satu token setelahnya (biasanya clock speed). */
export function cpuWip(processorFull: string): string {
  const idx = processorFull.toLowerCase().indexOf('processor');
  if (idx === -1) return xlTrim(processorFull);

  const head = processorFull.slice(0, idx);
  const tail = processorFull.slice(idx + 10, idx + 10 + 20);
  const space = tail.indexOf(' ');
  const token = space === -1 ? tail : tail.slice(0, space);
  return head + token;
}

const CPU_FAMILY_MARKERS: [RegExp, string][] = [
  [/Core™ Ultra 7/, 'U7'],
  [/Core™ Ultra 5/, 'U5'],
  [/Core™ Ultra 3/, 'U3'],
  [/Core™ 7/, 'Core 7'],
  [/Core™ 5/, 'Core 5'],
  [/Core™ 3/, 'Core 3'],
  [/Ryzen™ 7/, 'R7'],
  [/Ryzen™ 5/, 'R5'],
  [/Ryzen™ 3/, 'R3'],
];

/** `CPU Short` — "i7-1355U", atau penanda keluarga seperti U7 / Core 5 / R5. */
export function cpuShort(normal: string, wip: string): string {
  if (normal.includes('Core™ i')) {
    const after = textAfter(normal, 'Core™ ');
    if (after !== null) {
      const token = textBefore(after, ' ');
      if (token !== null) return token;
    }
  }
  for (const [re, code] of CPU_FAMILY_MARKERS) {
    if (re.test(normal)) return code;
  }
  return wip;
}

const NEEDS_SPEED = new Set(['U3', 'U5', 'U7', 'Core 3', 'Core 5', 'Core 7']);
const NEEDS_MODEL = new Set(['R3', 'R5', 'R7']);

/** `CPU Final` — penanda keluarga masih perlu ditempeli angka dari `CPU WIP`. */
export function cpuFinal(short: string, wip: string): string {
  const words = wip.trim().split(/\s+/).filter(Boolean);
  if (NEEDS_SPEED.has(short)) {
    const last = words[words.length - 1] ?? '';
    return `${short} ${last}`.trim();
  }
  if (NEEDS_MODEL.has(short)) {
    const secondLast = words[words.length - 2] ?? '';
    return `${short} ${secondLast}`.trim();
  }
  return short;
}

/** `GPU (if good)` — hanya GPU diskrit RTX yang ikut masuk Title, dengan garis miring di depan. */
export function gpuIfGood(gpu: string): string {
  const i = gpu.indexOf('RTX');
  if (i === -1) return '';
  const window = gpu.slice(i, i + 15);
  const space = window.indexOf(' ', 5);
  return '/' + (space === -1 ? window : window.slice(0, space));
}

export function chassisOf(modelName: string): string {
  return textBefore(modelName, '-') ?? modelName;
}

/**
 * Tukar bagian nama di depan Title, pertahankan bagian spesifikasinya apa adanya.
 *
 * Title berbentuk `{Marketing Name} ({CPU}/{RAM}/{SSD}/{OS}/{Chassis})`. Saat seseorang
 * memperbaiki Marketing Name, Title-nya ikut perlu diperbaiki — kalau tidak, teks "Not found"
 * tetap tayang di sana. Menyusun ulang seluruh Title akan butuh nilai mentah datasheet yang sudah
 * tidak tersimpan di baris ini, jadi yang ditukar hanya bagian namanya. Bagian dalam kurung tidak
 * disentuh sama sekali.
 */
export function retitleWith(title: string, marketingName: string): string {
  const i = title.lastIndexOf(' (');
  return i === -1 ? marketingName : marketingName + title.slice(i);
}

/** Apakah Title masih memakai nama lain daripada Marketing Name yang tertulis sekarang? */
export function titleNeedsRetitle(title: string, marketingName: string): boolean {
  if (!title.trim() || !marketingName.trim()) return false;
  return title !== retitleWith(title, marketingName);
}

export function osShortOf(os: string): string {
  return os.replace(/Windows/g, 'Win');
}

/** OS dinormalkan ke tiga nilai saja; selain itu dikosongkan. */
export function normalizeOs(raw: string): string {
  if (raw.includes('Windows 11 Home')) return 'Windows 11 Home';
  if (raw.includes('Windows 11 Pro')) return 'Windows 11 Pro';
  if (raw.includes('Chrome')) return 'ChromeOS';
  return '';
}

// ─── Pemetaan label datasheet PM per product line ────────────────────

/**
 * Label kolom A datasheet PM yang dipakai tiap product line.
 *
 * Huruf kolom pada formula Excel sudah diterjemahkan kembali ke label aslinya, sehingga peta ini
 * tetap berlaku walau ASUS menggeser urutan kolom di price list berikutnya.
 */
const SOURCE_LABELS: Record<PmLine, Record<string, string | string[]>> = {
  NX: {
    partNo: 'Part No',
    modelName: 'Sales Model Name',
    processor: 'Processor',
    // ASUS mengosongkan "On board memory" pada model-model baru dan memindahkannya ke
    // "Total System Memory". Kedua ejaan diterima supaya price list lama dan baru sama-sama jalan.
    ram: ['On board memory', 'Total System Memory'],
    storage: 'Storage',
    gpu: ['GPU', 'Integrated GPU'],
    os: 'Operating System',
    office: 'Office',
    expansionSlot: 'Expansion Slot(includes used)',
    ioPorts: 'I/O ports',
    includedInBox: 'Included in the Box',
    panelSize: 'Panel Size',
    brightness: 'Brightness',
    resolution: 'Resolution',
    // The ACID monthly price list calls this "Starting Weight (with Battery)"; the ABP one just
    // "Weight (with Battery)". Both spellings appear in files still in use.
    weight: ['Weight (with Battery)', 'Starting Weight (with Battery)'],
    dimension: 'Dimension (WxHxD)',
    battery: 'Battery',
    acAdapter: 'AC Adapter',
    // Baris ini dulu bernama "Security Reader" dan sekarang "FingerPrint" — isinya kode
    // konfigurasi pendek seperti "W/FINGERPRINT,SMART CARD". Jangan tertukar dengan baris
    // "Security", yang berisi paragraf panjang dan bukan yang dipakai halaman lama.
    security: ['Security Reader', 'FingerPrint'],
  },
  PT: {
    partNo: 'Part No',
    modelName: 'Sales Model Name',
    processor: 'On board processor',
    ram: 'Total System Memory',
    storage: 'Storage',
    gpu: 'Integrated GPU',
    os: 'Operating System',
    office: 'Office',
    expansionSlot: 'Expansion Slot(includes used)',
    includedInBox: 'Included in the box',
    panelSize: 'Panel Size',
    brightness: 'Brightness',
    resolution: 'Resolution',
    weight: 'Weight',
    dimension: 'Dimension (WxHxD)',
    acAdapter: 'AC Adapter',
    security: 'Security',
    sidePort: 'Side I/O Port',
    backPort: 'Back I/O Port',
  },
  PF: {
    partNo: 'Part No',
    modelName: 'Sales Model Name',
    processor: 'Processor',
    ram: 'DIMM Memory',
    storage: 'Storage',
    graphics: 'Graphics',
    integratedGpu: 'Integrated GPU',
    os: 'Operating System',
    office: 'Office',
    expansionSlot: 'Expansion Slot(includes used)',
    includedInBox: 'Keyboard & Mouse',
    weight: 'Weight',
    dimension: 'Dimension (WxDxH)',
    powerSupply: 'Power Supply',
    security: 'Security',
    formFactor: 'Form Factor',
    // Sengaja tertukar — lihat LEGACY_QUIRKS 'pf-io-swapped'.
    frontPortSource: 'Rear I/O Port',
    rearPortSource: 'Front I/O Port',
  },
};

export type PmRecord = Map<string, string>;

/** Susun record datasheet PM (label -> nilai) dengan pencarian label yang tahan beda ejaan. */
export function makePmRecord(pairs: [string, string][]): PmRecord {
  const m = new Map<string, string>();
  for (const [label, value] of pairs) {
    const key = normalizeLabel(label);
    if (key && !m.has(key)) m.set(key, value);
  }
  return m;
}

/** Ambil nilai label pertama yang ada isinya. Alias dipakai karena ASUS mengubah ejaan baris. */
function get(rec: PmRecord, label: string | string[] | undefined): string {
  if (!label) return '';
  const candidates = Array.isArray(label) ? label : [label];
  for (const c of candidates) {
    const v = rec.get(normalizeLabel(c));
    if (v && v.trim()) return v;
  }
  return '';
}

/**
 * Apakah teks ini benar-benar 90PN, mis. `90NX04U1-M004P0`.
 *
 * Kedua workbook sumber punya baris sisa yang menghasilkan teks label, error formula, atau objek
 * ExcelJS di posisi part number. Tanpa penyaring ini semuanya akan masuk sebagai "produk".
 */
export function isPn90(value: string): boolean {
  return /^90[A-Z]{2}/i.test(value.trim());
}

/** 90PN menentukan product line — inilah yang dipakai formula FILTER di sheet Auto-*. */
export function productLineFromPn90(pn90: string): PmLine | null {
  const p = pn90.slice(0, 4).toUpperCase();
  if (p === '90NX') return 'NX';
  if (p === '90PT') return 'PT';
  if (p === '90PF') return 'PF';
  return null;
}

export interface DeriveOptions {
  /** Peta dari sheet "MKT name rule": kunci ternormalisasi -> marketing name. */
  mktNameRule: Map<string, string>;
}

export interface DeriveResult {
  row: ForWebRow;
  warnings: string[];
}

function lookupMktName(
  rule: Map<string, string>,
  key: string,
): { value: string; found: boolean } {
  const hit = rule.get(normalizeLabel(key));
  // Sistem lama menulis literal "Not found" saat lookup gagal, dan teks itu ikut tayang.
  return hit ? { value: hit, found: true } : { value: 'Not found', found: false };
}

/** Ubah satu kolom model dari datasheet PM menjadi satu baris ForWeb. */
export function deriveForWebRow(rec: PmRecord, options: DeriveOptions): DeriveResult {
  const warnings: string[] = [];
  const partNo = get(rec, 'Part No').trim();
  const line = productLineFromPn90(partNo);

  if (!line) {
    return {
      row: { ...EMPTY_FORWEB_ROW, partNumber: partNo } as ForWebRow,
      warnings: [`90PN "${partNo}" is not recognised (not 90NX/90PT/90PF) — cannot be derived`],
    };
  }

  const L = SOURCE_LABELS[line];
  const modelName = get(rec, L.modelName).trim();
  const processorFull = get(rec, L.processor);
  const chassis = chassisOf(modelName);

  // Marketing name: NX/PT pakai 5 huruf pertama model name, PF pakai 3 huruf pertama chassis.
  let mkt: string;
  if (line === 'PF') {
    const base = lookupMktName(options.mktNameRule, chassis.slice(0, 3));
    const suffix = chassis.startsWith('S') ? chassis : get(rec, L.formFactor).trim();
    mkt = `${base.value} ${suffix}`.trim();
    if (!base.found) warnings.push(`${modelName}: no marketing name for series key "${chassis.slice(0, 3)}"`);
  } else {
    const base = lookupMktName(options.mktNameRule, modelName.slice(0, 5));
    mkt = base.value;
    if (!base.found) warnings.push(`${modelName}: no marketing name for series key "${modelName.slice(0, 5)}"`);
  }

  const normal = cpuNormal(processorFull, line);
  const wip = cpuWip(processorFull);
  const short = cpuShort(normal, wip);
  const finalCpu = cpuFinal(short, wip);

  const gpu =
    line === 'PF'
      ? (() => {
          const g = get(rec, L.graphics).trim();
          return !g || g === '0' ? get(rec, L.integratedGpu) : g;
        })()
      : get(rec, L.gpu);

  const os = normalizeOs(get(rec, L.os));

  let ram = get(rec, L.ram);
  if (line === 'NX') {
    const paren = ram.indexOf('(');
    if (paren !== -1) ram = xlTrim(ram.slice(0, paren));
  }

  const ssd = get(rec, L.storage);

  // Panel: NX membaca angka lalu menambahkan tanda inci; PT mengganti "-inch"; PF di-hardcode.
  let panelSize = '';
  let resolution = '';
  let brightness = '';
  if (line === 'NX') {
    panelSize = formatPanelSize(get(rec, L.panelSize));
    resolution = get(rec, L.resolution);
    brightness = get(rec, L.brightness);
  } else if (line === 'PT') {
    panelSize = get(rec, L.panelSize).replace(/-inch/g, '"');
    resolution = get(rec, L.resolution);
    brightness = get(rec, L.brightness);
  } else {
    panelSize = '21.5"';
    resolution = 'FHD (1920 x 1080) 16:9';
    brightness = '250nits';
  }

  // IO Port: NX satu blok; PT "Side:/Back:"; PF "Front:/Rear:" dengan sumber yang tertukar.
  let ioPort = '';
  if (line === 'NX') {
    ioPort = get(rec, L.ioPorts);
  } else {
    const [aLabel, bLabel, aTitle, bTitle] =
      line === 'PT'
        ? [L.sidePort, L.backPort, 'Side', 'Back']
        : [L.frontPortSource, L.rearPortSource, 'Front', 'Rear'];
    const a = get(rec, aLabel).trim();
    const b = get(rec, bLabel).trim();
    if (a || b) {
      ioPort = (a ? `${aTitle}:\n${a}\n\n` : '') + (b ? `${bTitle}:\n${b}` : '');
    }
  }

  // NX meneruskan berat apa adanya — termasuk keterangan lbs dalam kurung, yang memang ikut
  // tayang untuk notebook ("1.40 kg (3.09 lbs)"). Hanya PT/PF yang memotongnya.
  const weightRaw = get(rec, L.weight);
  const weight = line === 'NX' ? weightRaw : xlTrim(textBefore(weightRaw, ' (') ?? weightRaw);

  const dimension = get(rec, L.dimension);
  const size = xlTrim(textBefore(dimension, ' (') ?? dimension);

  let battery = 'No Battery';
  if (line === 'NX') {
    const b = get(rec, L.battery);
    battery = xlTrim(textBefore(b, ',') ?? b);
  }

  let powerSupply: string;
  if (line === 'PF') {
    powerSupply = textBefore(get(rec, L.powerSupply), ' power') ?? 'Not Found';
  } else {
    powerSupply = textBefore(get(rec, L.acAdapter), ', Output:') ?? 'Not Found';
  }

  let office = get(rec, L.office);
  if (line === 'PF') office = office.includes('Office') ? office : 'Not Included';

  let includedInBox = get(rec, L.includedInBox);
  if (line === 'NX' && !includedInBox.trim()) includedInBox = 'Not included';

  const securityRaw = get(rec, L.security);
  const security = securityRaw === 'N/A' ? '' : securityRaw;

  const title = `${mkt} (${finalCpu}/${ramShortOf(ram, line)}/${ssdShortOf(ssd)}${gpuIfGood(gpu)}/${osShortOf(os)}/${chassis})`;

  const row: ForWebRow = {
    ...EMPTY_FORWEB_ROW,
    partNumber: partNo,
    productName: modelName,
    marketingName: mkt,
    // Tagline / KSP / Link tidak pernah ada di datasheet PM — diisi dari workbook lama.
    tagline: '',
    ksp: '',
    link: '',
    partNumberAuto: partNo,
    title,
    cpu: normal,
    gpu,
    os,
    office,
    panelSize,
    resolution,
    brightness,
    ram,
    ssd,
    ioPort,
    weight,
    size,
    expansionSlot: get(rec, L.expansionSlot),
    includeInTheBox: includedInBox,
    battery,
    powerSupply,
    security,
  } as ForWebRow;

  return { row, warnings };
}

/**
 * Ukuran panel jadi angka + tanda inci: "14.0-inch" dan 14 sama-sama jadi `14"`, "15.6-inch"
 * jadi `15.6"`.
 *
 * Sheet Auto-NX memakai `INT()`, yang error pada teks seperti "14.0-inch" sehingga menghasilkan
 * "Not Found" — saat ini 67 dari 139 notebook yang tayang terkena itu. `INT()` juga akan
 * memotong 15.6 menjadi 15. Lihat DELIBERATE_FIXES.
 */
function formatPanelSize(raw: string): string {
  const m = raw.match(/\d+(?:\.\d+)?/);
  if (!m) return raw.trim() ? raw.trim() : '';
  const n = Number(m[0]);
  if (!Number.isFinite(n)) return '';
  return `${n}"`;
}

/** Bagian RAM yang masuk Title: NX/PT ambil setelah spasi pertama, PF sebelum spasi pertama. */
function ramShortOf(ram: string, line: PmLine): string {
  const i = ram.indexOf(' ');
  if (i === -1) return ram;
  return line === 'PF' ? ram.slice(0, i) : xlTrim(ram.slice(i + 1));
}

function ssdShortOf(ssd: string): string {
  const i = ssd.indexOf(' ');
  return i === -1 ? ssd : ssd.slice(0, i);
}
