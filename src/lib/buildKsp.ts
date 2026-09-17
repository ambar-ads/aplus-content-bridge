/**
 * Drafts a KSP block from a product's own specification columns.
 *
 * What this is, and what it is not. The KSP already on the site is marketing copy — "Core i7 HX
 * vPro and RTX A2000 drive 3D, AI, BIM workloads" is a claim about who the machine is for, and
 * no amount of reading the spec columns produces that sentence. What it can do is turn a blank
 * cell into a correct, correctly-shaped starting point that a PM edits, instead of a blank cell
 * they have to fill from nothing.
 *
 * The shape follows what is already published: exactly three bullets, each opening with "- ",
 * around 90 characters each. Of the 354 products that have a KSP today, 333 have precisely three.
 *
 * Every value is checked before it is used. The published data is partly scrambled — 21 products
 * carry expansion-slot text in the Security column, 16 carry AC-adapter text there, 21 have
 * "Anti-glare display" where Brightness should be, some carry memory text in GPU, and 73 carry
 * Excel error text such as "#VALUE!" in an ordinary cell. Repeating any of that into a selling
 * point would be worse than leaving the cell empty, so a value that does not look like its own
 * field is skipped and the bullet is built from what remains.
 */
import { isBlankValue } from './legacyCsv';
import type { ForWebRow } from './types';

export const KSP_BULLET_TARGET = 3;

const clean = (v: string) => v.replace(/\s+/g, ' ').trim();

/** Same tidy-up, but newlines survive — several source fields hold one item per line. */
const cleanLines = (v: string) =>
  v
    .split(/\r?\n/)
    .map((l) => l.replace(/[^\S\n]+/g, ' ').trim())
    .filter(Boolean)
    .join('\n');

/** Drop a trailing parenthetical: "RTX A2000 8GB ( 8GB GDDR6 )" -> "RTX A2000 8GB". */
const stripParens = (v: string) => clean(v.replace(/\s*\([^)]*\)/g, ''));

/** The longest published bullet is 202 characters; past that it stops reading as a selling point. */
const MAX_BULLET = 170;

const trimTo = (v: string) => {
  if (v.length <= MAX_BULLET) return v;
  const cut = v.lastIndexOf(' ', MAX_BULLET - 1);
  return v.slice(0, cut > 0 ? cut : MAX_BULLET).replace(/[,;]$/, '');
};

/** Join a short list the way a person writes it. */
const listPhrase = (items: string[]) =>
  items.length <= 1
    ? (items[0] ?? '')
    : items.slice(0, -1).join(', ') + ' and ' + items[items.length - 1];

/**
 * Excel error text that the published data carries in ordinary cells.
 *
 * 73 of the 392 live products have at least one — the workbook's formulas failed against a raw
 * row and the error was exported as if it were a value. A draft that opens "#VALUE! with 8GB
 * DDR4" is worse than an empty cell, so these never count as a value anywhere.
 */
const EXCEL_ERROR = /#(VALUE|REF|NAME|NUM|NULL|DIV\/0|N\/A)[!?]?/i;

const usable = (v: string | undefined, keepLines = false) => {
  const raw = String(v ?? '');
  const s = keepLines ? cleanLines(raw) : clean(raw);
  if (!s || isBlankValue(s) || /^n\/?a$/i.test(s) || EXCEL_ERROR.test(s)) return '';
  return s;
};

/** Memory and storage wording, used to spot a value that has landed in the wrong column. */
const MEMORY_WORDS = /(DIMM|DDR\d|LPDDR|eMMC|\bSSD\b|\bHDD\b|M\.2|NVMe)/i;

/** A value that says what it is, not only how big it is: "512GB PCIe SSD" against "64GB". */
const DESCRIBES_ITSELF = /(SSD|HDD|NVMe|PCIe|SATA|eMMC|DDR|DIMM|M\.2)/i;

/** Each field is only trusted when its value actually looks like that field. */
const looksLike = {
  cpu: (v: string) =>
    /(core|ryzen|athlon|celeron|pentium|xeon|snapdragon|ultra|\bi[3579]-|\bR[357]\b|processor|GHz)/i.test(
      v,
    ),
  // Several products carry memory text in the GPU column, which then reads as "with 8GB DDR4
  // U-DIMM" — true of the machine, but not a graphics card.
  gpu: (v: string) =>
    /(rtx|gtx|geforce|quadro|radeon|vega|iris|arc|nvidia|graphics)/i.test(v) &&
    !MEMORY_WORDS.test(v),
  panelSize: (v: string) => /\d/.test(v),
  resolution: (v: string) => /\d{3,}|FHD|WUXGA|WQXGA|QHD|UHD|2\.?5K|3K|4K/i.test(v),
  brightness: (v: string) => /\d+\s*nits/i.test(v),
  memory: (v: string) => /\d+\s*GB|DDR/i.test(v),
  storage: (v: string) => /\d+\s*(GB|TB)/i.test(v),
  slots: (v: string) => /(slot|M\.2|DIMM|PCIe|SATA|bay)/i.test(v),
  security: (v: string) => /(fingerprint|password|kensington|tpm|smart.?card|lock)/i.test(v),
  weight: (v: string) => /\d\s*kg/i.test(v),
  battery: (v: string) => /\d+\s*wh/i.test(v),
  ports: (v: string) => /(usb|thunderbolt|hdmi|rj-?45|displayport|type-c)/i.test(v),
};

function field(
  row: ForWebRow,
  key: keyof ForWebRow,
  test?: (v: string) => boolean,
  keepLines = false,
): string {
  const v = usable(row[key], keepLines);
  if (!v) return '';
  return !test || test(v) ? v : '';
}

/** "W/FINGERPRINT,SMART CARD" and a paragraph of prose both have to come out readable. */
function securityPhrase(raw: string): string {
  const v = raw.replace(/^W\//i, '').trim();
  const short = v.toUpperCase();
  const hasPrint = /FINGER\s*PRINT/i.test(short);
  const hasCard = /SMART\s*CARD/i.test(short);

  if (hasPrint && hasCard) return 'fingerprint and smart card readers';
  if (hasPrint && short.length < 40) return 'a fingerprint reader';
  if (hasCard && short.length < 40) return 'a smart card reader';

  // Longer values are a list of measures; the first two carry the point.
  const parts = v
    .split(/\n|,(?![^(]*\))/)
    .map(clean)
    .filter(Boolean);
  if (!parts.length) return '';
  return parts
    .slice(0, 2)
    .join(' and ')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

/** Pick the ports worth naming rather than listing all fifteen lines. */
function portsPhrase(raw: string): string {
  const lines = raw.split(/\n/).map(clean).filter(Boolean);
  const wanted = [
    [/thunderbolt/i, 'Thunderbolt'],
    [/usb.?4/i, 'USB4'],
    [/type-?c/i, 'USB-C'],
    [/hdmi/i, 'HDMI'],
    [/rj-?45|ethernet/i, 'RJ-45 Ethernet'],
    [/displayport/i, 'DisplayPort'],
  ] as const;

  const found: string[] = [];
  for (const [re, label] of wanted) {
    if (lines.some((l) => re.test(l)) && !found.includes(label)) found.push(label);
  }
  if (found.length < 2) return '';
  return listPhrase(found.slice(0, 4));
}

function displayBullet(row: ForWebRow): string {
  const size = field(row, 'panelSize', looksLike.panelSize);
  // "WQXGA(WQ) 2560X1600" — the bracketed abbreviation repeats the name before it.
  const res = field(row, 'resolution', looksLike.resolution).replace(/\s*\([A-Z]{1,3}\)/g, '');
  const nits = field(row, 'brightness', looksLike.brightness);
  if (!size && !res) return '';

  const head = [size, res].filter(Boolean).join(' ');
  const touch = /touch/i.test(row.productName) || /touch/i.test(res) ? ' touchscreen' : ' display';
  return trimTo(clean(`${head}${touch}${nits ? ` at ${nits}` : ''}`));
}

function performanceBullet(row: ForWebRow): string {
  const cpu = field(row, 'cpu', looksLike.cpu);
  const gpu = stripParens(field(row, 'gpu', looksLike.gpu));
  const ram = field(row, 'ram', looksLike.memory);
  const ssd = stripParens(field(row, 'ssd', looksLike.storage));
  if (!cpu && !ram && !ssd) return '';

  const core = [cpu, gpu && `with ${gpu}`].filter(Boolean).join(' ');
  const tail = [ram && `${ram} memory`, ssd].filter(Boolean).join(' and ');
  if (core) return trimTo(clean([core, tail].filter(Boolean).join(', ')));

  // No processor survived the checks, so the bullet has to stand on the numbers alone. That is
  // fine for "512GB 2280 PCIe G4 SSD" and useless for "64GB": one names what it is, the other is
  // a size with nothing attached. 90PF0461-M00C20 has every spec column shifted by one and that
  // fragment is all that is left of it.
  if (!ram && !DESCRIBES_ITSELF.test(ssd)) return '';
  return trimTo(clean(tail));
}

function expandabilityBullet(row: ForWebRow): string {
  const slots = field(row, 'expansionSlot', looksLike.slots, true);
  const sec = field(row, 'security', looksLike.security, true);
  if (!slots && !sec) return '';

  const slotText = listPhrase(slots.split(/\n/).map(clean).filter(Boolean).slice(0, 3));
  const secText = sec ? securityPhrase(sec) : '';

  if (slotText && secText) return trimTo(clean(`Expandable with ${slotText}, plus ${secText}`));
  if (slotText) return trimTo(clean(`Expandable with ${slotText}`));
  return trimTo(clean(secText.charAt(0).toUpperCase() + secText.slice(1)));
}

function connectivityBullet(row: ForWebRow): string {
  const ports = field(row, 'ioPort', looksLike.ports, true);
  if (!ports) return '';
  const phrase = portsPhrase(ports);
  return phrase ? trimTo(clean(`Connectivity includes ${phrase}`)) : '';
}

function portabilityBullet(row: ForWebRow): string {
  const weight = field(row, 'weight', looksLike.weight);
  const battery = field(row, 'battery', looksLike.battery);
  if (!weight && !battery) return '';
  if (weight && battery) return trimTo(clean(`${weight} with a ${battery} battery`));
  return trimTo(clean(weight || `${battery} battery`));
}

/**
 * Build the draft.
 *
 * Bullets are tried in order of how much they say about the product, and the first three that
 * produce anything are kept — so a desktop with no panel data still gets three bullets rather
 * than a gap where the display one would have been.
 */
export function draftKsp(row: ForWebRow): string {
  const candidates = [
    displayBullet(row),
    performanceBullet(row),
    expandabilityBullet(row),
    connectivityBullet(row),
    portabilityBullet(row),
  ].filter(Boolean);

  const chosen = candidates.slice(0, KSP_BULLET_TARGET);
  if (!chosen.length) return '';
  return chosen.map((b) => `- ${b}`).join('\n');
}

/** Products whose KSP is empty and for which a draft can actually be built. */
export function draftableRows(rows: ForWebRow[]): ForWebRow[] {
  return rows.filter((r) => isBlankValue(r.ksp) && draftKsp(r) !== '');
}
