import type { ProductLineCode, SpecFieldDefinition } from '@/types/specForm';
import { readOverlay, applyOverride, withoutRemoved } from '@/services/masterDataService';

/**
 * Declared category table.
 *
 * Previously category layout was an *accident of array ordering*: `groupFieldsByCategory`
 * opened a new group whenever a field's `category` string differed from the previous
 * field's, so a category appearing twice non-adjacently silently rendered as two blocks and
 * adding a field meant inserting it at exactly the right physical index.
 *
 * Categories are now declared here with an explicit `order` and `section`. Fields resolve to
 * a category id (via `categoryId`, or by matching their legacy `category` label), and
 * adjacency in `SPEC_FIELDS` no longer matters.
 */

/** The three collapsible blocks rendered by ProductSpec.tsx. */
export type SpecSection = 'spec' | 'warranty' | 'sustainable';

/**
 * `merge` renders the category label spanning the Category + Item columns (only the value
 * shows) - the existing `ProductSpecCategory.mergeCategory` behaviour. Omit to let
 * `resolveCategoryLayout` decide per product line: a category left with a single field for
 * that line merges, otherwise it pairs two label/value sets per row.
 */
export type SpecCategoryLayout = 'pair' | 'merge';

export interface SpecCategoryDef {
  id: string;
  label: string;
  section: SpecSection;
  order: number;
  layout?: SpecCategoryLayout;
  /** Omit to allow every product line. */
  productLines?: ProductLineCode[];
}

/** Catch-all for imported columns we do not have a field for yet. */
export const OTHER_CATEGORY_ID = 'other';

export const SPEC_CATEGORIES: SpecCategoryDef[] = [
  { id: 'eanCode', label: 'EAN Code', section: 'spec', order: 10 },
  { id: 'upcCode', label: 'UPC Code', section: 'spec', order: 20 },
  { id: 'operatingSystem', label: 'Operating System', section: 'spec', order: 30 },
  { id: 'colorMaterial', label: 'Color & Material', section: 'spec', order: 40 },
  { id: 'processor', label: 'Processor', section: 'spec', order: 50 },
  { id: 'display', label: 'Display', section: 'spec', order: 60 },
  { id: 'displayDetail', label: 'Color & Material (Display)', section: 'spec', order: 70 },
  { id: 'ram', label: 'RAM', section: 'spec', order: 80 },
  { id: 'storage', label: 'Storage', section: 'spec', order: 90 },
  { id: 'dimensionNotebook', label: 'Dimension (WxHxD)', section: 'spec', order: 100 },
  { id: 'dimensionDesktop', label: 'Dimension (WxDxH)', section: 'spec', order: 110 },
  { id: 'militaryGrade', label: 'Military Grade', section: 'spec', order: 120 },
  { id: 'weight', label: 'Weight', section: 'spec', order: 130 },
  { id: 'ioPorts', label: 'I/O Ports', section: 'spec', order: 140 },
  { id: 'battery', label: 'Battery', section: 'spec', order: 150 },
  { id: 'acAdapter', label: 'AC Adapter', section: 'spec', order: 160 },
  { id: 'requiredChargingPower', label: 'Required Charging Power', section: 'spec', order: 170 },
  { id: 'powerSupply', label: 'Power Supply', section: 'spec', order: 180 },
  { id: 'desktopInfo', label: 'Desktop Info', section: 'spec', order: 190 },
  { id: 'aioInfo', label: 'AiO Info', section: 'spec', order: 200 },
  { id: 'keyboard', label: 'Keyboard Type', section: 'spec', order: 210 },
  { id: 'keyboardMouse', label: 'Keyboard & Mouse', section: 'spec', order: 220 },
  { id: 'camera', label: 'Front-Facing Camera', section: 'spec', order: 230 },
  { id: 'network', label: 'Network and Communication', section: 'spec', order: 240 },
  { id: 'expansionSlots', label: 'Expansion Slots', section: 'spec', order: 250 },
  { id: 'builtInApps', label: 'Built-in Apps', section: 'spec', order: 260 },
  { id: 'builtInAppsDisclaimer', label: 'Disclaimer for Built-in Apps', section: 'spec', order: 270 },
  { id: 'myAsusFeatures', label: 'MyASUS Features', section: 'spec', order: 280 },
  { id: 'myExpertFeatures', label: 'MyExpert Features', section: 'spec', order: 290 },
  { id: 'adobeCreativeCloud', label: 'Adobe Creative Cloud Hard Bundle', section: 'spec', order: 300 },
  { id: 'audio', label: 'Audio', section: 'spec', order: 310 },
  { id: 'antivirus', label: 'Antivirus', section: 'spec', order: 320 },
  { id: 'security', label: 'Security', section: 'spec', order: 330 },
  { id: 'includedInBox', label: 'Included in the Box (Optional)', section: 'spec', order: 340 },
  { id: 'accessory', label: 'Accessory', section: 'spec', order: 350 },
  { id: 'repairability', label: 'Repairability Index', section: 'spec', order: 360 },
  { id: 'ecolabels', label: 'Ecolabels & Compliances', section: 'spec', order: 370 },
  { id: OTHER_CATEGORY_ID, label: 'Other Specifications', section: 'spec', order: 900 },

  { id: 'baseWarranty', label: 'Base Warranty', section: 'warranty', order: 10 },
  { id: 'recommendedWarranty', label: 'Recommended Warranty Package', section: 'warranty', order: 20 },
  { id: 'batteryWarranty', label: 'Battery Warranty', section: 'warranty', order: 30 },

  { id: 'sustainableMaterial', label: 'Sustainable Material', section: 'sustainable', order: 10 },
];

const CATEGORY_BY_ID = new Map(SPEC_CATEGORIES.map((c) => [c.id, c]));

/** Normalize a label for tolerant matching: lowercase, alphanumerics only. */
export function normalizeLabel(label: string): string {
  return label.toLowerCase().replace(/[^a-z0-9]/g, '');
}

/** Legacy `category` display strings -> category id, so existing fields resolve unchanged. */
const CATEGORY_BY_NORMALIZED_LABEL = new Map(
  SPEC_CATEGORIES.map((c) => [normalizeLabel(c.label), c.id]),
);

/**
 * The categories as an administrator has them, ordered.
 *
 * `CATEGORY_BY_ID` and `CATEGORY_BY_NORMALIZED_LABEL` keep matching on the *shipped* labels on
 * purpose: those maps exist to resolve a field to its category, and renaming a heading on the
 * Spec Categories screen must not unhook the fields that resolve through it.
 */
export function getSpecCategories(): SpecCategoryDef[] {
  const overlay = readOverlay();
  const shipped = withoutRemoved(
    'categories',
    SPEC_CATEGORIES.map((c) => applyOverride(c, overlay.categories[c.id])),
    (c) => c.id,
  );
  const added: SpecCategoryDef[] = overlay.added.categories.map((c) => ({
    id: c.id,
    label: c.label ?? c.id,
    section: c.section,
    order: c.order ?? 800,
    layout: c.layout,
    productLines: c.productLines,
  }));
  return [...shipped, ...added].sort((a, b) => a.order - b.order);
}

export function getCategoryById(id: string): SpecCategoryDef | undefined {
  return getSpecCategories().find((c) => c.id === id);
}

/**
 * Resolve a field to its category id: explicit `categoryId` wins, then the legacy
 * `category` label, then the field's own label, and finally the catch-all.
 */
export function resolveCategoryId(field: SpecFieldDefinition): string {
  if (field.categoryId && CATEGORY_BY_ID.has(field.categoryId)) return field.categoryId;
  const label = field.category ?? field.label;
  return CATEGORY_BY_NORMALIZED_LABEL.get(normalizeLabel(label)) ?? OTHER_CATEGORY_ID;
}

export function isCategoryForLine(category: SpecCategoryDef, line: ProductLineCode): boolean {
  return !category.productLines || category.productLines.includes(line);
}

/**
 * A category with one visible field for this product line reads better merged (the label
 * spans two columns and the duplicate item label is dropped). An explicit `layout` wins.
 */
export function resolveCategoryLayout(
  category: SpecCategoryDef,
  fieldCount: number,
): SpecCategoryLayout {
  if (category.layout) return category.layout;
  return fieldCount <= 1 ? 'merge' : 'pair';
}

/**
 * Column-label aliases accepted by the Excel template parser, keyed by field key.
 *
 * Kept beside the schema rather than on each field so the 100+ entries of
 * `specFieldConfig.ts` stay pure field data. A field may also declare `sourceLabels`
 * directly; both are consulted, and the field's own `label` always matches implicitly.
 * Aliases cover the spellings ASUS HQ datasheets actually use, including their typos
 * (`Reslolution`, `Reparability`), so a hand-edited template still loads.
 */
export const FIELD_SOURCE_LABELS: Record<string, string[]> = {
  partNo: ['Part No', 'Part Number'],
  modelName: ['Sales Model Name'],
  eanCode: ['EAN'],
  upcCode: ['UPC'],
  baseUnit: ['BASE UNIT', 'BOPN'],
  specOs: ['OS'],
  lcdCoverMaterial: ['LCD cover-material'],
  lcdCoverColor: ['LCD cover-color'],
  topCaseMaterial: ['Top case-material'],
  topCaseColor: ['Top case-color'],
  bottomCaseMaterial: ['Bottom case-material'],
  bottomCaseColor: ['Bottom case-color'],
  specProcessor: ['CPU'],
  specOnBoardProcessor: ['On board processor'],
  integratedGpu: ['GPU'],
  neuralProcessor: ['NPU'],
  resolution: ['Reslolution'],
  ipsLevel: ['IPS-level'],
  colorGamut: ['Color gamut'],
  viewAngle: ['View angle(H/V)', 'View angle'],
  responseTime: ['Response time(Typ/Max) (ms)', 'Response time'],
  contrast: ['Contrast (Typ)'],
  refreshRate: ['Refresh rate'],
  onBoardMemory: ['On board memory'],
  howToUpgradeMemory: ['How to upgrade memory'],
  memoryMax: ['Memory Max'],
  militaryGrade: ['Military grade'],
  weightWithBattery: ['Starting Weight (with Battery)', 'Weight (with Battery)'],
  weightWithoutBattery: ['Starting Weight (w/o Battery)', 'Weight (w/o Battery)'],
  weightAio2: ['Weight_2'],
  ioPorts: ['I/O ports', 'IO Port'],
  requiredChargingPower: ['Required charging power'],
  standType: ['Stand type'],
  voiceControl: ['Voice control'],
  keyboardType: ['Keyboard type'],
  keyboardLanguage: ['Keyboard language'],
  numberPad: ['NumberPad'],
  frontFacingCamera: ['Front-facing camera'],
  wlan: ['Wireless', 'WLAN'],
  onboardWireless: ['On board Wireless'],
  expansionSlot: ['Expansion Slot(includes used)', 'Expansion Slot'],
  myAsusFeatures: ['My ASUS features', 'MyASUS feature'],
  includedInBox: ['Included in the Box', 'Include in the box'],
  batteryWarranty: ['Battery warranty'],
  recommendedWarrantyPackage: [
    'Recommended Warranty Package Part number//Extension and/or upgrade service from base warranty',
  ],
  repairabilityFrance: ['Reparability Index (for France)', 'Repairability Index (for France)'],
  repairabilityBelgium: ['Reparability Index (for Belgium)', 'Repairability Index (for Belgium)'],
};

/**
 * Source labels that legitimately feed more than one field.
 *
 * ASUS sends one `Operating System` row, but the app keeps the value twice on purpose: the
 * Product Info copy drives `Product.operatingSystem` (the summary row, the grid filter, the
 * info card) and the spec copy renders inside the specification table. A single-winner lookup
 * has to drop one of them — which is why `info.operatingSystem` could never import.
 *
 * Only for genuine duplicates. Two fields that merely share a label and mean different things
 * (`battery` vs `sustainableBattery`) are separated with `templateLabel` instead.
 */
export const LABEL_FANOUT: Record<string, string[]> = {
  'Operating System': ['info.operatingSystem', 'specOs'],
  Processor: ['info.processor', 'specProcessor'],
  'On board Processor': ['info.onBoardProcessor', 'specOnBoardProcessor'],
};

/**
 * Shortest label the prefix rule will consider.
 *
 * ASUS truncates long column headings at a width that varies per sheet: the Recommended
 * Warranty row ends `...from base warranty` on the Notebook datasheet and `...from base warran`
 * on Desktop and All-in-One. A literal alias fixes this month and breaks next month, so the
 * resolver falls back to a prefix match — but only for labels long enough that a prefix cannot
 * be a coincidence. 24 normalized characters keeps `Weight` away from `Weight_2`.
 */
export const MIN_PREFIX_MATCH_CHARS = 24;
