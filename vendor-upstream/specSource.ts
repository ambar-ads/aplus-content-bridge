import type { ProductLineCode } from '@/types/specForm';
import {
  FIELD_SOURCE_LABELS,
  LABEL_FANOUT,
  MIN_PREFIX_MATCH_CHARS,
  normalizeLabel,
} from '@/config/specSchema';
import { getFieldsForProductLine, getProductInfoFieldsForLine } from '@/config/specFieldConfig';

/**
 * Everything shared by the two things that read ASUS spec data: the app's own Excel template
 * (`specTemplate.ts`) and the ASUS PM datasheet (`pmDatasheetImport.ts`).
 *
 * It lives apart from both so the PM importer can reuse the label resolver without importing
 * the template generator, and so the resolver has one implementation rather than the copy that
 * the paste path used to keep by round-tripping a whole workbook through ExcelJS.
 */

// ─── Parsed shapes ──────────────────────────────────────────────────

export interface ParsedProduct {
  /** Column heading, e.g. "Product 1" or the Sales Model Name from a PM datasheet. */
  columnLabel: string;
  modelName: string;
  /** Catalog spec fields, keyed by field key. */
  specs: Record<string, string>;
  /** Product Info fields, keyed *without* the `info.` prefix. */
  productInfo: Record<string, string>;
  /** Rows whose label matched no field — kept, never dropped. */
  unmapped: { label: string; value: string }[];
}

export interface ParsedFootnote {
  id: string;
  text: string;
  /** Field keys this disclaimer applies to. */
  appliesTo: string[];
}

export interface FootnoteSuggestion {
  /** Field key the disclaimer was found inside. */
  fieldKey: string;
  fieldLabel: string;
  text: string;
  /** The value with the disclaimer line removed. */
  cleanedValue: string;
}

/** Stable identity for a suggestion, used to track which ones the reviewer accepted. */
export function suggestionKey(s: FootnoteSuggestion): string {
  return `${s.fieldKey}|${s.text}`;
}

export interface ParsedTemplate {
  productLine: ProductLineCode | '';
  schemaVersion: number | null;
  products: ParsedProduct[];
  footnotes: ParsedFootnote[];
  /** Inline `*` disclaimers detected in values, awaiting confirmation. */
  suggestions: FootnoteSuggestion[];
  errors: string[];
  warnings: string[];
}

// ─── Cell text ──────────────────────────────────────────────────────

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

// ─── Inline disclaimers ─────────────────────────────────────────────

/**
 * Decide whether a `*`-prefixed line is a disclaimer or just an asterisk in the data.
 *
 * The source data is full of asterisks that are NOT footnotes — `2*2` for Wi-Fi spatial
 * streams, `SO-DIMM *2` for quantity, `4Y Accidental Damage Protection*` as an inline marker.
 * Requiring a line to *start* with the asterisk and then carry a real sentence keeps those
 * out. Anything that passes is only ever *suggested*, never applied silently.
 */
const MIN_FOOTNOTE_CHARS = 25;
const MIN_FOOTNOTE_WORDS = 3;

function looksLikeDisclaimer(line: string): boolean {
  const trimmed = line.trim();
  if (!trimmed.startsWith('*')) return false;
  const body = trimmed.replace(/^\*+\s*/, '');
  if (body.length < MIN_FOOTNOTE_CHARS) return false;
  return body.split(/\s+/).filter(Boolean).length >= MIN_FOOTNOTE_WORDS;
}

/** Split a value into its real content and any trailing disclaimer lines. */
export function extractInlineDisclaimers(value: string): {
  cleaned: string;
  disclaimers: string[];
} {
  const lines = value.split('\n');
  const kept: string[] = [];
  const disclaimers: string[] = [];
  lines.forEach((line) => {
    if (looksLikeDisclaimer(line)) disclaimers.push(line.trim().replace(/^\*+\s*/, ''));
    else kept.push(line);
  });
  return { cleaned: kept.join('\n').trim(), disclaimers };
}

// ─── Label resolution ───────────────────────────────────────────────

export interface FieldLookup {
  /**
   * Normalized label -> the field keys it fills. Usually one; a declared fan-out fills several.
   * Product Info keys carry the `info.` prefix.
   */
  byLabel: Map<string, string[]>;
  labelOf: Map<string, string>;
}

/**
 * Build the label resolver for one product line.
 *
 * Registration happens in three tiers with explicit, documented precedence. It used to be a
 * single `Map.set` per label, so the *last* field registered silently won: `Battery` resolved
 * to `sustainableBattery` rather than `battery`, and `Operating System` to `specOs`, which is
 * why `info.operatingSystem` could never import at all.
 *
 *   1. `templateLabel` — a hard claim. Deliberately set to separate two fields that share a
 *      display label, so it overwrites.
 *   2. `label` — a soft claim, first registration wins. Product Info registers first, so the
 *      order is deterministic rather than accidental.
 *   3. `sourceLabels` / `FIELD_SOURCE_LABELS` — soft claims, first wins.
 *
 * `LABEL_FANOUT` is applied last and is the only way one label reaches several fields.
 */
export function buildFieldLookup(line: ProductLineCode): FieldLookup {
  const byLabel = new Map<string, string[]>();
  const labelOf = new Map<string, string>();
  const hardClaimed = new Set<string>();
  const known = new Set<string>();

  const claim = (norm: string, key: string, hard: boolean) => {
    if (!norm) return;
    if (hard) {
      if (import.meta.env.DEV && hardClaimed.has(norm) && byLabel.get(norm)?.[0] !== key) {
        console.warn(
          `[specSource] templateLabel conflict on "${norm}": ${byLabel.get(norm)} vs ${key}`,
        );
      }
      byLabel.set(norm, [key]);
      hardClaimed.add(norm);
      return;
    }
    if (!byLabel.has(norm)) byLabel.set(norm, [key]);
  };

  const register = (
    key: string,
    label: string,
    templateLabel: string | undefined,
    aliases: string[] | undefined,
  ) => {
    known.add(key);
    labelOf.set(key, templateLabel ?? label);
    if (templateLabel) claim(normalizeLabel(templateLabel), key, true);
    claim(normalizeLabel(label), key, false);
    const bare = key.replace(/^info\./, '');
    [...(aliases ?? []), ...(FIELD_SOURCE_LABELS[bare] ?? [])].forEach((alias) =>
      claim(normalizeLabel(alias), key, false),
    );
  };

  getProductInfoFieldsForLine(line).forEach((f) =>
    register(`info.${f.key}`, f.label, f.templateLabel, f.sourceLabels),
  );
  getFieldsForProductLine(line).forEach((f) =>
    register(f.key, f.label, f.templateLabel, f.sourceLabels),
  );

  // Fan-out last, and only for keys this product line actually has.
  Object.entries(LABEL_FANOUT).forEach(([label, keys]) => {
    const present = keys.filter((k) => known.has(k));
    if (present.length > 0) byLabel.set(normalizeLabel(label), present);
  });

  return { byLabel, labelOf };
}

/**
 * Resolve an incoming column label to the field keys it fills.
 *
 * Exact match first, then a *unique* prefix match. The prefix rule exists for one reason: ASUS
 * truncates long headings at a width that differs per sheet, so the Recommended Warranty row
 * arrives with two spellings in the same workbook. Requiring exactly one candidate, and a
 * minimum length, is what keeps it from matching things it should not — deliberately not a
 * fuzzy/edit-distance match, which would be unpredictable and would still need this guard.
 *
 * Returns an empty array for an unknown label; callers keep those as custom specs.
 */
export function resolveLabel(lookup: FieldLookup, rawLabel: string): string[] {
  const norm = normalizeLabel(rawLabel);
  if (!norm) return [];

  const exact = lookup.byLabel.get(norm);
  if (exact) return exact;

  if (norm.length < MIN_PREFIX_MATCH_CHARS) return [];

  // Several registered labels routinely match the same field — its own `label` and a long
  // `sourceLabels` alias both prefix-match the truncated heading. That is not an ambiguity, so
  // compare the resolved *keys*, not the number of candidate labels.
  let hit: string[] | undefined;
  for (const [candidate, keys] of lookup.byLabel) {
    if (candidate.length < MIN_PREFIX_MATCH_CHARS) continue;
    if (!candidate.startsWith(norm) && !norm.startsWith(candidate)) continue;
    if (hit && hit.join('|') !== keys.join('|')) return []; // genuinely ambiguous — keep as custom
    hit = keys;
  }
  return hit ?? [];
}

/**
 * Fail loudly in development when two fields end up fighting over one label.
 *
 * This is the guard that stops the class of bug the tiered lookup just fixed from coming back
 * the next time somebody adds a field: the collision is invisible at the definition site and
 * only shows up as an import quietly filling the wrong row.
 */
export function assertNoAmbiguousLabels(line: ProductLineCode): string[] {
  const seen = new Map<string, string[]>();
  const note = (key: string, label: string | undefined) => {
    if (!label) return;
    const norm = normalizeLabel(label);
    seen.set(norm, [...(seen.get(norm) ?? []), key]);
  };
  getProductInfoFieldsForLine(line).forEach((f) => note(`info.${f.key}`, f.templateLabel ?? f.label));
  getFieldsForProductLine(line).forEach((f) => note(f.key, f.templateLabel ?? f.label));

  const fanout = new Set(Object.keys(LABEL_FANOUT).map(normalizeLabel));
  const problems: string[] = [];
  seen.forEach((keys, norm) => {
    if (keys.length > 1 && !fanout.has(norm)) {
      problems.push(`${line}: "${norm}" is claimed by ${keys.join(', ')}`);
    }
  });
  return problems;
}
