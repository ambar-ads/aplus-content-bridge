/**
 * Katalog akumulatif, berkunci 90PN.
 *
 * Dua alasan kenapa bentuknya begini, keduanya berasal dari cara sistem lama bekerja:
 *
 * 1. Akumulasi. `RAW.KSP.LInk` tidak pernah dihapus isinya - PM menambah baris di bawah sejak
 *    Agustus 2025. Price list bulanan hanya memuat yang sedang dijual (155 dari 392 produk yang
 *    tayang). Kalau CSV dibangun ulang dari price list saja, 237 produk hilang dari situs.
 *    Karena itu katalog inilah sumbernya, dan price list hanya menambahi.
 *
 * 2. Join berbasis kunci. Sheet `ForWeb` menempelkan dua array hasil `SORT` secara berdampingan
 *    dan hanya diselaraskan oleh urutan sort. Kalau satu sisi punya baris yang tidak ada di sisi
 *    lain, seluruh spec di bawahnya menempel ke produk yang salah. Di sini semuanya dijoin lewat
 *    90PN, sehingga mode kegagalan itu tidak bisa terjadi.
 */
import { CATALOG_SCHEMA_VERSION, VISIBILITY_FILTER, EMPTY_FORWEB_ROW } from './config';
import { isBlankValue } from './legacyCsv';
import type { Catalog, CatalogEntry, ForWebRow, Provenance } from './types';

/** Panjang 90PN yang normal, mis. `90NX04U1-M004P0`. Yang lebih pendek berarti terpotong. */
const PN90_NORMAL_LENGTH = 15;

export function pn90Of(row: ForWebRow): string {
  // "0" adalah penanda kosong dari workbook lama, bukan nilai - jangan sampai jadi kunci.
  for (const candidate of [row.partNumber, row.partNumberAuto]) {
    const trimmed = (candidate ?? '').trim();
    if (trimmed && !isBlankValue(trimmed)) return trimmed;
  }
  return '';
}

/** Urutan yang sama dengan `SORT(..., 1, TRUE)` di workbook lama: menaik menurut Part Number. */
function byPartNumber(a: CatalogEntry, b: CatalogEntry): number {
  if (a.pn90 === b.pn90) return 0;
  return a.pn90 < b.pn90 ? -1 : 1;
}

export function buildCatalogFromRows(
  rows: ForWebRow[],
  base: { origin: Provenance['origin']; sourceFile?: string },
): { catalog: Catalog; warnings: string[] } {
  const now = new Date().toISOString();
  const warnings: string[] = [];
  const byPn = new Map<string, CatalogEntry>();
  const nameSeen = new Map<string, string[]>();

  rows.forEach((row, i) => {
    const pn90 = pn90Of(row);
    const label = row.productName || `row ${i + 2}`;

    if (!pn90) {
      warnings.push(`${label}: has no Part Number — skipped`);
      return;
    }
    if (pn90.length < PN90_NORMAL_LENGTH) {
      warnings.push(
        `${label}: Part Number "${pn90}" looks truncated (${pn90.length} characters; a full one has ${PN90_NORMAL_LENGTH})`,
      );
    }
    if (byPn.has(pn90)) {
      const prev = byPn.get(pn90);
      warnings.push(
        `Duplicate Part Number "${pn90}" (${prev ? prev.row.productName : '?'} vs ${row.productName}) — the first one is used`,
      );
      return;
    }

    const names = nameSeen.get(row.productName) ?? [];
    names.push(pn90);
    nameSeen.set(row.productName, names);

    byPn.set(pn90, {
      pn90,
      row,
      provenance: { ...base, firstSeenAt: now, lastUpdatedAt: now },
    });
  });

  for (const [name, pns] of nameSeen) {
    if (pns.length > 1) {
      warnings.push(
        `Duplicate Product Name "${name}" (${pns.join(', ')}) — the site matches with .find(), so only the first is reachable by name`,
      );
    }
  }

  return {
    catalog: {
      schemaVersion: CATALOG_SCHEMA_VERSION,
      updatedAt: now,
      entries: [...byPn.values()].sort(byPartNumber),
    },
    warnings,
  };
}

export interface MergeReport {
  added: CatalogEntry[];
  refreshed: CatalogEntry[];
  keptUntouched: number;
  /** Ada di katalog tapi tidak lagi di price list. Dipertahankan - inilah inti akumulasi. */
  notInSource: number;
  warnings: string[];
}

export interface IncomingModel {
  pn90: string;
  row: ForWebRow;
  provenance: Omit<Provenance, 'firstSeenAt' | 'lastUpdatedAt'>;
}

/**
 * Gabungkan hasil import price list ke katalog.
 *
 * Default-nya hanya menambah. Baris yang sudah ada dibiarkan apa adanya kecuali 90PN-nya
 * disebut di `refreshPn90`. Ini disengaja: 392 baris yang sudah terbit teksnya sudah benar, dan
 * aturan derivasi `Auto-*` yang keliru tidak boleh bisa merusaknya secara massal.
 */
export function mergePriceList(
  catalog: Catalog,
  incoming: IncomingModel[],
  options: { refreshPn90?: Set<string> } = {},
): { catalog: Catalog; report: MergeReport } {
  const now = new Date().toISOString();
  const refresh = options.refreshPn90 ?? new Set<string>();
  const byPn = new Map(catalog.entries.map((e) => [e.pn90, e]));
  const report: MergeReport = {
    added: [],
    refreshed: [],
    keptUntouched: 0,
    notInSource: 0,
    warnings: [],
  };
  const incomingPns = new Set<string>();

  for (const item of incoming) {
    if (!item.pn90) {
      report.warnings.push('A model column has no Part No — skipped');
      continue;
    }
    incomingPns.add(item.pn90);
    const existing = byPn.get(item.pn90);

    if (!existing) {
      const entry: CatalogEntry = {
        pn90: item.pn90,
        row: item.row,
        provenance: { ...item.provenance, firstSeenAt: now, lastUpdatedAt: now },
      };
      byPn.set(item.pn90, entry);
      report.added.push(entry);
      continue;
    }

    if (refresh.has(item.pn90)) {
      // Tagline / KSP / Link tidak pernah ada di price list - pertahankan yang sudah ada.
      const merged: ForWebRow = {
        ...item.row,
        tagline: existing.row.tagline,
        ksp: existing.row.ksp,
        link: existing.row.link,
      };
      const entry: CatalogEntry = {
        pn90: item.pn90,
        row: merged,
        provenance: { ...existing.provenance, ...item.provenance, lastUpdatedAt: now },
      };
      byPn.set(item.pn90, entry);
      report.refreshed.push(entry);
    } else {
      report.keptUntouched += 1;
    }
  }

  report.notInSource = catalog.entries.filter((e) => !incomingPns.has(e.pn90)).length;

  return {
    catalog: {
      schemaVersion: CATALOG_SCHEMA_VERSION,
      updatedAt: now,
      entries: [...byPn.values()].sort(byPartNumber),
    },
    report,
  };
}

export interface CopyRecord {
  tagline: string;
  ksp: string;
  link: string;
  modelName?: string;
}

export interface CopyReport {
  updated: number;
  unchanged: number;
  /** Copy yang 90PN-nya tidak ada di katalog - biasanya produk yang spec-nya belum di-import. */
  orphanCopy: { pn90: string; modelName?: string }[];
  /** Produk di katalog yang sama sekali belum punya Tagline/KSP/Link. */
  missingCopy: { pn90: string; productName: string }[];
}

/** Tempelkan Tagline / KSP / Link dari workbook lama ke katalog, dijoin lewat 90PN. */
export function applyCopy(
  catalog: Catalog,
  copyByPn90: Map<string, CopyRecord>,
): { catalog: Catalog; report: CopyReport } {
  const now = new Date().toISOString();
  const report: CopyReport = { updated: 0, unchanged: 0, orphanCopy: [], missingCopy: [] };
  const known = new Set(catalog.entries.map((e) => e.pn90));

  for (const [pn90, rec] of copyByPn90) {
    if (!known.has(pn90)) report.orphanCopy.push({ pn90, modelName: rec.modelName });
  }

  const entries = catalog.entries.map((entry) => {
    const rec = copyByPn90.get(entry.pn90);
    const hasCopy =
      !isBlankValue(entry.row.tagline) ||
      !isBlankValue(entry.row.ksp) ||
      !isBlankValue(entry.row.link);

    if (!rec) {
      if (!hasCopy) {
        report.missingCopy.push({ pn90: entry.pn90, productName: entry.row.productName });
      }
      return entry;
    }

    // Copy dari workbook hanya menimpa kalau memang berisi. Nilai lama yang sudah terbit tidak
    // boleh terhapus hanya karena sel di workbook kebetulan kosong bulan ini.
    const next: ForWebRow = {
      ...entry.row,
      tagline: isBlankValue(rec.tagline) ? entry.row.tagline : rec.tagline,
      ksp: isBlankValue(rec.ksp) ? entry.row.ksp : rec.ksp,
      link: isBlankValue(rec.link) ? entry.row.link : rec.link,
    };
    const changed =
      next.tagline !== entry.row.tagline ||
      next.ksp !== entry.row.ksp ||
      next.link !== entry.row.link;

    if (changed) report.updated += 1;
    else report.unchanged += 1;

    return changed
      ? { ...entry, row: next, provenance: { ...entry.provenance, lastUpdatedAt: now } }
      : entry;
  });

  return {
    catalog: { schemaVersion: CATALOG_SCHEMA_VERSION, updatedAt: now, entries },
    report,
  };
}

/** Katalog -> baris CSV, dengan urutan dan filter yang sama seperti sistem lama. */
export function catalogToRows(catalog: Catalog): ForWebRow[] {
  const blocked = new Set(VISIBILITY_FILTER.blockedPn90);
  return catalog.entries
    .filter((e) => (VISIBILITY_FILTER.includeAll ? true : !blocked.has(e.pn90)))
    .filter((e) => !blocked.has(e.pn90))
    .slice()
    .sort(byPartNumber)
    .map((e) => ({ ...EMPTY_FORWEB_ROW, ...e.row }) as ForWebRow);
}
