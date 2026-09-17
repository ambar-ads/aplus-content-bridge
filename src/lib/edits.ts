/**
 * Perubahan yang diketik langsung di halaman, disimpan terpisah dari data sumbernya.
 *
 * Bentuknya overlay berkunci 90PN, bukan salinan seluruh baris. Alasannya: sumber datanya bisa
 * berganti di tengah jalan — pengguna mengunggah file master lain, atau menambahkan price list —
 * dan perubahan yang sudah diketik tidak boleh ikut hilang atau justru menimpa data yang lebih baru
 * di kolom yang tidak disentuh.
 *
 * Overlay ini membersihkan dirinya sendiri: begitu nilai yang diketik sama dengan nilai
 * sumbernya, entrinya dibuang. Tanpa itu, perubahan bulan lalu akan diam-diam menimpa data bulan
 * ini di kolom yang sama.
 */
import { FORWEB_COLUMNS } from './config';
import type { ForWebKey, ForWebRow } from './types';

export type ProductEdit = Partial<Record<ForWebKey, string>>;
export type EditMap = Record<string, ProductEdit>;

const STORAGE_KEY = 'aplus-bridge-edits-v1';

/** Kolom yang boleh diubah. 90PN adalah kunci join, jadi tidak ikut. */
export const EDITABLE_COLUMNS = FORWEB_COLUMNS.filter(
  (c) => c.key !== 'partNumber' && c.key !== 'partNumberAuto',
);

export function loadEdits(): EditMap {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return {};
    const parsed: unknown = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? (parsed as EditMap) : {};
  } catch {
    // Mode penyamaran atau penyimpanan diblokir — editor tetap jalan, hanya tidak bertahan.
    return {};
  }
}

export function saveEdits(edits: EditMap): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(edits));
  } catch {
    /* tidak apa-apa: perubahan tetap berlaku untuk sesi ini */
  }
}

export function countEdits(edits: EditMap): { products: number; fields: number } {
  const products = Object.keys(edits).length;
  const fields = Object.values(edits).reduce((n, e) => n + Object.keys(e).length, 0);
  return { products, fields };
}

export interface ApplyResult {
  rows: ForWebRow[];
  /** Overlay setelah dibersihkan dari entri yang sudah sama dengan sumbernya. */
  cleaned: EditMap;
  /** 90PN -> daftar kolom yang benar-benar berbeda dari sumbernya. */
  changedByPn90: Record<string, ForWebKey[]>;
  /** Perubahan untuk 90PN yang tidak ada di data sekarang. Disimpan, tapi tidak berlaku. */
  orphanCount: number;
}

const norm = (v: unknown) => String(v ?? '').replace(/\r\n/g, '\n').trim();

export function applyEdits(rows: ForWebRow[], edits: EditMap): ApplyResult {
  const cleaned: EditMap = {};
  const changedByPn90: Record<string, ForWebKey[]> = {};
  const present = new Set<string>();

  const next = rows.map((row) => {
    const pn90 = (row.partNumber || '').trim();
    present.add(pn90);
    const edit = edits[pn90];
    if (!edit) return row;

    const applied: ForWebRow = { ...row };
    const changed: ForWebKey[] = [];
    const keep: ProductEdit = {};

    for (const [key, value] of Object.entries(edit) as [ForWebKey, string][]) {
      if (norm(value) === norm(row[key])) continue; // sudah sama dengan sumbernya
      applied[key] = value;
      changed.push(key);
      keep[key] = value;
    }

    if (changed.length) {
      cleaned[pn90] = keep;
      changedByPn90[pn90] = changed;
    }
    return changed.length ? applied : row;
  });

  let orphanCount = 0;
  for (const [pn90, edit] of Object.entries(edits)) {
    if (present.has(pn90)) continue;
    // Produknya sedang tidak ada di data — mungkin file master lain. Simpan, jangan buang.
    cleaned[pn90] = edit;
    orphanCount += 1;
  }

  return { rows: next, cleaned, changedByPn90, orphanCount };
}

export function setField(
  edits: EditMap,
  pn90: string,
  key: ForWebKey,
  value: string,
): EditMap {
  return { ...edits, [pn90]: { ...(edits[pn90] ?? {}), [key]: value } };
}

export interface FieldChange {
  pn90: string;
  key: ForWebKey;
  value: string;
}

/**
 * Apply many changes at once.
 *
 * Pasting a column from Excel can touch several hundred cells. Folding them into one new overlay
 * keeps that a single render rather than one per cell, and keeps the whole paste undoable as one
 * step rather than leaving it half-applied if something throws part-way.
 */
export function setFields(edits: EditMap, changes: FieldChange[]): EditMap {
  if (!changes.length) return edits;
  const next: EditMap = { ...edits };
  for (const { pn90, key, value } of changes) {
    next[pn90] = { ...(next[pn90] ?? {}), [key]: value };
  }
  return next;
}

export function revertProduct(edits: EditMap, pn90: string): EditMap {
  const next = { ...edits };
  delete next[pn90];
  return next;
}
