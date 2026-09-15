import type { ForWebKey } from './types';

export const CATALOG_SCHEMA_VERSION = 1;

/** Nama file yang diharapkan situs lama. Harus persis — halaman mem-fetch-nya relatif. */
export const FORWEB_FILENAME = 'A+ Content Source(ForWeb).csv';

/**
 * Produk mana yang boleh masuk CSV publik.
 *
 * Keputusannya: samakan dengan perilaku sistem lama, yaitu semua produk di katalog ikut terbit.
 * Perlu dicatat bahwa workbook sumbernya bertanda "INTERNAL USE ONLY. DO NOT SHARE THIS FILE TO
 * PARTNERS" sementara hasilnya tayang di GitHub Pages tanpa autentikasi. Filternya sengaja
 * dikumpulkan di satu tempat ini supaya bisa diperketat tanpa menyentuh logika lain.
 */
export const VISIBILITY_FILTER: { includeAll: boolean; blockedPn90: string[] } = {
  includeAll: true,
  blockedPn90: [],
};

export interface ForWebColumn {
  key: ForWebKey;
  header: string;
  /**
   * Apa yang ditulis saat nilainya kosong.
   *
   * Di workbook lama kolom 2-6 berasal dari spill `SORT(FILTER(...))` yang memunculkan `0`
   * untuk sel kosong, sedangkan kolom Auto-* meninggalkannya benar-benar kosong. Halaman
   * memetakan keduanya ke "No Available", jadi ini soal kesetiaan pada format, bukan tampilan.
   */
  emptyAs: '0' | '';
  /** Field yang dipecah per baris oleh halaman lama. */
  multiline?: boolean;
}

/**
 * 25 kolom CSV ForWeb, berurutan.
 *
 * `Part Number` memang muncul dua kali — itu bentuk asli filenya, dan halaman lama
 * mengandalkannya (`delete row['Part Number']` menyasar kemunculan kedua). Jangan dirapikan.
 */
export const FORWEB_COLUMNS: ForWebColumn[] = [
  { key: 'partNumber', header: 'Part Number', emptyAs: '' },
  { key: 'productName', header: 'Product Name', emptyAs: '0' },
  { key: 'marketingName', header: 'Marketing name', emptyAs: '0' },
  { key: 'tagline', header: 'Tagline', emptyAs: '0' },
  { key: 'ksp', header: 'KSP', emptyAs: '0', multiline: true },
  { key: 'link', header: 'Link', emptyAs: '0' },
  { key: 'partNumberAuto', header: 'Part Number', emptyAs: '' },
  { key: 'title', header: 'Title', emptyAs: '' },
  { key: 'cpu', header: 'CPU', emptyAs: '' },
  { key: 'gpu', header: 'GPU', emptyAs: '' },
  { key: 'os', header: 'OS', emptyAs: '' },
  { key: 'office', header: 'Office', emptyAs: '' },
  { key: 'panelSize', header: 'Panel Size', emptyAs: '' },
  { key: 'resolution', header: 'Resolution', emptyAs: '' },
  { key: 'brightness', header: 'Brightness', emptyAs: '' },
  { key: 'ram', header: 'RAM', emptyAs: '' },
  { key: 'ssd', header: 'SSD', emptyAs: '' },
  { key: 'ioPort', header: 'IO Port', emptyAs: '', multiline: true },
  { key: 'weight', header: 'Weight (with Battery)', emptyAs: '' },
  { key: 'size', header: 'Size', emptyAs: '' },
  { key: 'expansionSlot', header: 'Expansion Slot', emptyAs: '', multiline: true },
  { key: 'includeInTheBox', header: 'Include in the box', emptyAs: '', multiline: true },
  { key: 'battery', header: 'Battery', emptyAs: '' },
  { key: 'powerSupply', header: 'Power Supply', emptyAs: '' },
  { key: 'security', header: 'Security', emptyAs: '', multiline: true },
];

export const EMPTY_FORWEB_ROW: Record<ForWebKey, string> = Object.fromEntries(
  FORWEB_COLUMNS.map((c) => [c.key, '']),
) as Record<ForWebKey, string>;
