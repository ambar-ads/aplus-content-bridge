import type { ProductLineCode } from '@/vendor/types';

/** Satu baris pada CSV ForWeb, dipakai sebagai bentuk kanonik di seluruh app. */
export interface ForWebRow {
  partNumber: string;
  productName: string;
  marketingName: string;
  tagline: string;
  ksp: string;
  link: string;
  /** Kolom "Part Number" kedua. Di workbook lama berasal dari sisi Auto-*, bukan RAW. */
  partNumberAuto: string;
  title: string;
  cpu: string;
  gpu: string;
  os: string;
  office: string;
  panelSize: string;
  resolution: string;
  brightness: string;
  ram: string;
  ssd: string;
  ioPort: string;
  weight: string;
  size: string;
  expansionSlot: string;
  includeInTheBox: string;
  battery: string;
  powerSupply: string;
  security: string;
}

export type ForWebKey = keyof ForWebRow;

/** Dari mana sebuah entri katalog berasal — tidak pernah ada di workbook lama. */
export interface Provenance {
  /** 'seed' = ikut dari CSV yang sudah tayang; 'priceList' = hasil import. */
  origin: 'seed' | 'priceList' | 'manual';
  /** Nama file price list saat entri ini pertama masuk / terakhir di-refresh. */
  sourceFile?: string;
  sheetName?: string;
  productLine?: ProductLineCode;
  group?: 'main' | 'tkdn';
  firstSeenAt: string;
  lastUpdatedAt: string;
}

export interface CatalogEntry {
  /** Kunci katalog: 90PN. Inilah pengganti join berbasis posisi di workbook lama. */
  pn90: string;
  row: ForWebRow;
  provenance: Provenance;
}

export interface Catalog {
  schemaVersion: number;
  updatedAt: string;
  entries: CatalogEntry[];
}
