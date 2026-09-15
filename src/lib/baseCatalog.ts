/**
 * Katalog dasar yang ikut di-bundle saat build.
 *
 * Dihasilkan `npm run seed:catalog` dari CSV yang sedang tayang, lalu ikut di-commit. Karena
 * app ini statis dan tanpa server, file inilah "database"-nya — bentuk yang sama dengan cara
 * tim sekarang bekerja, yaitu file di dalam repo.
 */
import raw from '../../data/catalog.json';
import type { Catalog } from './types';

export const BASE_CATALOG = raw as unknown as Catalog;
