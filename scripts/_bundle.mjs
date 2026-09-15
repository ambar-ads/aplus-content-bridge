/**
 * Bundling modul TypeScript dari src/ supaya bisa dijalankan di Node.
 *
 * Sengaja mem-bundle kode `src/` yang sebenarnya — bukan menyalin ulang logikanya ke dalam
 * script — supaya yang diuji memang kode yang dipakai app. Pola ini mengikuti scripts/ di repo
 * APLUS utama.
 */
import { build } from 'esbuild';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

export const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

/**
 * Hasil bundle ditaruh di dalam node_modules/.cache, bukan temp dir sistem, supaya paket yang
 * di-`external` (exceljs, papaparse) tetap bisa di-resolve Node dari lokasi file bundle-nya.
 */
const CACHE_DIR = path.join(ROOT, 'node_modules/.cache/aplus-legacy');

export async function loadModule(entryRelative) {
  fs.mkdirSync(CACHE_DIR, { recursive: true });
  const outfile = path.join(
    CACHE_DIR,
    entryRelative.replace(/[\\/]/g, '_').replace(/\.ts$/, '') + '.mjs',
  );
  await build({
    entryPoints: [path.join(ROOT, entryRelative)],
    outfile,
    bundle: true,
    format: 'esm',
    platform: 'node',
    target: 'node20',
    logLevel: 'silent',
    external: ['exceljs', 'papaparse'],
    alias: { '@': path.join(ROOT, 'src') },
  });
  return import(pathToFileURL(outfile).href);
}
