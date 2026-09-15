/**
 * Locating the input workbooks, in one place.
 *
 * ASUS ships price lists under at least two naming schemes ("ABP Price List …", "ACID Monthly
 * Pricelist …"), so the scripts match on the words rather than a fixed prefix and take whichever
 * file is newest.
 */
import fs from 'node:fs';
import path from 'node:path';

const PRICE_LIST = /price\s*list/i;
const MASTER_SOURCE = /^A\+ Content Source.*\.xlsx$/i;

function newestMatching(dir, test) {
  if (!fs.existsSync(dir)) return undefined;
  return fs
    .readdirSync(dir)
    .filter((f) => /\.xlsx$/i.test(f) && !f.startsWith('~$') && test(f))
    .map((f) => ({ f, m: fs.statSync(path.join(dir, f)).mtimeMs }))
    .sort((a, b) => b.m - a.m)[0]?.f;
}

export const newestPriceList = (dir) => newestMatching(dir, (f) => PRICE_LIST.test(f));
export const legacySourceWorkbook = (dir) => newestMatching(dir, (f) => MASTER_SOURCE.test(f));
