# vendor/

Kode yang berasal dari repo APLUS utama (`../asus-aplus-2-0`). Repo itu **hanya dibaca**.

| File | Asal | Cara pemeliharaan |
|---|---|---|
| `pmLayout.ts` | `src/config/pmDatasheetLayout.ts` | Disalin otomatis oleh `npm run sync:vendor`. **Jangan diedit.** |
| `pmCell.ts` | `specSource.ts`, `specSchema.ts`, `pmDatasheetImport.ts` | Ditulis ulang manual — lihat komentar di dalamnya. Tinjau bila `check:vendor` mengeluh. |
| `types.ts` | `src/types/specForm.ts` | Hanya satu tipe, ditulis manual. |

Salinan verbatim untuk keperluan diff ada di `../../vendor-upstream/`, beserta `manifest.json`
yang menyimpan sha256 tiap file sumber.

```bash
npm run check:vendor   # peringatkan kalau repo utama sudah berubah
npm run sync:vendor    # ambil ulang setelah perubahannya ditinjau
```

Kalau repo utama ada di lokasi lain, set `APLUS_MAIN_REPO`.
