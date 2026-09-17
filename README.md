# A+ Content Bridge

Maintains `A+ Content Master.xlsx` and produces `A+ Content Source(ForWeb).csv` for the
[ASUS Business Product Spec Search](https://asus-commercial.github.io/AplusContent/) site.

**Demo:** [the tool](https://ambar-ads.github.io/aplus-content-bridge/) · [the A+ Content page](https://ambar-ads.github.io/aplus-content-bridge/site/)

Everything runs in the browser. No server, no database, no upload — the Excel files you pick are
read locally and never leave your machine.

---

## Running it

```bash
npm install
npm run dev            # http://localhost:5180
```

To build the published bundle:

```bash
npm run build:demo     # tool at /, the A+ Content page at /site/
```

---

## The monthly flow

```
A+ Content Master.xlsx ──┐
                         ├──►  A+ Content Bridge  ──►  Master.xlsx (updated)
Monthly price list    ───┘        (in the browser)  └─►  ForWeb.csv  ──► GitHub ──► site
   (only when there are new products)
```

1. **Master file** — upload `A+ Content Master.xlsx`. Don't have it yet? Skip it; a built-in copy
   of the products currently on the site is used, and you can download the starting file in step 5.
2. **Price list** — only needed when there are new products this month. Specifications of existing
   products are never changed, and products missing from this month's list are never dropped.
3. **Reconciliation** — what was added, who is still missing copy, which 90PNs have a problem.
4. **Edit** — a spreadsheet-style grid. Sort and filter per column, select a block of cells, and
   paste a column straight out of Excel. Or open it in a separate window with Save / Cancel.

   | | |
   |---|---|
   | Click, drag, Shift+click | select a cell or a block |
   | Arrows, Shift+arrows | move or extend the selection |
   | Ctrl+C / Ctrl+V | copy and paste, tab-separated like Excel |
   | Ctrl+D | fill the top row of the selection downward |
   | Ctrl+A | select everything |
   | Delete | clear the selected cells |
   | Enter, F2, double-click, or just typing | edit a cell |

   Paste is positional: row three of the clipboard lands on row three of the selection. Copy a
   single cell and paste it over a block to give every row the same value — which is how a column
   of identical Links gets filled in one action.
5. **Download** — save the Excel file over your own copy, and upload the CSV to the
   `AplusContent` repository, replacing the old one. The filename must match exactly.

---

## Commands

| Command | What it does |
|---|---|
| `npm run dev` | The tool, port 5180 |
| `npm run dev:legacy` | A copy of the A+ Content page, port 5181 |
| `npm run build:demo` | Build the published bundle |
| `npm run check` | Run every check |
| `npm run check:forweb` | **The safety gate** — the generated CSV must match the published one byte for byte |
| `npm run build:master` | Generate `A+ Content Master.xlsx` and verify nothing was lost |
| `npm run dryrun` | Run the whole monthly flow from the terminal |
| `npm run emit` | Write the CSV into `legacy-site/` to try before uploading |
| `npm run patch:page` | Re-apply the site changes from the untouched original |
| `npm run check:derive` | Accuracy report for the derivation rules |
| `npm run seed:catalog` | Rebuild `data/catalog.json` from the published CSV |
| `npm run check:vendor` | Warn when the main APLUS repo has moved on |

`check:forweb` is the important one: it proves the published text cannot change by accident.

---

## Layout

```
assets/            input workbooks (not committed — ASUS internal)
data/catalog.json  starting catalogue, used to generate the first master file
legacy-site/       a copy of the A+ Content page for trying the CSV locally
  .original/       the untouched page from GitHub, for comparison
src/lib/           parsing, derivation, merge, CSV, the master workbook
src/vendor/        code copied from the main APLUS repo
```

The project is self-contained. It only ever **reads** the main APLUS repo, via
`npm run sync:vendor`.

---

## Publishing

`.github/workflows/deploy.yml` builds and publishes to GitHub Pages on every push to `main`.
Enable it once under Settings → Pages → Source: **GitHub Actions**.

`base` in `vite.config.ts` is `./`, so the build runs from any path — a domain root or a
`/repo-name/` subpath — without rebuilding.
