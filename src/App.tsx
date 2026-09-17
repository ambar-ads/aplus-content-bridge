import { useMemo, useState } from 'react';
import { FileDrop } from './components/FileDrop';
import { Reconciliation } from './components/Reconciliation';
import { ProductEditor } from './components/ProductEditor';
import { EditorWindow } from './components/EditorWindow';
import { BASE_CATALOG } from './lib/baseCatalog';
import { FORWEB_FILENAME } from './lib/config';
import { serializeForWebCsv } from './lib/legacyCsv';
import { readPriceList } from './lib/priceListImport';
import { readMasterWorkbook, buildMasterWorkbook, MASTER_FILENAME } from './lib/masterWorkbook';
import {
  deriveForWebRow,
  retitleWith,
  titleNeedsRetitle,
  LEGACY_QUIRKS,
  DELIBERATE_FIXES,
} from './lib/deriveForWeb';
import { buildCatalogFromRows, mergePriceList, catalogToRows } from './lib/catalogStore';
import { assertXlsx, XlsxError } from './lib/xlsxGuard';
import { loadEdits, saveEdits, applyEdits, setField, setFields } from './lib/edits';
import { normalizeLabel } from './vendor/pmCell';
import type { EditMap, FieldChange } from './lib/edits';
import type { IncomingModel, MergeReport } from './lib/catalogStore';
import type { PriceListResult } from './lib/priceListImport';
import type { MasterReadResult } from './lib/masterWorkbook';
import type { ForWebKey, ForWebRow } from './lib/types';

function saveFile(name: string, data: BlobPart, mime: string) {
  const blob = new Blob([data], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

export default function App() {
  const [master, setMaster] = useState<MasterReadResult | null>(null);
  const [masterName, setMasterName] = useState('');
  const [priceList, setPriceList] = useState<PriceListResult | null>(null);
  const [priceListName, setPriceListName] = useState('');
  const [busy, setBusy] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);

  const [edits, setEdits] = useState<EditMap>(() => loadEdits());
  // While the pop-out editor is open its changes stay in a draft, so Cancel really cancels.
  const [draftEdits, setDraftEdits] = useState<EditMap | null>(null);
  // Corporate browsers often refuse window.open outright, so the same editor can also take over
  // the page instead. Same component, same draft — only where it is mounted differs.
  const [popout, setPopout] = useState<null | 'window' | 'overlay'>(null);

  function commitEdits(next: EditMap) {
    setEdits(next);
    saveEdits(next);
  }

  const windowOpen = popout !== null;
  const activeEdits = draftEdits ?? edits;

  function openPopout() {
    setDraftEdits(edits);
    setPopout('window');
  }

  function closePopout() {
    setPopout(null);
    setDraftEdits(null);
  }

  function savePopout() {
    if (draftEdits) commitEdits(draftEdits);
    closePopout();
  }

  /** Catalogue as it stands before any hand edits — both views are derived from this. */
  const base = useMemo(() => {
    const baseRows = master ? master.rows : catalogToRows(BASE_CATALOG);
    const { catalog: baseCatalog } = buildCatalogFromRows(baseRows, {
      origin: master ? 'manual' : 'seed',
      sourceFile: masterName || undefined,
    });

    let catalog = baseCatalog;
    let merge: MergeReport | null = null;
    const newPn90 = new Set<string>();

    if (priceList) {
      // The series table only ever names NEW products. Existing rows keep whatever Marketing
      // name is written in the master file.
      const series = new Map(
        (master?.series ?? []).map((s) => [normalizeLabel(s.code), s.marketingName]),
      );

      const incoming: IncomingModel[] = [];
      const deriveWarnings: string[] = [];
      for (const m of priceList.models) {
        if (!m.line) continue;
        const { row, warnings } = deriveForWebRow(m.record, { mktNameRule: series });
        deriveWarnings.push(...warnings);
        incoming.push({
          pn90: m.pn90,
          row,
          provenance: {
            origin: 'priceList',
            sourceFile: priceListName,
            sheetName: m.sheetName,
            productLine: m.sheetLine ?? undefined,
            group: m.group,
          },
        });
      }

      const merged = mergePriceList(catalog, incoming);
      catalog = merged.catalog;
      merge = { ...merged.report, warnings: [...merged.report.warnings, ...deriveWarnings] };
      merged.report.added.forEach((e) => newPn90.add(e.pn90));
    }

    return { rows: catalogToRows(catalog), merge, newPn90, baseCount: baseRows.length };
  }, [master, masterName, priceList, priceListName]);

  /** What the page shows and downloads: committed edits only. */
  const committed = useMemo(() => applyEdits(base.rows, edits), [base.rows, edits]);

  /** What the pop-out window shows: the draft, so Cancel can throw it away. */
  const shown = useMemo(
    () => (draftEdits ? applyEdits(base.rows, draftEdits) : committed),
    [base.rows, draftEdits, committed],
  );

  function editField(pn90: string, key: ForWebKey, value: string) {
    const next = setField(activeEdits, pn90, key, value);
    if (draftEdits) setDraftEdits(next);
    else commitEdits(next);
  }

  /** Pasting a column from Excel touches hundreds of cells; fold them into one update. */
  function editFields(changes: FieldChange[]) {
    if (!changes.length) return;
    const next = setFields(activeEdits, changes);
    if (draftEdits) setDraftEdits(next);
    else commitEdits(next);
  }

  /**
   * Fixing a wrong name means fixing two columns: Marketing name and Title. Doing that by hand
   * across 92 products is 184 edits, so it is offered as one action over everything out of sync.
   */
  function syncTitles(rows: ForWebRow[]) {
    let next = activeEdits;
    for (const row of rows) {
      if (!titleNeedsRetitle(row.title, row.marketingName)) continue;
      next = setField(next, row.partNumber.trim(), 'title', retitleWith(row.title, row.marketingName));
    }
    if (draftEdits) setDraftEdits(next);
    else commitEdits(next);
  }

  function discardAll() {
    if (draftEdits) setDraftEdits({});
    else commitEdits({});
  }

  async function readFile(file: File, kind: 'master' | 'priceList') {
    setBusy(kind === 'master' ? 'Reading master file…' : 'Reading price list…');
    setError(null);
    try {
      const buf = await file.arrayBuffer();
      assertXlsx(buf, file.name);
      if (kind === 'master') {
        setMaster(await readMasterWorkbook(buf));
        setMasterName(file.name);
      } else {
        setPriceList(await readPriceList(buf));
        setPriceListName(file.name);
      }
    } catch (e) {
      setError(
        e instanceof XlsxError
          ? e.message
          : `Could not read "${file.name}": ${e instanceof Error ? e.message : String(e)}`,
      );
    } finally {
      setBusy(null);
    }
  }

  async function downloadMaster() {
    setBusy('Building the Excel file…');
    try {
      const buffer = await buildMasterWorkbook({
        rows: committed.rows,
        series: master?.series ?? [],
        newPn90: base.newPn90,
      });
      saveFile(
        MASTER_FILENAME,
        buffer,
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      );
    } catch (e) {
      setError(`Could not build the Excel file: ${e instanceof Error ? e.message : String(e)}`);
    } finally {
      setBusy(null);
    }
  }

  const editedCount = Object.keys(committed.changedByPn90).length;

  return (
    <div className="wrap">
      <header className="app">
        <h1>A+ Content Bridge</h1>
        <p>
          Maintains <code className="inline">{MASTER_FILENAME}</code> and produces{' '}
          <code className="inline">{FORWEB_FILENAME}</code> for the{' '}
          <a href="https://asus-commercial.github.io/AplusContent/" target="_blank" rel="noreferrer">
            ASUS Business Product Spec Search
          </a>{' '}
          site.
        </p>
      </header>

      {error && <div className="note danger">{error}</div>}

      <section className="card">
        <h2>
          <span className={`step${master ? ' done' : ''}`}>1</span> Master file
        </h2>
        <p className="hint">
          <code className="inline">{MASTER_FILENAME}</code> — one row per product, no formulas.
          Don&rsquo;t have it yet? Skip this step and download the starting copy in step 5.
        </p>
        <div className="body">
          <FileDrop
            accept=".xlsx"
            fileName={masterName}
            onFile={(f) => readFile(f, 'master')}
            label={`Drop ${MASTER_FILENAME} here, or click to choose`}
          />
          <div className="stats">
            <div className="stat ok">
              <div className="n">{base.baseCount}</div>
              <div className="k">{master ? 'products in your file' : 'products built in'}</div>
            </div>
            <div className="stat">
              <div className="n">{master ? master.series.length : '—'}</div>
              <div className="k">series codes</div>
            </div>
          </div>
          {!master && (
            <div className="note ok">
              Nothing uploaded yet, so the built-in copy of the {BASE_CATALOG.entries.length}{' '}
              products currently on the site is being used.
            </div>
          )}
          {master && master.warnings.length > 0 && (
            <details>
              <summary>{master.warnings.length} note(s) while reading the master file</summary>
              <div className="note warn">
                <ul>
                  {master.warnings.map((w, i) => (
                    <li key={i}>{w}</li>
                  ))}
                </ul>
              </div>
            </details>
          )}
        </div>
      </section>

      <section className="card">
        <h2>
          <span className={`step${priceList ? ' done' : ''}`}>2</span> ABP Price List
        </h2>
        <p className="hint">
          Only needed when there are new products this month. Specifications of existing products
          are never changed.
        </p>
        <div className="body">
          <FileDrop
            accept=".xlsx"
            fileName={priceListName}
            onFile={(f) => readFile(f, 'priceList')}
            label="Drop ABP Price List.xlsx here, or click to choose"
          />
          {priceList && (
            <>
              <div className="scroll">
                <table>
                  <thead>
                    <tr>
                      <th>Datasheet sheet</th>
                      <th>Catalogue</th>
                      <th className="mono">Models</th>
                      <th className="mono">Labels</th>
                    </tr>
                  </thead>
                  <tbody>
                    {priceList.sheets.map((s) => (
                      <tr key={s.name}>
                        <td>{s.name}</td>
                        <td>{s.group === 'tkdn' ? 'TKDN' : 'Main'}</td>
                        <td className="mono">{s.models}</td>
                        <td className="mono">{s.labels}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
              {priceList.sheets.length === 0 && (
                <div className="note danger">
                  No datasheet sheets in this file. They are recognised by structure — A2 must be
                  &ldquo;Part No&rdquo; and A9 must be &ldquo;Sales Model Name&rdquo;.
                </div>
              )}
              {!master && (
                <div className="note warn">
                  No master file uploaded, so new products cannot be given a Marketing name
                  automatically. Upload the master file in step 1 first.
                </div>
              )}
            </>
          )}
        </div>
      </section>

      <section className="card">
        <h2>
          <span className="step done">3</span> Reconciliation
        </h2>
        <p className="hint">
          What Excel never showed: what was added, who is still missing copy, and which 90PNs have
          a problem.
        </p>
        <div className="body">
          <Reconciliation
            rows={committed.rows}
            merge={base.merge}
            baseCount={base.baseCount}
          />
        </div>
      </section>

      <section className="card">
        <h2>
          <span className={`step${editedCount ? ' done' : ''}`}>4</span> Edit
        </h2>
        <p className="hint">
          Small corrections don&rsquo;t need a round trip through Excel. Anything changed here goes
          into both files in the next step.
        </p>
        <div className="body">
          {master && editedCount > 0 && (
            <div className="note warn">
              Changes stored in this browser override the values in{' '}
              <code className="inline">{masterName}</code>. If that file already contains the
              corrections, discard the changes so they don&rsquo;t override newer data.
            </div>
          )}

          {windowOpen ? (
            <div className="note ok">
              The editor is open {popout === 'window' ? 'in a separate window' : 'full screen'}.
              Changes made there are applied when you choose <strong>Save changes</strong>.
            </div>
          ) : (
            <ProductEditor
              rows={committed.rows}
              changedByPn90={committed.changedByPn90}
              newPn90={base.newPn90}
              onChange={editField}
              onBulkChange={editFields}
              onRevertAll={discardAll}
              onSyncTitles={() => syncTitles(committed.rows)}
              onOpenWindow={openPopout}
            />
          )}
        </div>
      </section>

      <section className="card">
        <h2>
          <span className="step">5</span> Download
        </h2>
        <div className="body">
          <div className="row">
            <button disabled={busy !== null} onClick={downloadMaster}>
              {busy ? <span className="spinner" /> : null}
              Download {MASTER_FILENAME}
            </button>
            <button
              className="ghost"
              disabled={busy !== null}
              onClick={() =>
                saveFile(
                  FORWEB_FILENAME,
                  serializeForWebCsv(committed.rows),
                  'text/csv;charset=utf-8',
                )
              }
            >
              Download {FORWEB_FILENAME}
            </button>
          </div>

          <ol className="steps">
            <li>Save the Excel file over your own copy. That is what you start from next month.</li>
            <li>
              Upload the CSV to <code className="inline">ASUS-commercial/AplusContent</code>,
              replacing the old one. The filename must match <strong>exactly</strong>.
            </li>
            <li>Wait about 3 minutes, open the site, and search for one product to confirm.</li>
          </ol>

          <div className="note ok">
            The CSV keeps the same 25-column contract as before, UTF-8 with BOM. The site does not
            need any change to read it.
          </div>
        </div>
      </section>

      <section className="card">
        <h2>How the old system behaves</h2>
        <div className="body">
          <details>
            <summary>{LEGACY_QUIRKS.length} quirks reproduced on purpose</summary>
            <div className="note warn">
              <ul>
                {LEGACY_QUIRKS.map((q) => (
                  <li key={q.id}>
                    <strong>{q.affects}</strong> — {q.detail}
                  </li>
                ))}
              </ul>
            </div>
          </details>
          <details>
            <summary>{DELIBERATE_FIXES.length} things fixed on purpose</summary>
            <div className="note ok">
              <ul>
                {DELIBERATE_FIXES.map((f) => (
                  <li key={f.id}>{f.detail}</li>
                ))}
              </ul>
            </div>
          </details>
        </div>
      </section>

      {windowOpen &&
        (() => {
          const editor = (
            <ProductEditor
              inWindow
              rows={shown.rows}
              changedByPn90={shown.changedByPn90}
              newPn90={base.newPn90}
              onChange={editField}
              onBulkChange={editFields}
              onRevertAll={discardAll}
              onSyncTitles={() => syncTitles(shown.rows)}
              onSave={savePopout}
              onCancel={closePopout}
            />
          );

          if (popout === 'overlay') {
            return (
              <div className="editor-overlay" role="dialog" aria-label="Edit product data">
                {editor}
              </div>
            );
          }

          return (
            <EditorWindow
              title="Edit product data — A+ Content Bridge"
              onClose={closePopout}
              onBlocked={() => setPopout('overlay')}
            >
              {editor}
            </EditorWindow>
          );
        })()}
    </div>
  );
}
