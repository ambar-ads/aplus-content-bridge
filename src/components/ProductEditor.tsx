import { useMemo, useState } from 'react';
import { DataGrid } from './DataGrid';
import { EDITABLE_COLUMNS } from '../lib/edits';
import { isBlankValue } from '../lib/legacyCsv';
import { titleNeedsRetitle } from '../lib/deriveForWeb';
import type { ColumnFilters, SelectionInfo, SortState } from './DataGrid';
import type { FieldChange } from '../lib/edits';
import type { ForWebKey, ForWebRow } from '../lib/types';

interface Props {
  rows: ForWebRow[];
  changedByPn90: Record<string, ForWebKey[]>;
  newPn90: Set<string>;
  onChange: (pn90: string, key: ForWebKey, value: string) => void;
  onBulkChange: (changes: FieldChange[]) => void;
  onRevertAll: () => void;
  onSyncTitles: () => void;
  /** Rendered inside the pop-out window, which owns the Save / Cancel pair. */
  inWindow?: boolean;
  onOpenWindow?: () => void;
  onSave?: () => void;
  onCancel?: () => void;
}

type FilterId = 'all' | 'needsCopy' | 'notFound' | 'new' | 'changed';
type SortKey = ForWebKey | 'partNumber';

/** A filter box only offers a value list when the list is short enough to be worth reading. */
const MAX_SUGGESTIONS = 40;

function isNotFound(row: ForWebRow): boolean {
  return /not found/i.test(row.marketingName) || /not found/i.test(row.title);
}

function needsCopy(row: ForWebRow): boolean {
  return isBlankValue(row.tagline) || isBlankValue(row.ksp) || isBlankValue(row.link);
}

const text = (row: ForWebRow, key: SortKey) => String(row[key] ?? '').replace(/\n/g, ' ').trim();

/**
 * Toolbar, filters and sorting around the grid.
 *
 * The quick filters are the reconciliation findings, turned into something you can work through:
 * "92 products named Not found" stops being a number and becomes a queue. Column filters and
 * sorting sit on top of that, the way they would in a spreadsheet.
 */
export function ProductEditor({
  rows,
  changedByPn90,
  newPn90,
  onChange,
  onBulkChange,
  onRevertAll,
  onSyncTitles,
  inWindow = false,
  onOpenWindow,
  onSave,
  onCancel,
}: Props) {
  const [quick, setQuick] = useState<FilterId>('all');
  const [query, setQuery] = useState('');
  const [sort, setSort] = useState<SortState | null>(null);
  const [selection, setSelection] = useState<SelectionInfo | null>(null);
  const [columnFilters, setColumnFilters] = useState<ColumnFilters>({});

  const counts = useMemo(
    () => ({
      all: rows.length,
      needsCopy: rows.filter(needsCopy).length,
      notFound: rows.filter(isNotFound).length,
      new: rows.filter((r) => newPn90.has(r.partNumber.trim())).length,
      changed: Object.keys(changedByPn90).length,
    }),
    [rows, newPn90, changedByPn90],
  );

  const outOfSyncTitles = useMemo(
    () => rows.filter((r) => titleNeedsRetitle(r.title, r.marketingName)).length,
    [rows],
  );

  const suggestions = useMemo(() => {
    const out: Partial<Record<SortKey, string[]>> = {};
    for (const col of EDITABLE_COLUMNS) {
      const seen = new Set<string>();
      for (const row of rows) {
        const v = text(row, col.key);
        if (v) seen.add(v);
        if (seen.size > MAX_SUGGESTIONS) break;
      }
      if (seen.size > 1 && seen.size <= MAX_SUGGESTIONS) {
        out[col.key] = [...seen].sort((a, b) => a.localeCompare(b, 'en'));
      }
    }
    return out;
  }, [rows]);

  const visible = useMemo(() => {
    const q = query.trim().toLowerCase();
    const active = Object.entries(columnFilters).filter(([, v]) => (v ?? '').trim() !== '') as [
      SortKey,
      string,
    ][];

    const filtered = rows.filter((r) => {
      if (quick === 'needsCopy' && !needsCopy(r)) return false;
      if (quick === 'notFound' && !isNotFound(r)) return false;
      if (quick === 'new' && !newPn90.has(r.partNumber.trim())) return false;
      if (quick === 'changed' && !changedByPn90[r.partNumber.trim()]) return false;

      for (const [key, needle] of active) {
        if (!text(r, key).toLowerCase().includes(needle.trim().toLowerCase())) return false;
      }

      if (!q) return true;
      return (
        r.partNumber.toLowerCase().includes(q) ||
        r.productName.toLowerCase().includes(q) ||
        r.marketingName.toLowerCase().includes(q)
      );
    });

    if (!sort) return filtered;

    // Blank cells sink to the bottom whichever way the column is sorted — otherwise ascending
    // order buries every filled row under the empty ones.
    const dir = sort.dir === 'asc' ? 1 : -1;
    return [...filtered].sort((a, b) => {
      const av = text(a, sort.key);
      const bv = text(b, sort.key);
      if (!av && !bv) return 0;
      if (!av) return 1;
      if (!bv) return -1;
      return dir * av.localeCompare(bv, 'en', { numeric: true, sensitivity: 'base' });
    });
  }, [rows, quick, query, columnFilters, sort, newPn90, changedByPn90]);

  function cycleSort(key: SortKey) {
    setSort((prev) => {
      if (prev?.key !== key) return { key, dir: 'asc' };
      if (prev.dir === 'asc') return { key, dir: 'desc' };
      return null; // third click clears the sort and restores catalogue order
    });
  }

  const activeFilterCount = Object.values(columnFilters).filter((v) => (v ?? '').trim() !== '').length;

  const quickFilters: { id: FilterId; label: string }[] = [
    { id: 'all', label: 'All' },
    { id: 'needsCopy', label: 'Needs copy' },
    { id: 'notFound', label: '"Not found"' },
    { id: 'new', label: 'New' },
    { id: 'changed', label: 'Edited' },
  ];

  return (
    <div className={inWindow ? 'editor-shell in-window' : 'editor-shell'}>
      {inWindow && (
        <header className="window-head">
          <div>
            <strong>Edit product data</strong>
            <span className="window-sub">
              {counts.changed > 0
                ? `${counts.changed} product${counts.changed === 1 ? '' : 's'} edited`
                : 'No changes yet'}
            </span>
          </div>
          <div className="row">
            <button type="button" className="ghost" onClick={onCancel}>
              Cancel
            </button>
            <button type="button" onClick={onSave}>
              Save changes
            </button>
          </div>
        </header>
      )}

      <div className="editor-bar">
        <div className="bar-line">
          <input
            id="editor-search"
            className="editor-search"
            type="search"
            placeholder="Search 90PN, model name, or marketing name…"
            aria-label="Search products"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
          />
          {!inWindow && (
            <button type="button" className="ghost" onClick={onOpenWindow}>
              Open in new window
            </button>
          )}
        </div>

        <div className="bar-line">
          <div className="chips">
            {quickFilters.map((f) => (
              <button
                key={f.id}
                type="button"
                className={`chip${quick === f.id ? ' on' : ''}`}
                aria-pressed={quick === f.id}
                onClick={() => setQuick(f.id)}
              >
                {f.label}
                <span className="chip-n">{counts[f.id]}</span>
              </button>
            ))}
          </div>
          <div className="bar-actions">
            {(activeFilterCount > 0 || sort) && (
              <button
                type="button"
                className="ghost small"
                onClick={() => {
                  setColumnFilters({});
                  setSort(null);
                }}
              >
                Clear {activeFilterCount > 0 ? `${activeFilterCount} column filter${activeFilterCount === 1 ? '' : 's'}` : 'sort'}
              </button>
            )}
            {outOfSyncTitles > 0 && (
              <button type="button" className="ghost small" onClick={onSyncTitles}>
                Sync {outOfSyncTitles} Title{outOfSyncTitles === 1 ? '' : 's'} with Marketing name
              </button>
            )}
            {counts.changed > 0 && (
              <button type="button" className="ghost small" onClick={onRevertAll}>
                Discard all changes
              </button>
            )}
          </div>
        </div>
      </div>

      <DataGrid
        rows={visible}
        changedByPn90={changedByPn90}
        newPn90={newPn90}
        sort={sort}
        filters={columnFilters}
        suggestions={suggestions}
        onChange={onChange}
        onBulkChange={onBulkChange}
        onSelectionChange={setSelection}
        onSort={cycleSort}
        onFilter={(key, value) => setColumnFilters((prev) => ({ ...prev, [key]: value }))}
      />

      <div className="grid-status">
        <span>
          Showing <strong>{visible.length}</strong> of {rows.length} products
          {selection && (
            <>
              {' · '}
              <strong>
                {selection.rows} × {selection.columns}
              </strong>{' '}
              selected ({selection.cells.toLocaleString('en')} cells)
            </>
          )}
        </span>
        <span className="keys">
          <kbd>Click</kbd> select · <kbd>drag</kbd> or <kbd>Shift</kbd>+click extend ·{' '}
          <kbd>Ctrl</kbd>+<kbd>C</kbd>/<kbd>V</kbd> copy &amp; paste ·{' '}
          <kbd>Ctrl</kbd>+<kbd>D</kbd> fill down · <kbd>Del</kbd> clear ·{' '}
          <kbd>Enter</kbd> or double-click to edit
        </span>
      </div>

      <p className="editor-foot">
        Paste a column straight from Excel: select the first cell, press <kbd>Ctrl</kbd>+
        <kbd>V</kbd>, and the block fills downward from there. Copy one cell and paste it over a
        selection to give every row the same value.
        {outOfSyncTitles > 0 && (
          <>
            {' '}
            <strong>{outOfSyncTitles}</strong> Title
            {outOfSyncTitles === 1 ? ' still uses' : 's still use'} a different name from the
            Marketing name column — the sync button replaces just the name and leaves the
            specification in brackets untouched.
          </>
        )}
      </p>
    </div>
  );
}
