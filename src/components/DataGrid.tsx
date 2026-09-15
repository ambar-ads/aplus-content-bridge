import { useCallback, useEffect, useRef, useState } from 'react';
import { EDITABLE_COLUMNS } from '../lib/edits';
import { isBlankValue } from '../lib/legacyCsv';
import { MASTER_COLUMNS } from '../lib/masterWorkbook';
import type { ForWebKey, ForWebRow } from '../lib/types';

export type SortDir = 'asc' | 'desc';
export interface SortState {
  key: ForWebKey | 'partNumber';
  dir: SortDir;
}
export type ColumnFilters = Partial<Record<ForWebKey | 'partNumber', string>>;

interface Props {
  rows: ForWebRow[];
  changedByPn90: Record<string, ForWebKey[]>;
  newPn90: Set<string>;
  sort: SortState | null;
  filters: ColumnFilters;
  /** Distinct values per column, for the filter boxes' autocomplete. */
  suggestions: Partial<Record<ForWebKey | 'partNumber', string[]>>;
  onChange: (pn90: string, key: ForWebKey, value: string) => void;
  onSort: (key: ForWebKey | 'partNumber') => void;
  onFilter: (key: ForWebKey | 'partNumber', value: string) => void;
}

/** Fixed row height is what makes windowing possible, and it suits a spreadsheet grid anyway. */
const ROW_H = 30;
/** Rows rendered above and below the viewport, so fast scrolling does not flash empty space. */
const OVERSCAN = 8;

const WIDTH_BY_KEY = new Map(MASTER_COLUMNS.map((c) => [c.key, c.width]));
const TALL = new Set(MASTER_COLUMNS.filter((c) => c.wrap).map((c) => c.key));

/** Excel column widths are in characters; this is the pixel equivalent, clamped to stay readable. */
function pxWidth(key: ForWebKey): number {
  const chars = WIDTH_BY_KEY.get(key) ?? 20;
  return Math.min(340, Math.max(96, Math.round(chars * 7.6)));
}

interface Cell {
  pn90: string;
  key: ForWebKey;
}

/**
 * A wide, spreadsheet-style grid: every column visible, edited in place.
 *
 * Two things make this workable at 400+ products by 23 columns. Only the rows inside the viewport
 * are rendered, so the table stays responsive however long the catalogue grows. And only the cell
 * being edited becomes an input — the rest are plain text — because ten thousand live inputs is
 * what makes grids like this crawl.
 *
 * Sorting and filtering are decided by the parent and arrive as props, so the row count shown in
 * the footer and the rows drawn here can never disagree.
 */
export function DataGrid({
  rows,
  changedByPn90,
  newPn90,
  sort,
  filters,
  suggestions,
  onChange,
  onSort,
  onFilter,
}: Props) {
  const scroller = useRef<HTMLDivElement>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportH, setViewportH] = useState(560);
  const [active, setActive] = useState<Cell | null>(null);
  const [draft, setDraft] = useState('');

  useEffect(() => {
    const el = scroller.current;
    if (!el) return;
    const measure = () => setViewportH(el.clientHeight);
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  const start = Math.max(0, Math.floor(scrollTop / ROW_H) - OVERSCAN);
  const end = Math.min(rows.length, Math.ceil((scrollTop + viewportH) / ROW_H) + OVERSCAN);
  const visible = rows.slice(start, end);

  const beginEdit = useCallback((pn90: string, key: ForWebKey, value: string) => {
    setActive({ pn90, key });
    // "0" is the legacy blank sentinel; never put it in front of someone as editable text.
    setDraft(value.trim() === '0' ? '' : value);
  }, []);

  const commit = useCallback(() => {
    if (!active) return;
    onChange(active.pn90, active.key, draft);
    setActive(null);
  }, [active, draft, onChange]);

  const cancel = useCallback(() => setActive(null), []);

  const sortMark = (key: ForWebKey | 'partNumber') =>
    sort?.key === key ? (sort.dir === 'asc' ? '▲' : '▼') : '';

  const headCell = (key: ForWebKey | 'partNumber', label: string, sticky: boolean) => (
    <th key={key} className={sticky ? 'sticky-col' : undefined} aria-sort={
      sort?.key === key ? (sort.dir === 'asc' ? 'ascending' : 'descending') : 'none'
    }>
      <button type="button" className="sort-btn" onClick={() => onSort(key)}>
        <span className="sort-label">{label}</span>
        <span className="sort-mark">{sortMark(key)}</span>
      </button>
    </th>
  );

  const filterCell = (key: ForWebKey | 'partNumber', sticky: boolean) => {
    const listId = `sug-${key}`;
    const options = suggestions[key];
    return (
      <th key={key} className={`filter-th${sticky ? ' sticky-col' : ''}`}>
        <input
          type="text"
          className={filters[key] ? 'on' : undefined}
          value={filters[key] ?? ''}
          placeholder="Filter…"
          aria-label={`Filter by ${key}`}
          list={options ? listId : undefined}
          onChange={(e) => onFilter(key, e.target.value)}
        />
        {options && (
          <datalist id={listId}>
            {options.map((o) => (
              <option key={o} value={o} />
            ))}
          </datalist>
        )}
      </th>
    );
  };

  return (
    <div
      className="grid-scroll"
      ref={scroller}
      onScroll={(e) => setScrollTop((e.target as HTMLDivElement).scrollTop)}
    >
      <table className="grid">
        <colgroup>
          <col style={{ width: 154 }} />
          {EDITABLE_COLUMNS.map((c) => (
            <col key={c.key} style={{ width: pxWidth(c.key) }} />
          ))}
        </colgroup>
        <thead>
          <tr className="head-row">
            {headCell('partNumber', '90PN', true)}
            {EDITABLE_COLUMNS.map((c) => headCell(c.key, c.header, false))}
          </tr>
          <tr className="filter-row">
            {filterCell('partNumber', true)}
            {EDITABLE_COLUMNS.map((c) => filterCell(c.key, false))}
          </tr>
        </thead>
        <tbody>
          {start > 0 && <tr style={{ height: start * ROW_H }} aria-hidden="true" />}

          {visible.length === 0 && (
            <tr>
              <td className="empty" colSpan={EDITABLE_COLUMNS.length + 1}>
                No products match the current search and filters.
              </td>
            </tr>
          )}

          {visible.map((row) => {
            const pn90 = row.partNumber.trim();
            const changed = new Set(changedByPn90[pn90] ?? []);
            return (
              <tr key={pn90} style={{ height: ROW_H }}>
                <th className="sticky-col" scope="row">
                  <span
                    className={`dot${changed.size ? ' changed' : newPn90.has(pn90) ? ' new' : ''}`}
                  />
                  {pn90}
                </th>
                {EDITABLE_COLUMNS.map((col) => {
                  const value = row[col.key] ?? '';
                  const editing = active?.pn90 === pn90 && active.key === col.key;
                  const blank = isBlankValue(value);
                  const flag =
                    blank && (col.key === 'tagline' || col.key === 'ksp' || col.key === 'link');

                  if (editing) {
                    return (
                      <td key={col.key} className="editing">
                        {TALL.has(col.key) ? (
                          <textarea
                            autoFocus
                            value={draft}
                            onChange={(e) => setDraft(e.target.value)}
                            onBlur={commit}
                            onKeyDown={(e) => {
                              if (e.key === 'Escape') cancel();
                            }}
                          />
                        ) : (
                          <input
                            autoFocus
                            type="text"
                            value={draft}
                            onChange={(e) => setDraft(e.target.value)}
                            onBlur={commit}
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') commit();
                              if (e.key === 'Escape') cancel();
                            }}
                          />
                        )}
                      </td>
                    );
                  }

                  return (
                    <td
                      key={col.key}
                      tabIndex={0}
                      role="button"
                      className={`cell${changed.has(col.key) ? ' changed' : ''}${flag ? ' flag' : ''}`}
                      title={value.length > 40 ? value : undefined}
                      onClick={() => beginEdit(pn90, col.key, value)}
                      onFocus={() => beginEdit(pn90, col.key, value)}
                    >
                      {blank ? '' : value.replace(/\n/g, ' · ')}
                    </td>
                  );
                })}
              </tr>
            );
          })}

          {end < rows.length && (
            <tr style={{ height: (rows.length - end) * ROW_H }} aria-hidden="true" />
          )}
        </tbody>
      </table>
    </div>
  );
}
