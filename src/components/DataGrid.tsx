import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { EDITABLE_COLUMNS } from '../lib/edits';
import { isBlankValue } from '../lib/legacyCsv';
import { MASTER_COLUMNS } from '../lib/masterWorkbook';
import { buildClipboardTable, parseClipboardTable } from '../lib/clipboard';
import type { FieldChange } from '../lib/edits';
import type { ForWebKey, ForWebRow } from '../lib/types';

export type SortDir = 'asc' | 'desc';
export interface SortState {
  key: ForWebKey | 'partNumber';
  dir: SortDir;
}
export type ColumnFilters = Partial<Record<ForWebKey | 'partNumber', string>>;

export interface SelectionInfo {
  rows: number;
  columns: number;
  cells: number;
}

interface Props {
  rows: ForWebRow[];
  changedByPn90: Record<string, ForWebKey[]>;
  newPn90: Set<string>;
  sort: SortState | null;
  filters: ColumnFilters;
  suggestions: Partial<Record<ForWebKey | 'partNumber', string[]>>;
  onChange: (pn90: string, key: ForWebKey, value: string) => void;
  onBulkChange: (changes: FieldChange[]) => void;
  onSort: (key: ForWebKey | 'partNumber') => void;
  onFilter: (key: ForWebKey | 'partNumber', value: string) => void;
  onSelectionChange?: (info: SelectionInfo | null) => void;
}

/** Fixed row height is what makes windowing possible, and it suits a spreadsheet grid anyway. */
const ROW_H = 30;
/** Rows rendered above and below the viewport, so fast scrolling does not flash empty space. */
const OVERSCAN = 8;
/** Width of the frozen 90PN column, needed to keep a keyboard-selected cell clear of it. */
const STICKY_W = 154;

const WIDTH_BY_KEY = new Map(MASTER_COLUMNS.map((c) => [c.key, c.width]));
const TALL = new Set(MASTER_COLUMNS.filter((c) => c.wrap).map((c) => c.key));

/** Excel column widths are in characters; this is the pixel equivalent, clamped to stay readable. */
function pxWidth(key: ForWebKey): number {
  const chars = WIDTH_BY_KEY.get(key) ?? 20;
  return Math.min(340, Math.max(96, Math.round(chars * 7.6)));
}

const COL_X: number[] = [];
{
  let x = STICKY_W;
  for (const c of EDITABLE_COLUMNS) {
    COL_X.push(x);
    x += pxWidth(c.key);
  }
}

interface Cell {
  r: number;
  c: number;
}
interface Range {
  r1: number;
  c1: number;
  r2: number;
  c2: number;
}

const bounds = (s: Range) => ({
  top: Math.min(s.r1, s.r2),
  bottom: Math.max(s.r1, s.r2),
  left: Math.min(s.c1, s.c2),
  right: Math.max(s.c1, s.c2),
});

const clamp = (n: number, lo: number, hi: number) => Math.max(lo, Math.min(hi, n));

/**
 * A wide, spreadsheet-style grid: every column visible, edited in place, selected in blocks.
 *
 * Two things make this workable at 400+ products by 23 columns. Only the rows inside the viewport
 * are rendered, so the table stays responsive however long the catalogue grows. And only the cell
 * being edited becomes an input — the rest are plain text — because ten thousand live inputs is
 * what makes grids like this crawl.
 *
 * Selection follows spreadsheet convention rather than the earlier click-to-edit: a single click
 * selects, and editing starts on double click, Enter, F2, or simply typing. Click-to-edit and
 * drag-to-select cannot both own the mouse down event, and filling a column of Taglines one click
 * at a time is the thing this is meant to stop.
 *
 * Copy and paste go through the browser's own clipboard events rather than `navigator.clipboard`,
 * which needs a permission the async API does not always get in a corporate browser. Ctrl+C and
 * Ctrl+V fire these events natively, so they work everywhere the page does.
 */
export function DataGrid({
  rows,
  changedByPn90,
  newPn90,
  sort,
  filters,
  suggestions,
  onChange,
  onBulkChange,
  onSort,
  onFilter,
  onSelectionChange,
}: Props) {
  const scroller = useRef<HTMLDivElement>(null);
  const [scrollTop, setScrollTop] = useState(0);
  const [viewportH, setViewportH] = useState(560);
  const [editing, setEditing] = useState<Cell | null>(null);
  const [draft, setDraft] = useState('');
  const [anchor, setAnchor] = useState<Cell | null>(null);
  const [sel, setSel] = useState<Range | null>(null);
  const dragging = useRef(false);

  useEffect(() => {
    const el = scroller.current;
    if (!el) return;
    const measure = () => setViewportH(el.clientHeight);
    measure();
    const ro = new ResizeObserver(measure);
    ro.observe(el);
    return () => ro.disconnect();
  }, []);

  useEffect(() => {
    const stop = () => {
      dragging.current = false;
    };
    window.addEventListener('mouseup', stop);
    return () => window.removeEventListener('mouseup', stop);
  }, []);

  // Selection is held as indices into the rows currently on screen, so re-sorting or filtering
  // would leave it pointing at different products. Drop it rather than silently move it.
  useEffect(() => {
    setSel(null);
    setAnchor(null);
    setEditing(null);
  }, [rows, sort, filters]);

  const info = useMemo<SelectionInfo | null>(() => {
    if (!sel) return null;
    const b = bounds(sel);
    const r = b.bottom - b.top + 1;
    const c = b.right - b.left + 1;
    return { rows: r, columns: c, cells: r * c };
  }, [sel]);

  useEffect(() => {
    onSelectionChange?.(info);
  }, [info, onSelectionChange]);

  const start = Math.max(0, Math.floor(scrollTop / ROW_H) - OVERSCAN);
  const end = Math.min(rows.length, Math.ceil((scrollTop + viewportH) / ROW_H) + OVERSCAN);
  const visible = rows.slice(start, end);

  const valueAt = useCallback(
    (r: number, c: number) => String(rows[r]?.[EDITABLE_COLUMNS[c].key] ?? ''),
    [rows],
  );

  /** Keep a keyboard-moved cell on screen, past the frozen column and the sticky header. */
  const reveal = useCallback((r: number, c: number) => {
    const el = scroller.current;
    if (!el) return;
    const top = r * ROW_H;
    const headerH = 55;
    if (top < el.scrollTop + headerH) el.scrollTop = Math.max(0, top - headerH);
    else if (top + ROW_H > el.scrollTop + el.clientHeight) {
      el.scrollTop = top + ROW_H - el.clientHeight;
    }
    const x = COL_X[c];
    const w = pxWidth(EDITABLE_COLUMNS[c].key);
    if (x - STICKY_W < el.scrollLeft) el.scrollLeft = Math.max(0, x - STICKY_W);
    else if (x + w > el.scrollLeft + el.clientWidth) el.scrollLeft = x + w - el.clientWidth;
  }, []);

  const beginEdit = useCallback(
    (r: number, c: number, initial?: string) => {
      const raw = valueAt(r, c);
      setEditing({ r, c });
      // "0" is the legacy blank sentinel; never put it in front of someone as editable text.
      setDraft(initial !== undefined ? initial : raw.trim() === '0' ? '' : raw);
    },
    [valueAt],
  );

  const commit = useCallback(() => {
    if (!editing) return;
    const row = rows[editing.r];
    if (row) onChange(row.partNumber.trim(), EDITABLE_COLUMNS[editing.c].key, draft);
    setEditing(null);
    scroller.current?.focus();
  }, [editing, draft, rows, onChange]);

  const cancel = useCallback(() => {
    setEditing(null);
    scroller.current?.focus();
  }, []);

  const select = useCallback((r: number, c: number, extend: boolean) => {
    setSel((prev) =>
      extend && prev ? { ...prev, r2: r, c2: c } : { r1: r, c1: c, r2: r, c2: c },
    );
    if (!extend) setAnchor({ r, c });
  }, []);

  /** Every cell in the current selection, as changes ready to apply in one go. */
  const changesOver = useCallback(
    (fn: (r: number, c: number) => string | null): FieldChange[] => {
      if (!sel) return [];
      const b = bounds(sel);
      const out: FieldChange[] = [];
      for (let r = b.top; r <= b.bottom; r += 1) {
        const row = rows[r];
        if (!row) continue;
        for (let c = b.left; c <= b.right; c += 1) {
          const value = fn(r, c);
          if (value === null) continue;
          out.push({ pn90: row.partNumber.trim(), key: EDITABLE_COLUMNS[c].key, value });
        }
      }
      return out;
    },
    [sel, rows],
  );

  const clearSelection = useCallback(() => {
    onBulkChange(changesOver(() => ''));
  }, [changesOver, onBulkChange]);

  /** Ctrl+D — take the top row of the selection and repeat it down the rest. */
  const fillDown = useCallback(() => {
    if (!sel) return;
    const b = bounds(sel);
    if (b.bottom === b.top) return;
    onBulkChange(
      changesOver((r, c) => (r === b.top ? null : valueAt(b.top, c))),
    );
  }, [sel, changesOver, onBulkChange, valueAt]);

  function handleCopy(e: React.ClipboardEvent) {
    if (!sel || editing) return;
    const b = bounds(sel);
    const grid: string[][] = [];
    for (let r = b.top; r <= b.bottom; r += 1) {
      const line: string[] = [];
      for (let c = b.left; c <= b.right; c += 1) {
        const v = valueAt(r, c);
        line.push(isBlankValue(v) ? '' : v);
      }
      grid.push(line);
    }
    e.clipboardData.setData('text/plain', buildClipboardTable(grid));
    e.preventDefault();
  }

  /**
   * Paste lands positionally from the anchor. A single copied cell fills the whole selection —
   * that is how a column of identical Links gets filled in one action.
   */
  function handlePaste(e: React.ClipboardEvent) {
    if (!sel || editing) return;
    const text = e.clipboardData.getData('text/plain');
    if (!text) return;
    e.preventDefault();

    const table = parseClipboardTable(text);
    if (!table.length) return;

    const b = bounds(sel);
    const single = table.length === 1 && table[0].length === 1;
    const height = single ? b.bottom - b.top + 1 : table.length;
    const width = single ? b.right - b.left + 1 : Math.max(...table.map((t) => t.length));

    const out: FieldChange[] = [];
    for (let dr = 0; dr < height; dr += 1) {
      const r = b.top + dr;
      const row = rows[r];
      if (!row) break;
      for (let dc = 0; dc < width; dc += 1) {
        const c = b.left + dc;
        if (c >= EDITABLE_COLUMNS.length) break;
        const value = single ? table[0][0] : (table[dr]?.[dc] ?? '');
        out.push({ pn90: row.partNumber.trim(), key: EDITABLE_COLUMNS[c].key, value });
      }
    }

    onBulkChange(out);
    setSel({
      r1: b.top,
      c1: b.left,
      r2: Math.min(rows.length - 1, b.top + height - 1),
      c2: Math.min(EDITABLE_COLUMNS.length - 1, b.left + width - 1),
    });
  }

  function handleKeyDown(e: React.KeyboardEvent) {
    if (editing) return;
    const a = anchor;
    const maxR = rows.length - 1;
    const maxC = EDITABLE_COLUMNS.length - 1;
    const mod = e.ctrlKey || e.metaKey;

    if (mod && e.key.toLowerCase() === 'd') {
      e.preventDefault();
      fillDown();
      return;
    }
    if (mod && e.key.toLowerCase() === 'a') {
      e.preventDefault();
      setAnchor({ r: 0, c: 0 });
      setSel({ r1: 0, c1: 0, r2: maxR, c2: maxC });
      return;
    }
    if (e.key === 'Escape') {
      setSel(null);
      setAnchor(null);
      return;
    }
    if (!a) return;

    if (e.key === 'Delete' || e.key === 'Backspace') {
      e.preventDefault();
      clearSelection();
      return;
    }
    if (e.key === 'Enter' || e.key === 'F2') {
      e.preventDefault();
      beginEdit(a.r, a.c);
      return;
    }

    const move = (dr: number, dc: number) => {
      e.preventDefault();
      const r = clamp(a.r + dr, 0, maxR);
      const c = clamp(a.c + dc, 0, maxC);
      if (e.shiftKey) setSel((prev) => (prev ? { ...prev, r2: r, c2: c } : null));
      else {
        setAnchor({ r, c });
        setSel({ r1: r, c1: c, r2: r, c2: c });
      }
      reveal(r, c);
    };

    if (e.key === 'ArrowUp') return move(-1, 0);
    if (e.key === 'ArrowDown') return move(1, 0);
    if (e.key === 'ArrowLeft') return move(0, -1);
    if (e.key === 'ArrowRight') return move(0, 1);
    if (e.key === 'Tab') return move(0, e.shiftKey ? -1 : 1);
    if (e.key === 'Home') return move(0, -maxC);
    if (e.key === 'End') return move(0, maxC);
    if (e.key === 'PageUp') return move(-Math.floor(viewportH / ROW_H), 0);
    if (e.key === 'PageDown') return move(Math.floor(viewportH / ROW_H), 0);

    // Typing over a selected cell replaces it, the way a spreadsheet does.
    if (!mod && !e.altKey && e.key.length === 1) {
      e.preventDefault();
      beginEdit(a.r, a.c, e.key);
    }
  }

  const sortMark = (key: ForWebKey | 'partNumber') =>
    sort?.key === key ? (sort.dir === 'asc' ? '▲' : '▼') : '';

  const headCell = (key: ForWebKey | 'partNumber', label: string, sticky: boolean) => (
    <th
      key={key}
      className={sticky ? 'sticky-col' : undefined}
      aria-sort={sort?.key === key ? (sort.dir === 'asc' ? 'ascending' : 'descending') : 'none'}
    >
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

  const b = sel ? bounds(sel) : null;

  return (
    <div
      className="grid-scroll"
      ref={scroller}
      tabIndex={0}
      onScroll={(e) => setScrollTop((e.target as HTMLDivElement).scrollTop)}
      onKeyDown={handleKeyDown}
      onCopy={handleCopy}
      onPaste={handlePaste}
    >
      <table className="grid">
        <colgroup>
          <col style={{ width: STICKY_W }} />
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

          {visible.map((row, i) => {
            const r = start + i;
            const pn90 = row.partNumber.trim();
            const changed = new Set(changedByPn90[pn90] ?? []);
            const rowSelected = b !== null && r >= b.top && r <= b.bottom;
            return (
              <tr key={pn90} style={{ height: ROW_H }}>
                <th className={`sticky-col${rowSelected ? ' in-selection' : ''}`} scope="row">
                  <span
                    className={`dot${changed.size ? ' changed' : newPn90.has(pn90) ? ' new' : ''}`}
                  />
                  {pn90}
                </th>
                {EDITABLE_COLUMNS.map((col, c) => {
                  const value = row[col.key] ?? '';
                  const isEditing = editing?.r === r && editing.c === c;
                  const blank = isBlankValue(value);
                  const flag =
                    blank && (col.key === 'tagline' || col.key === 'ksp' || col.key === 'link');
                  const selected = rowSelected && b !== null && c >= b.left && c <= b.right;
                  const isAnchor = anchor?.r === r && anchor.c === c;

                  if (isEditing) {
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
                      className={
                        'cell' +
                        (changed.has(col.key) ? ' changed' : '') +
                        (flag ? ' flag' : '') +
                        (selected ? ' selected' : '') +
                        (isAnchor ? ' anchor' : '')
                      }
                      title={value.length > 40 ? value : undefined}
                      onMouseDown={(e) => {
                        if (e.button !== 0) return;
                        e.preventDefault();
                        scroller.current?.focus();
                        if (e.shiftKey && anchor) select(r, c, true);
                        else {
                          select(r, c, false);
                          dragging.current = true;
                        }
                      }}
                      onMouseEnter={() => {
                        // Guarded on the ref alone. Checking the `anchor` state here loses the
                        // first cell of a fast drag, because the mousedown that sets it has not
                        // flushed by the time the pointer has already moved on.
                        if (dragging.current) select(r, c, true);
                      }}
                      onDoubleClick={() => beginEdit(r, c)}
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
