/**
 * Moving blocks of cells between this grid and Excel.
 *
 * The tokenizer follows `parseClipboardTable` in the main APLUS repo, for the reason given there:
 * splitting each line on tabs is wrong for this data, because Excel wraps any cell holding a
 * newline or a tab in double quotes and writes the newline literally — and ASUS values are full
 * of them. IO Port, Expansion Slot and Include in the box are multi-line on nearly every model,
 * so a naive split turns one product into a dozen half-rows.
 *
 * One deliberate difference: the upstream version drops rows that are entirely blank, which suits
 * reading a datasheet. Here a paste is positional — row three of the clipboard lands on row three
 * of the selection — so dropping a blank row would silently shift everything below it onto the
 * wrong product. Blank rows are kept, and a blank cell clears its target.
 */

/** Split clipboard TSV into rows and cells, honouring Excel's quoting. */
export function parseClipboardTable(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cell = '';
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const ch = text[i];

    if (quoted) {
      if (ch === '"') {
        if (text[i + 1] === '"') {
          cell += '"';
          i += 1;
        } else quoted = false;
      } else cell += ch;
      continue;
    }

    if (ch === '"' && cell === '') {
      quoted = true;
    } else if (ch === '\t') {
      row.push(cell);
      cell = '';
    } else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i += 1;
      row.push(cell);
      rows.push(row);
      row = [];
      cell = '';
    } else {
      cell += ch;
    }
  }

  if (cell !== '' || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }

  // A trailing newline produces one empty row that was never really there.
  while (rows.length > 1 && rows[rows.length - 1].length === 1 && rows[rows.length - 1][0] === '') {
    rows.pop();
  }

  return rows;
}

/** Build clipboard TSV the way Excel does, so a copied block pastes back into Excel intact. */
export function buildClipboardTable(grid: string[][]): string {
  return grid
    .map((row) =>
      row
        .map((cell) => {
          const v = cell ?? '';
          return /[\t\n\r"]/.test(v) ? `"${v.replace(/"/g, '""')}"` : v;
        })
        .join('\t'),
    )
    .join('\n');
}
