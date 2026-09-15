import { isBlankValue } from '../lib/legacyCsv';
import type { MergeReport } from '../lib/catalogStore';
import { FORWEB_COLUMNS } from '../lib/config';
import type { ForWebKey, ForWebRow } from '../lib/types';

/** Excel writes these into a cell when a formula fails, and the old pipeline published them. */
const EXCEL_ERROR = /#(VALUE|REF|N\/A|NAME|DIV\/0|NUM|NULL)[!?]?/;
const EXCEL_ERROR_COLUMNS: ForWebKey[] = FORWEB_COLUMNS.map((c) => c.key).filter(
  (k) => k !== 'partNumber' && k !== 'partNumberAuto',
);
const HEADER_BY_KEY = new Map(FORWEB_COLUMNS.map((c) => [c.key, c.header]));

interface Props {
  rows: ForWebRow[];
  merge: MergeReport | null;
  baseCount: number;
}

/**
 * The reconciliation report.
 *
 * In the old workbook everything was joined by spill formulas that reported nothing — products
 * could shift or lose their copy with no sign at all. This panel makes those things visible
 * before the CSV is uploaded.
 */
export function Reconciliation({ rows, merge, baseCount }: Props) {
  const total = rows.length;

  // The old system wrote the literal text "Not found" whenever the marketing name lookup failed,
  // and that text went live.
  const notFound = rows.filter(
    (r) => /not found/i.test(r.marketingName) || /not found/i.test(r.title),
  );
  const noCopy = rows.filter(
    (r) => isBlankValue(r.tagline) && isBlankValue(r.ksp) && isBlankValue(r.link),
  );
  const noLink = rows.filter((r) => isBlankValue(r.link));
  const panelNotFound = rows.filter((r) => /not found/i.test(r.panelSize));

  // Excel error text that the old formulas wrote straight into the published CSV. It is showing
  // on the site right now, so it is worth naming rather than leaving to be noticed by a customer.
  const withExcelErrors = rows
    .map((r) => ({
      row: r,
      columns: EXCEL_ERROR_COLUMNS.filter((k) => EXCEL_ERROR.test(String(r[k] ?? ''))),
    }))
    .filter((x) => x.columns.length > 0);

  const nameCount = new Map<string, number>();
  rows.forEach((r) => nameCount.set(r.productName, (nameCount.get(r.productName) ?? 0) + 1));
  const dupNames = [...nameCount.entries()].filter(([, c]) => c > 1);

  return (
    <>
      <div className="stats">
        <div className="stat ok">
          <div className="n">{total}</div>
          <div className="k">products in CSV</div>
        </div>
        <div className="stat ok">
          <div className="n">+{merge ? merge.added.length : 0}</div>
          <div className="k">new products</div>
        </div>
        <div className="stat">
          <div className="n">{merge ? merge.notInSource : baseCount}</div>
          <div className="k">carried over</div>
        </div>
        <div className={noCopy.length ? 'stat warn' : 'stat'}>
          <div className="n">{noCopy.length}</div>
          <div className="k">no copy at all</div>
        </div>
      </div>

      {total < baseCount && (
        <div className="note danger">
          The product count dropped from {baseCount} to {total}. That should never happen — check
          before uploading anything.
        </div>
      )}

      {merge && merge.added.length > 0 && (
        <details open>
          <summary>{merge.added.length} new products from the price list</summary>
          <div className="scroll">
            <table>
              <thead>
                <tr>
                  <th className="mono">90PN</th>
                  <th>Product Name</th>
                  <th>Marketing name</th>
                  <th>Copy</th>
                </tr>
              </thead>
              <tbody>
                {merge.added.map((e) => {
                  const has =
                    !isBlankValue(e.row.tagline) ||
                    !isBlankValue(e.row.ksp) ||
                    !isBlankValue(e.row.link);
                  return (
                    <tr key={e.pn90}>
                      <td className="mono">{e.pn90}</td>
                      <td>{e.row.productName}</td>
                      <td>{e.row.marketingName}</td>
                      <td style={{ color: has ? 'var(--ok)' : 'var(--warn)' }}>
                        {has ? 'complete' : 'missing'}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </details>
      )}

      {noCopy.length > 0 && (
        <details>
          <summary>{noCopy.length} products with no Tagline, KSP or Link</summary>
          <div className="note warn">
            The site will show &ldquo;No Available&rdquo; for these, and the Materials button will
            say &ldquo;Coming Soon&rdquo;. Fill the Tagline, KSP and Link columns in step 4, or in
            the master file where those cells are highlighted yellow.
          </div>
          <div className="scroll">
            <table>
              <thead>
                <tr>
                  <th className="mono">90PN</th>
                  <th>Product Name</th>
                </tr>
              </thead>
              <tbody>
                {noCopy.map((r) => (
                  <tr key={r.partNumber}>
                    <td className="mono">{r.partNumber}</td>
                    <td>{r.productName}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </details>
      )}

      {notFound.length > 0 && (
        <details>
          <summary>{notFound.length} products showing &ldquo;Not found&rdquo; as their name</summary>
          <div className="note danger">
            The marketing name lookup failed and the old system wrote the literal text &ldquo;Not
            found&rdquo; into both Marketing name and Title, so it went live. For desktops the
            lookup used a 3-character key against a table keyed on 5 characters, which can never
            match. Fix them in step 4: type the real name, then use the Title sync button.
          </div>
          <div className="scroll">
            <table>
              <thead>
                <tr>
                  <th className="mono">90PN</th>
                  <th>Product Name</th>
                  <th>Marketing name</th>
                </tr>
              </thead>
              <tbody>
                {notFound.slice(0, 100).map((r) => (
                  <tr key={r.partNumber}>
                    <td className="mono">{r.partNumber}</td>
                    <td>{r.productName}</td>
                    <td>{r.marketingName}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </details>
      )}

      {withExcelErrors.length > 0 && (
        <details>
          <summary>
            {withExcelErrors.length} products showing Excel error text such as #VALUE!
          </summary>
          <div className="note danger">
            These are live on the site right now. They come from the old <code className="inline">
            Auto-*</code> formulas failing and writing the error straight into the CSV — usually
            after ASUS moved a row in the datasheet, which the formulas referenced by column
            letter. Overwrite the affected cells in step 4, or import a price list that covers
            these products and refresh them.
          </div>
          <div className="scroll">
            <table>
              <thead>
                <tr>
                  <th className="mono">90PN</th>
                  <th>Product Name</th>
                  <th>Columns affected</th>
                </tr>
              </thead>
              <tbody>
                {withExcelErrors.slice(0, 100).map(({ row, columns }) => (
                  <tr key={row.partNumber}>
                    <td className="mono">{row.partNumber}</td>
                    <td>{row.productName}</td>
                    <td>{columns.map((k) => HEADER_BY_KEY.get(k) ?? k).join(', ')}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </details>
      )}

      {panelNotFound.length > 0 && (
        <details>
          <summary>
            {panelNotFound.length} products showing &ldquo;Not Found&rdquo; for Panel Size
          </summary>
          <div className="note warn">
            Inherited from the old <code className="inline">Auto-NX</code> sheet, which could not
            read a panel size written as text. New products are no longer affected; existing ones
            can be corrected by typing the size in step 4.
          </div>
        </details>
      )}

      {dupNames.length > 0 && (
        <details>
          <summary>{dupNames.length} Product Names used by more than one product</summary>
          <div className="note warn">
            The site matches with <code className="inline">.find()</code>, so only the first one is
            reachable by model name. Each of them can still be found by its own 90PN.
          </div>
          <div className="scroll">
            <table>
              <thead>
                <tr>
                  <th>Product Name</th>
                  <th className="mono">Count</th>
                  <th className="mono">90PN</th>
                </tr>
              </thead>
              <tbody>
                {dupNames.map(([name, count]) => (
                  <tr key={name}>
                    <td>{name}</td>
                    <td className="mono">{count}</td>
                    <td className="mono">
                      {rows
                        .filter((r) => r.productName === name)
                        .map((r) => r.partNumber)
                        .join(', ')}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </details>
      )}

      <div className="note ok" style={{ marginTop: 14 }}>
        {noLink.length} of {total} products have no materials Link yet — the Materials button will
        show &ldquo;Coming Soon&rdquo; for those.
      </div>

      {merge && merge.warnings.length > 0 && (
        <details>
          <summary>{merge.warnings.length} note(s) while merging the price list</summary>
          <div className="note warn">
            <ul>
              {merge.warnings.slice(0, 30).map((w, i) => (
                <li key={i}>{w}</li>
              ))}
            </ul>
          </div>
        </details>
      )}
    </>
  );
}
