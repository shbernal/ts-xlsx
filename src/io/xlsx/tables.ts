// Worksheet tables (OOXML `<table>` parts) — the reader that reconstructs a table's model from its
// `xl/tables/table{n}.xml` part. The writer (write.ts) is the serialization gatekeeper; this module
// is its inverse, turning a stored table part back into the `TableOptions` a worksheet re-registers.
//
// The part stores the table's *full* occupied range (`ref="A1:B3"`), whereas the model anchors at a
// single top-left cell plus a data-row count. The two are equivalent: the data-row count is the
// range height minus the header row (present unless `headerRowCount="0"`) and the totals row (present
// only when `totalsRowCount` is positive), so reconstructing one from the other is lossless.

import {decodeRange, encodeAddress} from '../../core/address.ts';
import type {TableColumn, TableOptions} from '../../core/table.ts';
import {localName, parseXml} from './xml-read.ts';

/**
 * Parse a `<table>` part into the options that reconstruct it, or `undefined` when the XML is not a
 * usable table (no name, no ref, or no columns — Excel treats such a part as corrupt, so we drop it
 * rather than fabricate a degenerate table).
 */
/**
 * Rename any column whose name collides (case-insensitively) with an earlier one, in place, by
 * appending the smallest numeric suffix that makes it unique — the same disambiguation Excel applies
 * when loading a table with duplicate column names.
 */
function disambiguateColumnNames(columns: TableColumn[]): void {
  const seen = new Set<string>();
  for (let i = 0; i < columns.length; i++) {
    const column = columns[i];
    if (column === undefined) continue;
    let candidate = column.name;
    for (let n = 2; seen.has(candidate.toLowerCase()); n++) candidate = `${column.name}${n}`;
    seen.add(candidate.toLowerCase());
    if (candidate !== column.name) columns[i] = {...column, name: candidate};
  }
}

export function parseTable(xml: string): TableOptions | undefined {
  let name: string | undefined;
  let displayName: string | undefined;
  let ref: string | undefined;
  let headerRowCount = 1; // OOXML default: a table carries a header row unless it says otherwise.
  let totalsRowCount = 0; // OOXML default: no totals row.
  let hasAutoFilter = false; // Only present when the part carries an `<autoFilter>` element.
  const columns: TableColumn[] = [];

  parseXml(xml, {
    onOpen(elementName, attrs) {
      switch (localName(elementName)) {
        case 'table':
          // OOXML makes `displayName` the required identifier and `name` an optional alias; the
          // model inverts the roles (`name` is the formula identifier, `displayName` the label),
          // so read each from its own attribute and fall back across the pair when one is absent.
          name = attrs.name ?? attrs.displayName;
          displayName = attrs.displayName ?? attrs.name;
          ref = attrs.ref;
          if (attrs.headerRowCount !== undefined) headerRowCount = Number(attrs.headerRowCount);
          if (attrs.totalsRowCount !== undefined) totalsRowCount = Number(attrs.totalsRowCount);
          break;
        case 'autoFilter':
          hasAutoFilter = true;
          break;
        case 'tableColumn': {
          if (attrs.name === undefined) break;
          const column: {name: string; totalsRowLabel?: string; totalsRowFunction?: string} = {
            name: attrs.name,
          };
          if (attrs.totalsRowLabel !== undefined) column.totalsRowLabel = attrs.totalsRowLabel;
          if (attrs.totalsRowFunction !== undefined) column.totalsRowFunction = attrs.totalsRowFunction;
          columns.push(column);
          break;
        }
      }
    },
    onText() {},
    onClose() {},
  });

  if (name === undefined || ref === undefined || columns.length === 0) return undefined;

  // Real workbooks exist with duplicate table-column names (Excel writes the file, then disambiguates
  // on load by suffixing the clashes). The model's authoring guard rejects duplicates outright, so the
  // reader must resolve them here — matching Excel's suffixing — rather than reject a file Excel opens.
  disambiguateColumnNames(columns);

  const {top, left, bottom} = decodeRange(ref);
  if (top === undefined || left === undefined || bottom === undefined) return undefined;

  const headerRow = headerRowCount !== 0;
  const totalsRow = totalsRowCount > 0;
  const dataRows = bottom - top + 1 - (headerRow ? 1 : 0) - (totalsRow ? 1 : 0);

  return {
    name,
    displayName: displayName ?? name,
    ref: encodeAddress(left, top),
    columns,
    rowCount: Math.max(0, dataRows),
    headerRow,
    totalsRow,
    // Reconstruct the autoFilter state explicitly from the part: a header table read without an
    // `<autoFilter>` must not have one fabricated on the next write.
    autoFilter: hasAutoFilter,
  };
}
