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
export function parseTable(xml: string): TableOptions | undefined {
  let name: string | undefined;
  let ref: string | undefined;
  let headerRowCount = 1; // OOXML default: a table carries a header row unless it says otherwise.
  let totalsRowCount = 0; // OOXML default: no totals row.
  const columns: TableColumn[] = [];

  parseXml(xml, {
    onOpen(elementName, attrs) {
      switch (localName(elementName)) {
        case 'table':
          name = attrs.name;
          ref = attrs.ref;
          if (attrs.headerRowCount !== undefined) headerRowCount = Number(attrs.headerRowCount);
          if (attrs.totalsRowCount !== undefined) totalsRowCount = Number(attrs.totalsRowCount);
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

  const {top, left, bottom} = decodeRange(ref);
  if (top === undefined || left === undefined || bottom === undefined) return undefined;

  const headerRow = headerRowCount !== 0;
  const totalsRow = totalsRowCount > 0;
  const dataRows = bottom - top + 1 - (headerRow ? 1 : 0) - (totalsRow ? 1 : 0);

  return {
    name,
    ref: encodeAddress(left, top),
    columns,
    rowCount: Math.max(0, dataRows),
    headerRow,
    totalsRow,
  };
}
