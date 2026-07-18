// Worksheet tables (OOXML `<table>` parts) — the reader that reconstructs a table's model from its
// `xl/tables/table{n}.xml` part. The writer (write.ts) is the serialization gatekeeper; this module
// is its inverse, turning a stored table part back into the `TableOptions` a worksheet re-registers.
//
// The part stores the table's *full* occupied range (`ref="A1:B3"`), whereas the model anchors at a
// single top-left cell plus a data-row count. The two are equivalent: the data-row count is the
// range height minus the header row (present unless `headerRowCount="0"`) and the totals row (present
// only when `totalsRowCount` is positive), so reconstructing one from the other is lossless.

import {decodeRange, encodeAddress} from '../../core/address.ts';
import type {TableColumn, TableOptions, TableStyleInfo} from '../../core/table.ts';
import {localName, parseXml} from './xml-read.ts';

// OOXML booleans spell false as "0" or "false"; every other spelling (including "1"/"true") is true.
function parseOoxmlBool(value: string): boolean {
  return value !== '0' && value !== 'false';
}

/**
 * Parse a `<table>` part into the options that reconstruct it, or `undefined` when the XML is not a
 * usable table (no name, no ref, or no columns — Excel treats such a part as corrupt, so we drop it
 * rather than fabricate a degenerate table). Duplicate column names are not resolved here — the
 * {@link Table} constructor disambiguates them, so authoring and loading share one implementation.
 */
export function parseTable(xml: string): TableOptions | undefined {
  let name: string | undefined;
  let displayName: string | undefined;
  let ref: string | undefined;
  let headerRowCount = 1; // OOXML default: a table carries a header row unless it says otherwise.
  let totalsRowCount = 0; // OOXML default: no totals row.
  let totalsRowShown: boolean | undefined; // Absent unless the part states the attribute.
  let style: TableStyleInfo | undefined; // Absent unless the part carries a `<tableStyleInfo>`.
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
          // Capture the flag verbatim so it re-emits exactly (or, absent, stays absent) rather
          // than being normalised.
          if (attrs.totalsRowShown !== undefined) totalsRowShown = parseOoxmlBool(attrs.totalsRowShown);
          break;
        case 'autoFilter':
          hasAutoFilter = true;
          break;
        case 'tableStyleInfo': {
          // Keep each attribute off the literal so an absent one stays absent (not `key: undefined`),
          // preserving the round-trip — the writer re-emits only the attributes we actually saw.
          const captured: {-readonly [K in keyof TableStyleInfo]: TableStyleInfo[K]} = {};
          if (attrs.name !== undefined) captured.name = attrs.name;
          if (attrs.showFirstColumn !== undefined) captured.showFirstColumn = parseOoxmlBool(attrs.showFirstColumn);
          if (attrs.showLastColumn !== undefined) captured.showLastColumn = parseOoxmlBool(attrs.showLastColumn);
          if (attrs.showRowStripes !== undefined) captured.showRowStripes = parseOoxmlBool(attrs.showRowStripes);
          if (attrs.showColumnStripes !== undefined) {
            captured.showColumnStripes = parseOoxmlBool(attrs.showColumnStripes);
          }
          style = captured;
          break;
        }
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

  const options: TableOptions = {
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
  // Kept off the literal so an absent attribute stays absent (not `totalsRowShown: undefined`),
  // preserving the round-trip: a table that never stated the flag must not gain one.
  if (totalsRowShown !== undefined) options.totalsRowShown = totalsRowShown;
  if (style !== undefined) options.style = style;
  return options;
}
