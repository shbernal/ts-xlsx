import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

// Author a workbook whose single sheet carries one table, round-trip it, and hand back the
// reconstructed table for assertions.
function roundtripTable(options: {
  name: string;
  displayName?: string;
  ref: string;
  columns: {name: string; totalsRowLabel?: string; totalsRowFunction?: string}[];
  rowCount: number;
  headerRow?: boolean;
  totalsRow?: boolean;
  totalsRowShown?: boolean;
  autoFilter?: boolean;
  style?: {
    name?: string;
    showFirstColumn?: boolean;
    showLastColumn?: boolean;
    showRowStripes?: boolean;
    showColumnStripes?: boolean;
  };
}) {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable(options);
  const back = readXlsx(writeXlsx(wb));
  const sheet = back.getWorksheet('S');
  assert.ok(sheet !== undefined);
  return sheet.tables;
}

test('a table read back from a written package exposes its name, columns, and ref', () => {
  const [table, ...rest] = roundtripTable({
    name: 'Inventory',
    ref: 'A1',
    columns: [{name: 'Item'}, {name: 'Qty'}],
    rowCount: 3,
  });
  assert.equal(rest.length, 0, 'exactly one table is reconstructed');
  assert.ok(table !== undefined);
  assert.equal(table.name, 'Inventory');
  assert.deepEqual(
    table.columns.map((c) => c.name),
    ['Item', 'Qty'],
  );
  assert.equal(table.ref, 'A1:B4', 'header + 3 data rows spans four rows');
  assert.equal(table.options.ref, 'A1', 'the anchor reconstructs to the top-left cell');
});

test('a loaded table exposes its data-row count, not an empty rows array', () => {
  const [table] = roundtripTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Name'}],
    rowCount: 5,
  });
  assert.ok(table !== undefined);
  assert.equal(table.options.rowCount, 5, 'the data-row count survives the round-trip');
});

test('an empty-body table (header only) round-trips with a zero data-row count', () => {
  const [table] = roundtripTable({
    name: 'Empty',
    ref: 'A1',
    columns: [{name: 'C1'}, {name: 'C2'}],
    rowCount: 0,
  });
  assert.ok(table !== undefined);
  assert.equal(table.options.rowCount, 0);
  assert.equal(table.ref, 'A1:B1', 'a header-only table occupies a single row');
});

test('a headerless table round-trips with headerRow false and no autofilter', () => {
  const [table] = roundtripTable({
    name: 'Bare',
    ref: 'A1',
    columns: [{name: 'C1'}],
    rowCount: 2,
    headerRow: false,
  });
  assert.ok(table !== undefined);
  assert.equal(table.options.headerRow, false);
  assert.equal(table.options.rowCount, 2, 'both rows are data rows when there is no header');
  assert.equal(table.autoFilterRef, null, 'a headerless table anchors no autofilter');
});

test('a totals-row table round-trips its totals flag and per-column totals behaviour', () => {
  const [table] = roundtripTable({
    name: 'Totalled',
    ref: 'A1',
    columns: [
      {name: 'Label', totalsRowLabel: 'Total'},
      {name: 'Amount', totalsRowFunction: 'sum'},
    ],
    rowCount: 2,
    totalsRow: true,
  });
  assert.ok(table !== undefined);
  assert.equal(table.options.totalsRow, true);
  assert.equal(table.options.rowCount, 2, 'the totals row is not counted as a data row');
  assert.equal(table.ref, 'A1:B4', 'header + 2 data + totals spans four rows');
  assert.equal(table.columns[0]?.totalsRowLabel, 'Total');
  assert.equal(table.columns[1]?.totalsRowFunction, 'sum');
});

test('a table with no totalsRowShown flag stays without one across a round-trip', () => {
  const [table] = roundtripTable({name: 'T', ref: 'A1', columns: [{name: 'A'}], rowCount: 2});
  assert.ok(table !== undefined);
  assert.equal(
    table.options.totalsRowShown,
    undefined,
    'an absent flag must not be fabricated on read-back',
  );
  assert.equal(
    'totalsRowShown' in table.options,
    false,
    'the key is omitted, not set to undefined',
  );
});

test('an explicit totalsRowShown flag survives a round-trip in both states', () => {
  const [off] = roundtripTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 2,
    totalsRowShown: false,
  });
  assert.equal(off?.options.totalsRowShown, false);

  const [on] = roundtripTable({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 2,
    totalsRowShown: true,
  });
  assert.equal(on?.options.totalsRowShown, true);
});

test('a custom table style name and banding flags survive a round-trip verbatim', () => {
  const [table] = roundtripTable({
    name: 'Styled',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 2,
    style: {
      name: 'TableStyleLight9',
      showFirstColumn: true,
      showLastColumn: false,
      showRowStripes: false,
      showColumnStripes: true,
    },
  });
  assert.ok(table !== undefined);
  assert.deepEqual(table.options.style, {
    name: 'TableStyleLight9',
    showFirstColumn: true,
    showLastColumn: false,
    showRowStripes: false,
    showColumnStripes: true,
  });
});

test('a style that omits its name round-trips still nameless, not defaulted', () => {
  const [table] = roundtripTable({
    name: 'NoName',
    ref: 'A1',
    columns: [{name: 'A'}],
    rowCount: 1,
    style: {showRowStripes: true},
  });
  assert.ok(table !== undefined);
  assert.equal(
    'name' in (table.options.style ?? {}),
    false,
    'an absent style name is not fabricated',
  );
  assert.equal(table.options.style?.showRowStripes, true);
});

test("a table authored without a style reads back with Excel's default style", () => {
  const [table] = roundtripTable({name: 'Plain', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  assert.ok(table !== undefined);
  assert.equal(table.options.style?.name, 'TableStyleMedium2');
  assert.equal(table.options.style?.showRowStripes, true);
});

test('a distinct display name round-trips independently of the internal name', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'MyTable',
    displayName: 'My Display Name',
    ref: 'A1',
    columns: [{name: 'C'}],
    rowCount: 1,
  });
  const table = readXlsx(writeXlsx(wb)).getWorksheet('S')?.tables[0];
  assert.ok(table !== undefined);
  assert.equal(table.name, 'MyTable', 'the internal identifier is unaffected');
  assert.equal(table.displayName, 'My Display Name', 'the display label survives the round-trip');
});

test('a display name defaults to the table name when unset', () => {
  const [table] = roundtripTable({name: 'Plain', ref: 'A1', columns: [{name: 'C'}], rowCount: 1});
  assert.ok(table !== undefined);
  assert.equal(table.displayName, 'Plain');
});

test('a header table read without an autoFilter does not gain one on round-trip', () => {
  const [table] = roundtripTable({
    name: 'Unfiltered',
    ref: 'A1',
    columns: [{name: 'C'}],
    rowCount: 2,
    autoFilter: false,
  });
  assert.ok(table !== undefined);
  assert.equal(table.headerRow, true, 'it keeps its header row');
  assert.equal(table.autoFilterRef, null, 'but exposes no autoFilter range');
  assert.equal(table.options.autoFilter, false, 'and the flag survives the round-trip');
});

test('a header table gains an autoFilter by default', () => {
  const [table] = roundtripTable({
    name: 'Filtered',
    ref: 'A1',
    columns: [{name: 'C'}],
    rowCount: 2,
  });
  assert.ok(table !== undefined);
  assert.equal(table.autoFilterRef, 'A1:A3', 'the default autoFilter spans header + data rows');
  assert.equal(table.options.autoFilter, true);
});

test('several tables on one sheet all read back in definition order', () => {
  const wb = new Workbook();
  const sheet = wb.addWorksheet('S');
  sheet.addTable({name: 'First', ref: 'A1', columns: [{name: 'A'}], rowCount: 1});
  sheet.addTable({name: 'Second', ref: 'D1', columns: [{name: 'B'}, {name: 'C'}], rowCount: 2});

  const back = readXlsx(writeXlsx(wb));
  const tables = back.getWorksheet('S')?.tables ?? [];
  assert.deepEqual(
    tables.map((t) => t.name),
    ['First', 'Second'],
  );
  assert.equal(tables[1]?.ref, 'D1:E3');
});

test('tables on distinct sheets each reconstruct on their own sheet', () => {
  const wb = new Workbook();
  wb.addWorksheet('One').addTable({name: 'TA', ref: 'A1', columns: [{name: 'X'}], rowCount: 1});
  wb.addWorksheet('Two').addTable({name: 'TB', ref: 'A1', columns: [{name: 'Y'}], rowCount: 1});

  const back = readXlsx(writeXlsx(wb));
  assert.deepEqual(
    back.getWorksheet('One')?.tables.map((t) => t.name),
    ['TA'],
  );
  assert.deepEqual(
    back.getWorksheet('Two')?.tables.map((t) => t.name),
    ['TB'],
  );
});

test('a table survives a second read → write → read round-trip unchanged', () => {
  const wb = new Workbook();
  wb.addWorksheet('S').addTable({
    name: 'Persist',
    ref: 'B2',
    columns: [{name: 'One'}, {name: 'Two'}, {name: 'Three'}],
    rowCount: 4,
  });
  const once = readXlsx(writeXlsx(wb));
  const twice = readXlsx(writeXlsx(once));
  const table = twice.getWorksheet('S')?.tables[0];
  assert.ok(table !== undefined);
  assert.equal(table.name, 'Persist');
  assert.equal(table.ref, 'B2:D6', 'the range is stable across two round-trips');
  assert.deepEqual(
    table.columns.map((c) => c.name),
    ['One', 'Two', 'Three'],
  );
});
