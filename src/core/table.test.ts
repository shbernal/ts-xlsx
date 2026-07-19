import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Table, type TableOptions} from './table.ts';

function table(overrides: Partial<TableOptions> = {}): Table {
  return new Table({
    name: 'T',
    ref: 'A3',
    columns: [{name: 'H1'}, {name: 'H2'}],
    rowCount: 2,
    ...overrides,
  });
}

test('duplicate column names are disambiguated with a numeric suffix', () => {
  const t = new Table({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Dup'}, {name: 'Dup'}, {name: 'Dup'}],
    rowCount: 1,
  });
  assert.deepStrictEqual(
    t.columns.map((c) => c.name),
    ['Dup', 'Dup2', 'Dup3'],
    'the first name is kept; later clashes gain the smallest resolving suffix',
  );
});

test('duplicate column names are disambiguated case-insensitively', () => {
  const t = new Table({
    name: 'T',
    ref: 'A1',
    columns: [{name: 'Name'}, {name: 'name'}],
    rowCount: 1,
  });
  const [first, second] = t.columns.map((c) => c.name);
  assert.strictEqual(first, 'Name');
  assert.notStrictEqual(
    second?.toLowerCase(),
    'name',
    'a case-insensitive clash is still resolved',
  );
});

test('distinct column names are accepted', () => {
  assert.doesNotThrow(() => table());
});

test('inserting a row above a table shifts its whole range down', () => {
  const t = table(); // A3:B5 (header + 2 data rows)
  assert.strictEqual(t.ref, 'A3:B5');
  const alive = t.shiftRows(1, 0, 1); // insert one row at the top
  assert.strictEqual(alive, true);
  assert.strictEqual(t.ref, 'A4:B6');
});

test('inserting rows inside a table grows its data rows', () => {
  const t = table(); // A3:B5
  t.shiftRows(4, 0, 2); // two rows inserted within the data body (row 4)
  assert.strictEqual(t.ref, 'A3:B7', 'the table absorbs the inserted rows');
});

test('a table left entirely above the splice is untouched', () => {
  const t = table(); // A3:B5
  t.shiftRows(10, 0, 5); // insert well below the table
  assert.strictEqual(t.ref, 'A3:B5');
});

test('deleting every row of a table reports it as removed', () => {
  const t = table(); // rows 3..5
  const alive = t.shiftRows(3, 3, -3); // delete the whole span
  assert.strictEqual(alive, false);
});

test('a column splice to the left shifts the table anchor', () => {
  const t = table(); // anchored at column A (A3:B5)
  t.shiftColumns(1, 0, 2); // insert two columns before it
  assert.strictEqual(t.ref, 'C3:D5');
});

test('the sentinel style name "None" is normalised to an absent name', () => {
  const t = table({style: {name: 'None', showRowStripes: true}});
  assert.strictEqual(
    t.style?.name,
    undefined,
    '"None" means unstyled, not a literal style reference',
  );
  assert.strictEqual(t.style?.showRowStripes, true, 'flags set alongside the theme survive');
});

test('a real style name is preserved', () => {
  const t = table({style: {name: 'TableStyleMedium2'}});
  assert.strictEqual(t.style?.name, 'TableStyleMedium2');
});

test('rowCount reports the data-row count', () => {
  assert.strictEqual(table().rowCount, 2);
});

test('addRow grows the data-row count and the range', () => {
  const t = table(); // A3:B5, 2 data rows
  t.addRow();
  assert.strictEqual(t.rowCount, 3);
  assert.strictEqual(t.ref, 'A3:B6');
});

test('addRow with values on a table not attached to a worksheet throws', () => {
  assert.throws(() => table().addRow(['a', 1]), /not attached to a worksheet/);
});

test('addRow rejects more values than the table has columns', () => {
  assert.throws(() => table().addRow(['a', 1, 'x']), /has 2 columns/);
});

test('addRow on a detached table with a totals row throws — relocation needs the grid', () => {
  assert.throws(
    () => table({totalsRow: true}).addRow(),
    /not attached to a worksheet.*relocate its totals row/,
  );
});
