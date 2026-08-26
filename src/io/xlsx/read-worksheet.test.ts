import assert from 'node:assert/strict';
import {test} from 'node:test';

import {MAX_COLUMN, MAX_ROW} from '../../core/address.ts';
import {Workbook} from '../../core/workbook.ts';
import {parseWorksheet} from './read-worksheet.ts';
import {sheetViewsXml} from './sheet-properties.ts';

// `customWidth`/`customHeight` are xsd:booleans, so a foreign producer may spell false either way.
// Excel writes the digit, which is why the long spelling went unnoticed for so long.
function sheet(body: string) {
  const worksheet = new Workbook().addWorksheet('S');
  parseWorksheet(
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      body +
      '</worksheet>',
    worksheet,
    [],
    [],
  );
  return worksheet;
}

test('customWidth="false" suppresses the width exactly as customWidth="0" does', () => {
  for (const spelling of ['0', 'false']) {
    const width = sheet(
      `<cols><col min="2" max="2" width="12" customWidth="${spelling}"/></cols>`,
    ).getColumn(2).width;
    assert.equal(width, undefined, `customWidth="${spelling}" is not a custom width`);
  }
  assert.equal(
    sheet('<cols><col min="2" max="2" width="12" customWidth="1"/></cols>').getColumn(2).width,
    12,
    'and the flag set still carries the width through',
  );
});

test('customHeight="false" suppresses the row height exactly as customHeight="0" does', () => {
  for (const spelling of ['0', 'false']) {
    const height = sheet(
      `<sheetData><row r="3" ht="30" customHeight="${spelling}"/></sheetData>`,
    ).getRow(3).height;
    assert.equal(height, undefined, `customHeight="${spelling}" is not a custom height`);
  }
  assert.equal(
    sheet('<sheetData><row r="3" ht="30" customHeight="1"/></sheetData>').getRow(3).height,
    30,
    'and the flag set still carries the height through',
  );
});

test('a `<col>` span past XFD is clamped, not walked', () => {
  // Unclamped, `max="99999999"` materialised 16.7 million column records before dying on the map's
  // size limit. Assert the record count rather than the elapsed time: the count is what bounds it.
  const worksheet = sheet('<cols><col min="1" max="99999999" width="12" customWidth="1"/></cols>');
  assert.equal([...worksheet.columns()].length, MAX_COLUMN);
  assert.equal(worksheet.getColumn(MAX_COLUMN).width, 12, 'the span reaches the last real column');
});

test('a `<col>` element wholly outside the grid is dropped', () => {
  const worksheet = sheet(
    `<cols><col min="${MAX_COLUMN + 1}" max="99999" width="12" customWidth="1"/></cols>`,
  );
  assert.equal([...worksheet.columns()].length, 0);
});

test('a `<row>` past the last row is dropped rather than clamped onto it', () => {
  const worksheet = sheet(
    `<sheetData><row r="${MAX_ROW + 1}" ht="30" customHeight="1"/></sheetData>`,
  );
  assert.equal([...worksheet.rows()].length, 0);
  assert.equal(worksheet.getRow(MAX_ROW).height, undefined, 'and nothing landed on the last row');
});

test('a `<pane>` split that is not a non-negative integer is dropped, not carried to the writer', () => {
  // Each of these used to land in `view` verbatim and surface later out of the serializer, as a
  // RangeError naming a column the file never mentioned.
  for (const attrs of ['xSplit="abc"', 'xSplit="1.5" ySplit="2"', 'ySplit="-3"']) {
    const worksheet = sheet(
      `<sheetViews><sheetView workbookViewId="0"><pane ${attrs} topLeftCell="B2" state="frozen"/></sheetView></sheetViews>`,
    );
    assert.equal(worksheet.view.state, 'frozen', `${attrs}: the pane itself still reads`);
    for (const axis of ['xSplit', 'ySplit'] as const) {
      const split = worksheet.view[axis];
      assert.ok(
        split === undefined || (Number.isInteger(split) && split >= 0),
        `${attrs}: ${axis} is ${String(split)}, which freeze() would refuse`,
      );
    }
    assert.doesNotThrow(() => sheetViewsXml(worksheet.view, true), `${attrs}: and re-writes clean`);
  }
});

test('a well-formed `<pane>` split still reads', () => {
  const worksheet = sheet(
    '<sheetViews><sheetView workbookViewId="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" state="frozen"/></sheetView></sheetViews>',
  );
  assert.equal(worksheet.view.xSplit, 2);
  assert.equal(worksheet.view.ySplit, 1);
  assert.equal(worksheet.view.topLeftCell, 'C2');
});
