import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Workbook} from '../../core/workbook.ts';
import {parseWorksheet} from './read-worksheet.ts';

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
