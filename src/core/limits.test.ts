import assert from 'node:assert/strict';
import {test} from 'node:test';

import {Column} from './column.ts';
import {MAX_COLUMN_WIDTH, MAX_ROW_HEIGHT} from './limits.ts';
import {Row} from './row.ts';
import {Worksheet} from './worksheet.ts';

// These numbers are Excel Desktop's, measured over COM against a blank workbook — see
// docs/knowledge/specs/grid-geometry-limits-are-excels-not-the-schemas.md. The point of pinning
// them is that they are *not* what Microsoft's published specifications table says (409 points),
// so a future reader who "corrects" them against that page fails here and finds the probe.
test('the geometry limits are the values Excel accepts, not the documented ones', () => {
  assert.equal(MAX_ROW_HEIGHT, 409.5, 'Excel takes 409.5 and refuses 409.6');
  assert.equal(MAX_COLUMN_WIDTH, 255, 'Excel takes 255 and refuses 255.4');
});

test('the model does not enforce them — a foreign file must survive being read', () => {
  const sheet = new Worksheet('S', 1);
  const row = new Row(sheet, 1);
  row.height = MAX_ROW_HEIGHT * 2;
  assert.equal(row.height, MAX_ROW_HEIGHT * 2, 'held as stated, not clamped or refused');

  const column = new Column(sheet, 1);
  column.width = MAX_COLUMN_WIDTH * 2;
  assert.equal(column.width, MAX_COLUMN_WIDTH * 2);
});
