// Field-agnostic cover for the model round-trip. The per-field cases live in worksheet.test.ts;
// these assert the properties the facet registry is what guarantees, so they keep their meaning —
// and keep catching a half-wired field — as WorksheetModel grows.

import assert from 'node:assert/strict';
import {test} from 'node:test';
import {Worksheet} from './worksheet.ts';
import {WORKSHEET_MODEL_FACETS} from './worksheet-model.ts';

// Every field of a WorksheetModel, populated, so a round-trip has something to lose in each of them.
function populatedSheet(): Worksheet {
  const sheet = new Worksheet('Src', 1);
  sheet.state = 'hidden';
  sheet.tabColor = {argb: 'FFFF0000'};
  sheet.properties.defaultRowHeight = 18;
  sheet.outline.summaryBelow = false;
  sheet.pageSetup.fitToPage = true;
  sheet.printOptions.gridLines = true;
  sheet.pageMargins.left = 0.25;
  sheet.headerFooter.oddHeader = '&Ctitle';
  sheet.rowBreaks.push({id: 3});
  sheet.columnBreaks.push({id: 2});
  sheet.getColumn(2).width = 20;
  sheet.getRow(7).height = 33; // a row that carries formatting and no cells
  sheet.getCell('A1').value = 'title';
  sheet.getCell('B2').value = 42;
  sheet.getCell('B2').font = {bold: true};
  sheet.mergeCells('D1:E1');
  sheet.addDataValidation('F1', {type: 'whole', operator: 'between', formulae: [0, 9]});
  sheet.addConditionalFormatting({
    ref: 'G1:G3',
    rules: [{type: 'cellIs', operator: 'greaterThan', formulae: [5], priority: 1}],
  });
  sheet.addTable({name: 'T1', ref: 'A10', columns: [{name: 'Col'}], rowCount: 2});
  sheet.autoFilter = 'A1:B2';
  sheet.protect('secret');
  return sheet;
}

test('the exported model carries exactly the fields the facet registry declares', () => {
  // The registry is proved exhaustive over `keyof WorksheetModel` at compile time; this is the
  // runtime half — that the getter actually emits a key for each facet and invents none besides.
  assert.deepEqual(
    Object.keys(populatedSheet().model).sort(),
    WORKSHEET_MODEL_FACETS.map((facet) => facet.key).sort(),
  );
});

test('a fully populated sheet round-trips to an identical model', () => {
  const src = populatedSheet();
  const dst = new Worksheet('Dst', 2);
  dst.model = src.model;

  // A facet whose two directions disagree — a field read but not written, or written into the wrong
  // place — shows up here as a difference, whichever field it is.
  assert.deepEqual(dst.model, src.model);
});

test('a round-trip through a sheet holding unrelated content leaves no residue', () => {
  const dst = new Worksheet('Dst', 2);
  dst.getCell('Z9').value = 'stale';
  dst.getColumn(26).width = 99;
  dst.getRow(9).height = 99;
  dst.mergeCells('X1:Y1');
  dst.addDataValidation('W1', {type: 'whole', formulae: [1]});
  dst.protect('other');

  const src = populatedSheet();
  dst.model = src.model;

  assert.deepEqual(dst.model, src.model);
});

test('the exported model orders cells, rows, and columns ascending whatever order they were written', () => {
  const sheet = new Worksheet('S', 1);
  for (const reference of ['C3', 'A1', 'B2', 'A3']) sheet.getCell(reference).value = reference;
  for (const index of [5, 1, 3]) sheet.getColumn(index).width = index;
  for (const number of [6, 2, 4]) sheet.getRow(number).height = number;

  const model = sheet.model;
  assert.deepEqual(
    model.cells.map((cell) => [cell.row, cell.col]),
    [
      [1, 1],
      [2, 2],
      [3, 1],
      [3, 3],
    ],
    'cells come out row-major, ascending within each row',
  );
  assert.deepEqual(
    model.columns.map((column) => column.index),
    [1, 3, 5],
  );
  assert.deepEqual(
    model.rows.map((row) => row.number),
    [2, 4, 6],
  );
});

test('a model exported from an empty sheet clears every field of the sheet it is assigned to', () => {
  const dst = populatedSheet();
  const empty = new Worksheet('Empty', 3).model;

  dst.model = empty;

  assert.deepEqual(dst.model, empty);
});
