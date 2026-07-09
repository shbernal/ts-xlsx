// Cluster: types
//
// Real-world scenario: a user assigns a genuine Date to a cell that carries a date number
// format — set either per-cell or via a column-level style. The cell must be written as a
// numeric-typed cell holding the date serial number (so a spreadsheet app treats it as a real
// date and the number format displays it), never as a text string. A date silently coerced to
// text sorts and computes wrong and cannot be reformatted. Applying a date number format at the
// column level must not, by itself, turn a date value into a text cell.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const ISO = '2020-01-15T00:00:00.000Z';
const SPEC = {
  sheets: [
    {
      name: 'S',
      columns: [{index: 1, numFmt: 'yyyy-mm-dd'}],
      cells: [
        {ref: 'A1', value: {date: ISO}}, // date under a column-level date format
        {ref: 'B1', value: {date: ISO}, numFmt: 'yyyy-mm-dd'}, // date under a per-cell format
      ],
    },
  ],
};

export default {
  id: 'date-value-written-as-serial-not-text',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1666},
  cluster: 'types',
  description:
    'A genuine date value under a date number format is written as a numeric (serial) cell ' +
    'and reads back as a date, not a text string — for both a column-level and a per-cell ' +
    'date format; a column date format alone does not coerce the value to text.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a date under a column-level date format reads back as a date',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(model.sheets.S.cells.A1.value, ISO, 'the value round-trips as a date, not the text "2020-01-15"');
      },
    },
    {
      name: 'a date under a per-cell date format reads back as a date and keeps its format',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const b1 = model.sheets.S.cells.B1;
        assert.strictEqual(b1.value, ISO, 'the per-cell date value round-trips as a date');
        assert.strictEqual(b1.numFmt, 'yyyy-mm-dd', 'the date number format is preserved');
      },
    },
  ],
};
