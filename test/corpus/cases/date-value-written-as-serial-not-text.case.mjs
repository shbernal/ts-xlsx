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

// A time-of-day under a duration/time number format: the fraction-of-a-day serial is what makes a
// SUM over a column of durations meaningful. Stored as text (the reported failure mode) the total
// is zero, so this must be a numeric cell, not an inline string.
const TIME_ISO = '1899-12-30T10:51:00.000Z'; // Excel epoch + 10:51, i.e. serial ≈ 0.4521
const TIME_SPEC = {
  sheets: [{name: 'T', cells: [{ref: 'A1', value: {date: TIME_ISO}, numFmt: '[h]:mm'}]}],
};

// The dual of the above: a date-LOOKING STRING under a date number format must stay a string. A date
// format is a display instruction, not a coercion — applying it cannot turn text into a date. So a
// caller who wants real, pivotable dates must supply Date values; a string like "2024/02/02" under a
// date column stays text (and pivots/sorts as text), which the caller can detect rather than being
// silently told it is a date.
const STRING_UNDER_DATE_FMT = {
  sheets: [
    {
      name: 'S',
      columns: [{index: 1, numFmt: 'yyyy/mm/dd'}],
      cells: [{ref: 'A1', value: '2024/02/02'}],
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
        assert.strictEqual(
          model.sheets.S.cells.A1.value,
          ISO,
          'the value round-trips as a date, not the text "2020-01-15"',
        );
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
    {
      // A time-of-day under a time/duration format must stay a numeric (fractional-serial) cell so
      // arithmetic over a column of durations works — storing it as text makes a SUM evaluate to 0.
      name: 'a time-of-day under a duration format stays a numeric date value, not text',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(TIME_SPEC);
        const a1 = model.sheets.T.cells.A1;
        assert.strictEqual(
          a1.value,
          TIME_ISO,
          'the time value round-trips as a date, not the text "10:51"',
        );
        assert.strictEqual(a1.numFmt, '[h]:mm', 'the duration number format is preserved');
      },
    },
    {
      name: 'a date-looking string under a date column stays a string (the format does not coerce it)',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(STRING_UNDER_DATE_FMT);
        assert.strictEqual(
          model.sheets.S.cells.A1.value,
          '2024/02/02',
          'the string keeps its exact text — a date number format must not turn text into a date',
        );
      },
    },
  ],
};
