// Cluster: types
//
// Real-world scenario: a cell carries a date-style number format but its stored value is not a valid
// date — most sharply, an already-invalid Date object (the result of parsing bad input). Because the
// number format is a date format, the writer runs the date→serial conversion, and an invalid Date
// converts to NaN. Emitting <v>NaN</v> as a numeric cell value produces malformed content Excel
// flags as corrupt and offers to repair on reopen. The writer must degrade gracefully — never leak
// NaN or "Invalid Date" into the sheet XML. A plain string or null under the same date format is the
// working control (no conversion is forced, nothing bogus is emitted).

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'date-numfmt-nonnumeric-value-serializes-valid-xml',
  provenance: {source: 'upstream-pr'},
  cluster: 'types',
  description:
    'A cell with a date number format over a non-numeric value must never serialize NaN or ' +
    '"Invalid Date" into the sheet XML — an invalid Date under a date numFmt degrades gracefully ' +
    'rather than emitting a bogus numeric serial Excel rejects; a string or null under the same ' +
    'format is unaffected.',

  behavior: [
    {
      name: 'a string under a date number format serializes no NaN (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, hasNaN, hasInvalidDate} = await api.dateNumFmtValueReport('string');
        assert.ok(ok, 'the workbook writes');
        assert.strictEqual(hasNaN, false, 'a string value emits no NaN');
        assert.strictEqual(hasInvalidDate, false, 'a string value emits no "Invalid Date"');
      },
    },
    {
      name: 'a null under a date number format serializes no NaN (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasNaN, hasInvalidDate} = await api.dateNumFmtValueReport('null');
        assert.strictEqual(hasNaN, false, 'an empty cell emits no NaN');
        assert.strictEqual(hasInvalidDate, false, 'an empty cell emits no "Invalid Date"');
      },
    },
    {
      name: 'an invalid Date under a date number format does not leak NaN into the sheet XML',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasNaN, cellXml} = await api.dateNumFmtValueReport('invalidDate');
        assert.strictEqual(
          hasNaN,
          false,
          `an invalid Date must not serialize as <v>NaN</v>; got cell XML ${JSON.stringify(cellXml)}`,
        );
      },
    },
    {
      name: 'an invalid Date under a date number format does not leak "Invalid Date" text',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {hasInvalidDate, cellXml} = await api.dateNumFmtValueReport('invalidDate');
        assert.strictEqual(
          hasInvalidDate,
          false,
          `an invalid Date must not serialize the literal "Invalid Date"; got ${JSON.stringify(cellXml)}`,
        );
      },
    },
  ],
} satisfies Case;
