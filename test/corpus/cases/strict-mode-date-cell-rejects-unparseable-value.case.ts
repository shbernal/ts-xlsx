// Cluster: dates
//
// Real-world scenario: an ISO/IEC 29500 Strict workbook stores a date cell as `t="d"` with an ISO 8601
// string in `<v>` rather than as a 1900-epoch serial. A generator that emits a locale-formatted or
// truncated timestamp there produces a cell whose text is not a date at all. `new Date(text)` does not
// fail on that: it returns a Date whose time is `NaN`, which passes `instanceof Date` and every guard
// the model applies, survives into the workbook, and serialises back out as the literal string
// `Invalid Date`: a value nobody stored, in a cell that used to say what was wrong with it.
//
// Dropping the value at the boundary is the only reading that cannot lie: the cell is empty, which is
// true, instead of holding a date that is not one.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'strict-mode-date-cell-rejects-unparseable-value',
  cluster: 'dates',
  description:
    'A Strict-mode `t="d"` cell whose `<v>` is not a parseable ISO 8601 date reads as empty rather ' +
    'than putting an Invalid Date into the model, while a well-formed sibling still reads as the ' +
    'instant it states.',
  provenance: {source: 'audit'},

  behavior: [
    {
      name: 'an unparseable `t="d"` value never reaches the model as a Date',
      expect(api: CorpusApi, assert: Assert) {
        const {unparseableIsDate} = api.strictModeDateCellReport();
        assert.strictEqual(
          unparseableIsDate,
          false,
          'an Invalid Date satisfies every downstream guard and writes back as the string `Invalid Date`',
        );
      },
    },
    {
      name: 'the cell holding it is empty rather than holding a value nobody stored',
      expect(api: CorpusApi, assert: Assert) {
        const {unparseable} = api.strictModeDateCellReport();
        assert.strictEqual(unparseable, null);
      },
    },
    {
      name: 'a well-formed `t="d"` sibling still reads as the instant it states (control)',
      expect(api: CorpusApi, assert: Assert) {
        const {parsed} = api.strictModeDateCellReport();
        assert.strictEqual(parsed, '2024-03-04T05:06:07.000Z');
      },
    },
  ],
} satisfies Case;
