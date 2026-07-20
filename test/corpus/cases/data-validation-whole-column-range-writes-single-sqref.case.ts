// Cluster: validations
//
// Real-world scenario: a user attaches one data validation — a list/dropdown — to an entire column
// (e.g. B2 down to the last row) rather than assigning it cell by cell. Writing must succeed and
// emit exactly one `dataValidation` whose `sqref` is the requested range, not one entry per covered
// cell (which would balloon the file and the parse). The whole-column form is the natural,
// performant way to put a dropdown on a column, so it must serialize without throwing and reload
// cleanly. (Reading such a validation in bounded memory is a separate hostile-input concern captured
// in `whole-column-data-validation-bounded-memory`; this case locks the write side.)

import type {Assert, Case, CorpusApi} from '../case.ts';

const RANGE = 'B2:B1048576';

export default {
  id: 'data-validation-whole-column-range-writes-single-sqref',
  provenance: {source: 'upstream-issue'},
  cluster: 'validations',
  description:
    'A single data validation applied over a whole-column range serializes without throwing, emits ' +
    'exactly one dataValidation whose sqref is the requested range (not per-cell entries), and the ' +
    'written file reloads without error.',

  behavior: [
    {
      name: 'writing a whole-column-range validation does not throw during serialization',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {writeOk, writeError} = await api.roundtripRangeValidation({range: RANGE});
        assert.strictEqual(
          writeOk,
          true,
          `serialization must not throw; got ${JSON.stringify(writeError)}`,
        );
      },
    },
    {
      name: 'the package contains exactly one dataValidation whose sqref is the requested range',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {count, sqrefs} = await api.roundtripRangeValidation({range: RANGE});
        assert.strictEqual(count, 1, 'one dataValidation entry, not one per covered cell');
        assert.deepStrictEqual(
          sqrefs,
          [RANGE],
          `the sqref is the whole-column range; got ${JSON.stringify(sqrefs)}`,
        );
      },
    },
    {
      name: 'the written file reloads without error',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {reloadOk} = await api.roundtripRangeValidation({range: RANGE});
        assert.strictEqual(
          reloadOk,
          true,
          'the whole-column validation reads back without throwing',
        );
      },
    },
  ],
} satisfies Case;
