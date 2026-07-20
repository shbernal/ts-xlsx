// Cluster: streaming
//
// Real-world scenario: with the streaming writer, a caller adds rows to a worksheet and commits it
// (or commits the whole workbook, which commits all sheets). If the caller then adds another row to
// that already-committed sheet — often because rows are appended from an async callback that races
// the commit — the write must fail with a clear "sheet already committed" error, not an internal
// null-property access crash that leaves the caller guessing. A cleanly-committed workbook still
// produces a valid, readable package.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'stream-add-row-after-commit-rejects-legibly',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'Adding a row to a streaming worksheet after it has been committed is rejected with a clear ' +
    '"already committed" error rather than an internal null-property crash, and a cleanly-committed ' +
    'workbook still reads back valid.',

  behavior: [
    {
      name: 'a row added after commit is rejected (not silently accepted)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {rejected} = await api.streamAddRowAfterCommit();
        assert.strictEqual(rejected, true, 'adding a row after commit must be rejected');
      },
    },
    {
      name: 'the rejection is a legible "already committed" error, not an internal null crash',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {legibleRejection, internalCrash, error} = await api.streamAddRowAfterCommit();
        assert.strictEqual(
          internalCrash,
          false,
          `the error must not be an internal null-property crash; got ${JSON.stringify(error)}`,
        );
        assert.strictEqual(
          legibleRejection,
          true,
          `the error must name the committed state; got ${JSON.stringify(error)}`,
        );
      },
    },
  ],
} satisfies Case;
