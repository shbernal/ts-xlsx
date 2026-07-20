// Cluster: streaming
//
// Real-world scenario: a caller uses the streaming workbook writer to produce a spreadsheet directly
// to a file destination, then awaits commit() to finalize the archive. If the destination cannot be
// opened for writing — an invalid path, or one exceeding the OS filename length limit — the underlying
// write stream errors. The commit promise must reject with that I/O error, not hang forever leaving
// the caller awaiting a promise that never settles. A commit to a valid destination still resolves.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'streaming-commit-rejects-on-unwritable-destination',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'Committing a streaming workbook whose output destination cannot be opened rejects the commit ' +
    'promise (carrying the I/O error) rather than hanging forever.',

  behavior: [
    {
      name: 'commit to an unwritable destination rejects rather than hanging',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {outcome, rejected} = await api.streamCommitBadDestination();
        assert.strictEqual(
          rejected,
          true,
          `commit must reject on a failed sink; outcome was "${outcome}"`,
        );
      },
    },
    {
      name: 'the rejection carries the underlying I/O error',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {carriesIoError, error} = await api.streamCommitBadDestination();
        assert.strictEqual(
          carriesIoError,
          true,
          `the rejection names the I/O failure; got ${JSON.stringify(error)}`,
        );
      },
    },
  ],
} satisfies Case;
