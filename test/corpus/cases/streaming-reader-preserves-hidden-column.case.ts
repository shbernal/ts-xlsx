// Cluster: streaming
//
// Real-world scenario: a worksheet is written with a hidden column, then read back through the
// streaming (row-by-row) reader rather than the in-memory reader. The streaming reader parses the
// <col> definitions but drops the hidden attribute, so a consumer iterating a streamed worksheet sees
// every column as visible even though the file marks one hidden. The eager (full-load) read is the
// oracle: it reports the hidden column correctly. The streaming reader must agree with it. This is
// the column companion to streaming-read-preserves-hidden-row, which locks the row facet.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'streaming-reader-preserves-hidden-column',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    "The streaming reader reports a column's hidden flag, agreeing with the eager read — a worksheet " +
    'written with a hidden column is not surfaced as all-visible when streamed.',

  behavior: [
    {
      name: 'the eager read sees the hidden column (oracle)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {eager} = await api.streamVsEagerColumnHidden();
        assert.strictEqual(eager.col2, true, 'the fixture hides column 2');
        assert.strictEqual(eager.col1, false, 'column 1 stays visible');
      },
    },
    {
      name: 'the streaming read reports the hidden column, not every column visible',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {eager, stream, error} = await api.streamVsEagerColumnHidden();
        assert.strictEqual(error, null, `streaming read must not error; got ${error}`);
        assert.strictEqual(
          stream.col2,
          eager.col2,
          `streaming must report column 2 hidden as the eager read does; stream=${JSON.stringify(stream)}`,
        );
      },
    },
    {
      name: 'a visible column is not reported hidden by the streaming reader (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {stream} = await api.streamVsEagerColumnHidden();
        assert.strictEqual(stream.col1, false, 'column 1 must stay visible when streamed');
        assert.strictEqual(stream.col3, false, 'column 3 must stay visible when streamed');
      },
    },
  ],
} satisfies Case;
