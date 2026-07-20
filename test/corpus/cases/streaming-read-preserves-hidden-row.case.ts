// Cluster: streaming
//
// Real-world scenario: a worksheet has a hidden row (its <row> carries hidden="true" while others
// carry hidden="false"). Read through the streaming (row-by-row) reader, each row's hidden flag must
// reflect the XML. In the reported failure the streaming reader returned hidden=false for every row,
// including the one explicitly hidden, so a consumer iterating in streaming mode could not tell
// hidden from visible rows. The boolean is written in string form ("true"/"false") by some
// generators, so the parser must interpret those, not only the "1"/"0" or attribute-presence forms.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'streaming-read-preserves-hidden-row/source.xlsx';

export default {
  id: 'streaming-read-preserves-hidden-row',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'The streaming reader reports a row\'s hidden flag (interpreting the string-form "true"/"false" ' +
    'some generators write), agreeing with the eager read — not reporting every row visible.',

  behavior: [
    {
      name: 'the eager read sees the hidden row (oracle)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {eager} = await api.streamVsEagerRowHidden(FIXTURE);
        assert.ok(
          eager.some((r: CorpusApi) => r.hidden),
          `the fixture has a hidden row; eager=${JSON.stringify(eager)}`,
        );
      },
    },
    {
      name: 'the streaming read reports the same hidden flags as the eager read',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {eager, streaming} = await api.streamVsEagerRowHidden(FIXTURE);
        const eagerHidden = eager
          .filter((r: CorpusApi) => r.hidden)
          .map((r: CorpusApi) => r.number);
        const streamHidden = streaming
          .filter((r: CorpusApi) => r.hidden)
          .map((r: CorpusApi) => r.number);
        assert.deepStrictEqual(
          streamHidden,
          eagerHidden,
          `streaming must report the hidden rows the eager read sees; eager=${JSON.stringify(eager)} streaming=${JSON.stringify(streaming)}`,
        );
      },
    },
  ],
} satisfies Case;
