// Cluster: streaming
//
// Real-world scenario: a sheet has data on row 1, then several blank rows, then data resuming on
// rows 6–8. Read row-by-row with the streaming reader, each yielded row must carry its true sheet
// index in `row.number` — so a consumer can map streamed rows back to their original positions —
// preserving the numeric gap where the blank rows sit (1 then 6, not a resequenced 1,2). The eager
// (fully-loaded) reader reports the true row numbers; the two read paths must agree.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'streaming-reader-preserves-blank-row-numbers/source.xlsx';

export default {
  id: 'streaming-reader-preserves-blank-row-numbers',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    'The streaming reader preserves each data row\'s true sheet index in row.number across interior ' +
    'blank rows, agreeing with the eager read — a row after a run of blanks keeps its absolute ' +
    'number rather than being shifted up by the count of skipped blanks.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'streaming row numbers match the eager row numbers, gaps preserved',
      baseline: 'pass',
      async expect(api, assert) {
        const {eager, streaming} = await api.streamVsEagerRowNumbers(FIXTURE);
        assert.deepStrictEqual(
          streaming,
          eager,
          `streaming must yield the same row numbers as the eager read; eager=${JSON.stringify(eager)} streaming=${JSON.stringify(streaming)}`
        );
      },
    },
    {
      name: 'a data row after a run of blank rows keeps its absolute number (not resequenced)',
      baseline: 'pass',
      async expect(api, assert) {
        const {streaming} = await api.streamVsEagerRowNumbers(FIXTURE);
        assert.ok(streaming.length >= 2, 'the fixture has multiple data rows');
        assert.strictEqual(streaming[0], 1, 'the first data row is row 1');
        assert.ok(
          streaming[1] > 2,
          `the row after the blank run must keep its true index (a gap), not be resequenced to 2; got ${JSON.stringify(streaming)}`
        );
      },
    },
  ],
};
