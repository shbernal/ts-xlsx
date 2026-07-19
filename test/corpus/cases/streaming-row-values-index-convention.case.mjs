// Cluster: streaming
//
// Real-world scenario: a user reads a large spreadsheet row by row with the streaming reader instead
// of loading the whole workbook, then indexes into each row's value collection to pull named columns.
// Spreadsheet columns are 1-based, so a row's value array carries an empty leading slot: index 0 is
// unused and the first real cell (column A) lands at index 1. A caller who wrote code against the
// full-load reader and switches to streaming (or vice versa) must not have to re-index — the durable
// contract is that both read modes expose the SAME 1-based convention, so the leading slot is empty
// in both and column A is at index 1 in both.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const spec = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'A'},
        {ref: 'B1', value: 'B'},
        {ref: 'C1', value: 'C'},
      ],
    },
  ],
};

export default {
  id: 'streaming-row-values-index-convention',
  provenance: {source: 'upstream-issue'},
  cluster: 'streaming',
  description:
    "A streamed row's value array uses the same 1-based column indexing as a full-load read: index 0 " +
    'is an empty leading slot and column A is at index 1, so a caller can switch between streaming and ' +
    'buffered reads without re-indexing.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a streamed row exposes column A at index 1 with an empty leading slot at index 0',
      baseline: 'pass',
      async expect(api, assert) {
        const {streamed} = await api.streamVsEagerRowValues(spec, [1]);
        assert.strictEqual(
          streamed[1][0],
          null,
          'index 0 is an empty leading slot (columns are 1-based)',
        );
        assert.strictEqual(streamed[1][1], 'A', 'column A lands at index 1');
        assert.strictEqual(streamed[1][3], 'C', 'column C lands at index 3');
      },
    },
    {
      name: "a streamed row's indexing matches the full-load reader's for the same file",
      baseline: 'pass',
      async expect(api, assert) {
        const {streamed, eager} = await api.streamVsEagerRowValues(spec, [1]);
        assert.deepStrictEqual(
          streamed[1],
          eager[1],
          'streaming and buffered reads must expose the identical row-values array',
        );
      },
    },
  ],
};
