// Cluster: tables
//
// Real-world scenario: a worksheet has a merged range, and the author edits rows through the
// row-level API — inserting a row above the merge, or duplicating a row — rather than through the
// splice API. Just like a splice, these edits move cell data; the merged range must move with it.
// The bug: insertRow and duplicateRow shift the cells but leave the merge record stranded at its
// original indices, so the range silently ends up covering the wrong cells (or a duplicated row is
// emitted with no merge at all). The companion case `splice-rows-preserves-merged-cells` locks the
// splice path; this one locks the distinct row-level insert/duplicate paths, which regress
// independently.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const bannerCells = [
  {ref: 'A1', value: 'header'},
  {ref: 'A2', value: 'banner'},
];

export default {
  id: 'row-insert-and-duplicate-shift-merged-cells',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'Row-level edits keep merged ranges aligned with their content: inserting a row above a merged ' +
    'banner shifts it down (A2:C2 → A3:C3), and duplicating rows above a merged range shifts that ' +
    'range down by the number of rows inserted, rather than stranding the merge at its old indices.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'inserting a row above a merged range shifts the merge down and keeps it merged',
      baseline: 'pass',
      async expect(api, assert) {
        const {merges} = await api.mutateWorksheet({
          cells: bannerCells,
          ops: [{op: 'mergeCells', range: 'A2:C2'}, {op: 'insertRow', pos: 1, value: ['inserted']}],
        });
        assert.ok(
          merges.includes('A3:C3'),
          `inserting a row above must shift the merge to A3:C3; got ${JSON.stringify(merges)}`
        );
      },
    },
    {
      name: 'duplicating rows above a merged range shifts that range down by the number inserted',
      baseline: 'pass',
      async expect(api, assert) {
        const {merges} = await api.mutateWorksheet({
          cells: [
            {ref: 'A1', value: 'a'},
            {ref: 'A3', value: 'banner'},
          ],
          // duplicate row 1 twice, inserting — this pushes the A3:C3 merge down to A5:C5.
          ops: [{op: 'mergeCells', range: 'A3:C3'}, {op: 'duplicateRow', start: 1, count: 2, insert: true}],
        });
        assert.ok(
          merges.includes('A5:C5'),
          `duplicating two rows above must shift the merge to A5:C5; got ${JSON.stringify(merges)}`
        );
      },
    },
    {
      name: 'a row insert below the merged range leaves it merged and untouched (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {merges} = await api.mutateWorksheet({
          cells: bannerCells,
          ops: [{op: 'mergeCells', range: 'A2:C2'}, {op: 'insertRow', pos: 10, value: ['below']}],
        });
        assert.deepStrictEqual(merges, ['A2:C2'], 'a row insert far below must not disturb the merge');
      },
    },
  ],
};
