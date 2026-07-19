// Cluster: tables
//
// Real-world scenario: a worksheet has a horizontally merged range below the top of the sheet — a
// header row, then a banner merged across A2:O2. The user deletes a row above the merged range
// (a row-splice delete), which shifts the merged range upward. After the delete the range must
// still be merged, now at A1:O1 — the same span, shifted. The same row-shifting logic underlies
// row insertion (shift down) and duplication, so merged ranges must survive those too. The bug:
// the splice moves the cell data but leaves the merge range stranded at its original indices, so
// the range silently un-merges (and points at the wrong, now-empty cells).

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const bannerSheet = (ops) => ({
  cells: [
    {ref: 'A1', value: 'header'},
    {ref: 'A2', value: 'banner'},
  ],
  ops: [{op: 'mergeCells', range: 'A2:O2'}, ...ops],
  read: ['A1'],
});

export default {
  id: 'splice-rows-preserves-merged-cells',
  provenance: {source: 'upstream-issue'},
  cluster: 'tables',
  description:
    'A row-splice that shifts a merged range keeps it merged at its new position: deleting a row ' +
    'above a merged banner shifts it up (A2:O2 → A1:O1), inserting a row above shifts it down ' +
    '(A2:O2 → A3:O3), and a splice below the range leaves it untouched.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a splice below the merged range leaves it merged and untouched (control)',
      baseline: 'pass',
      async expect(api, assert) {
        const {merges} = await api.mutateWorksheet(
          bannerSheet([{op: 'spliceRows', start: 10, count: 1}]),
        );
        assert.deepStrictEqual(merges, ['A2:O2'], 'a splice far below must not disturb the merge');
      },
    },
    {
      name: 'deleting a row above the merged range shifts it up and keeps it merged',
      baseline: 'pass',
      async expect(api, assert) {
        const {merges} = await api.mutateWorksheet(
          bannerSheet([{op: 'spliceRows', start: 1, count: 1}]),
        );
        assert.ok(
          merges.includes('A1:O1'),
          `deleting row 1 must shift the merge to A1:O1 and keep it merged; got ${JSON.stringify(merges)}`,
        );
      },
    },
    {
      name: 'inserting a row above the merged range shifts it down and keeps it merged',
      baseline: 'pass',
      async expect(api, assert) {
        const {merges} = await api.mutateWorksheet(
          bannerSheet([{op: 'spliceRows', start: 1, count: 0, inserts: [['inserted']]}]),
        );
        assert.ok(
          merges.includes('A3:O3'),
          `inserting a row above must shift the merge to A3:O3 and keep it merged; got ${JSON.stringify(merges)}`,
        );
      },
    },
  ],
};
