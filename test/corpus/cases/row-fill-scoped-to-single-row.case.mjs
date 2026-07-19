// Cluster: styles
//
// Real-world scenario: a solid pattern fill is set on a single row via the row accessor (e.g. an
// orange-red fill on row 3). After a write/read round-trip only that row's cells carry the fill;
// every other cell keeps its default (empty) fill. This guards against whole-sheet style leakage —
// a row-level fill spreading across the entire sheet — which is exactly the class of silent
// corruption a regression lock must catch, and it must hold regardless of the row's index.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'r1'},
        {ref: 'A2', value: 'r2'},
        {ref: 'A3', value: 'r3'},
        {ref: 'A4', value: 'r4'},
        {ref: 'A5', value: 'r5'},
      ],
      rows: [{index: 3, fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFF4500'}}}],
    },
  ],
};

export default {
  id: 'row-fill-scoped-to-single-row',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A solid fill assigned to a single row via the row accessor stays scoped to that row after a ' +
    'round-trip — the fill does not leak onto rows above or below it, independent of the row index.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the targeted row carries the fill after round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.roundtripWorkbook(SPEC);
        const fg = sheets.S.cells.A3.fill?.fgColor;
        assert.strictEqual(fg?.argb, 'FFFF4500', 'row 3 keeps its solid fill');
      },
    },
    {
      name: 'rows above and below the targeted row keep their default (empty) fill',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.roundtripWorkbook(SPEC);
        for (const ref of ['A1', 'A2', 'A4', 'A5']) {
          assert.strictEqual(
            sheets.S.cells[ref].fill,
            undefined,
            `${ref} must not inherit the row-3 fill`,
          );
        }
      },
    },
  ],
};
