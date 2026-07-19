// Cluster: styles
//
// Real-world scenario: a column has a shared style (e.g. a column-level number format) covering
// several populated cells. The user targets one cell — say A2 — and gives it a fill to make just
// that cell stand out. They expect only A2 to change. A reported failure was that every other cell in
// the column that shared the original style also picked up the fill, because the per-cell style
// setter mutated a style object aliased across the cells rather than giving the targeted cell its own
// copy (the documented workaround was spreading `{...cell.style, fill}` into a fresh object).
// Correct behavior is copy-on-write: a single-cell style assignment affects only that cell and never
// bleeds to its column siblings or the column default.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const YELLOW = 'FFFFFF00';
const SPEC = {
  sheets: [
    {
      name: 'S',
      columns: [{index: 1, numFmt: '0.00'}], // shared column-level style
      cells: [
        {ref: 'A1', value: 1},
        {ref: 'A2', value: 2, fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: YELLOW}}},
        {ref: 'A3', value: 3},
      ],
    },
  ],
};

export default {
  id: 'per-cell-style-does-not-leak-to-column-siblings',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Assigning a fill to one cell in a column that shares a style changes only that cell — its ' +
    'column siblings keep their original (unfilled) style and the column-level number format is ' +
    'preserved on all of them, with no shared-style bleed.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the targeted cell gets the fill',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const a2 = model.sheets.S.cells.A2;
        assert.strictEqual(a2.fill?.fgColor?.argb, YELLOW, 'A2 carries the yellow fill');
      },
    },
    {
      name: 'the column siblings do not pick up the fill',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const cells = model.sheets.S.cells;
        // The writer may add a benign patternFill pattern="none" (no visible fill) to an unfilled
        // cell; what must NOT happen is the sibling acquiring the yellow foreground.
        const notYellow = (fill) =>
          !fill || fill.pattern === 'none' || !fill.fgColor || fill.fgColor.argb !== YELLOW;
        assert.ok(
          notYellow(cells.A1.fill),
          `A1 must not pick up the yellow fill; got ${JSON.stringify(cells.A1.fill)}`,
        );
        assert.ok(
          notYellow(cells.A3.fill),
          `A3 must not pick up the yellow fill; got ${JSON.stringify(cells.A3.fill)}`,
        );
      },
    },
    {
      name: 'the shared column number format survives on every cell',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const cells = model.sheets.S.cells;
        for (const ref of ['A1', 'A2', 'A3']) {
          assert.strictEqual(cells[ref].numFmt, '0.00', `${ref} keeps the column number format`);
        }
      },
    },
  ],
};
