// Cluster: styles
//
// Real-world scenario: a workbook has many rows with light, repetitive formatting — every cell in a
// numeric column shares one number format, every header shares one bold font. Even when the visual
// style is identical, each cell may hold its own distinct style object in memory. On write, styles.xml
// is meant to be a SHARED table referenced by index: those identical cell styles must collapse to a
// single entry rather than emitting one style-table entry per cell. This is both a correctness
// expectation for well-formed OOXML and the mechanism that keeps write time bounded on large,
// lightly-formatted sheets — the historical performance cliff came from treating each cell's style as
// unique and re-serializing it, so the interned representation never got reused. Deduplication must
// also not over-collapse: a genuinely different style stays a distinct entry.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const fill = () => ({type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFDDEEFF'}});
const sharedCells = () =>
  Array.from({length: 40}, (_, i) => ({ref: `A${i + 1}`, value: i + 1, fill: fill()}));

export default {
  id: 'shared-styles-deduplicated-in-written-package',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'Cells carrying identical visual formatting collapse to a single shared style-table entry on ' +
    'write (one index, not one entry per cell), while a genuinely distinct style stays separate — the ' +
    'OOXML shared-table expectation that also keeps write cost bounded on large, lightly-formatted sheets.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: '40 cells carrying an identical fill resolve to a single shared style index',
      baseline: 'pass',
      async expect(api, assert) {
        const {indices} = await api.styleDedupReport(
          {sheets: [{name: 'S', cells: sharedCells()}]},
          ['A1', 'A20', 'A40'],
        );
        assert.strictEqual(
          indices.A1,
          indices.A40,
          'first and last identically-styled cells share one index',
        );
        assert.strictEqual(
          indices.A1,
          indices.A20,
          'a middle identically-styled cell shares the same index',
        );
      },
    },
    {
      name: 'the written style table does not emit one entry per identically-styled cell',
      baseline: 'pass',
      async expect(api, assert) {
        const {cellXfCount} = await api.styleDedupReport(
          {sheets: [{name: 'S', cells: sharedCells()}]},
          [],
        );
        // Default + the one shared fill = a small, bounded table — never ~40 entries.
        assert.ok(
          cellXfCount < 5,
          `40 identically-styled cells must not inflate the style table; got ${cellXfCount} entries`,
        );
      },
    },
    {
      name: 'a genuinely different style resolves to a distinct index (dedup does not over-collapse)',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = [...sharedCells(), {ref: 'B1', value: 'x', numFmt: '0.00%'}];
        const {indices} = await api.styleDedupReport({sheets: [{name: 'S', cells}]}, ['A1', 'B1']);
        assert.notStrictEqual(
          indices.A1,
          indices.B1,
          'a differently-formatted cell must keep its own style index',
        );
      },
    },
  ],
};
