// Cluster: styles
//
// Real-world scenario: a user styles a single cell — e.g. sets `A1.font = { bold:
// true }` to emphasize one heading — and expects only that cell to change. A
// regression in the legacy style engine (an "always keep the first font" change)
// caused a per-cell font assignment to bleed across the ENTIRE worksheet: every
// untouched cell reported the same font as the one cell that was set. Setting a
// font on cells of one column likewise leaked into other columns.
//
// Correct behavior: a font assigned to a cell is observable on that cell and on no
// other. Untouched cells keep the default (no font / null).

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => void }} Behavior */

export default {
  id: 'per-cell-font-isolation',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 879},
  cluster: 'styles',
  description:
    'A font set on a single cell must apply to that cell only and must not bleed ' +
    'onto untouched cells elsewhere in the worksheet.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a font set on A1 is observable on A1',
      baseline: 'pass',
      expect(api, assert) {
        const {A1} = api.probeCellFonts({
          apply: [{cell: 'A1', font: {bold: true}}],
          read: ['A1'],
        });
        assert.strictEqual(A1 && A1.bold, true, 'A1 should be bold');
      },
    },
    {
      name: 'a font set on A1 does not bleed onto an untouched cell (B2)',
      baseline: 'pass',
      expect(api, assert) {
        const {B2} = api.probeCellFonts({
          apply: [{cell: 'A1', font: {bold: true, color: {argb: 'FF3A80D5'}}}],
          read: ['B2'],
        });
        assert.ok(
          !B2 || B2.bold !== true,
          `B2 must not inherit A1's bold font, got ${JSON.stringify(B2)}`
        );
      },
    },
    {
      name: 'styling cells of one column does not bleed into another column',
      baseline: 'pass',
      expect(api, assert) {
        const {C5} = api.probeCellFonts({
          apply: [
            {cell: 'A1', font: {italic: true}},
            {cell: 'A2', font: {italic: true}},
          ],
          read: ['C5'],
        });
        assert.ok(
          !C5 || C5.italic !== true,
          `C5 (a different column) must not inherit column A's font, got ${JSON.stringify(C5)}`
        );
      },
    },
  ],
};
