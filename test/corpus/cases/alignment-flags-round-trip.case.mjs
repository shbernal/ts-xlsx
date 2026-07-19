// Cluster: styles
//
// Real-world scenario: cells carry boolean alignment flags — wrapText and
// shrinkToFit — that are OFF for the vast majority of cells. A reader must report
// these as they truly are: a cell that never enabled wrapText must not come back
// claiming wrapText is on. The failure mode is a reader defaulting every cell's
// flags to `true`, which silently rewraps or shrinks content the author never
// asked to touch. Setting a flag must round-trip as set; leaving it off must
// round-trip as off.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void> }} Behavior */

const WRAP_ON = {
  sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x', alignment: {wrapText: true}}]}],
};
const FLAGS_OFF = {
  sheets: [
    {name: 'S', cells: [{ref: 'A1', value: 'x', alignment: {wrapText: false, shrinkToFit: false}}]},
  ],
};
const PLAIN = {sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x'}]}]};
const INDENT = {sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x', alignment: {indent: 3}}]}]};

export default {
  id: 'alignment-flags-round-trip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1908},
  cluster: 'styles',
  description:
    'The boolean alignment flags wrapText and shrinkToFit round-trip as their true ' +
    'value — an enabled flag survives, and a disabled or unset flag does not come ' +
    'back spuriously enabled.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'wrapText set true survives the round-trip as true',
      baseline: 'pass',
      async expect(api, assert) {
        const {alignment} = (await api.roundtripWorkbook(WRAP_ON)).sheets.S.cells.A1;
        assert.strictEqual(alignment?.wrapText, true, 'wrapText should survive as true');
      },
    },
    {
      name: 'wrapText and shrinkToFit left off do not come back true',
      baseline: 'pass',
      async expect(api, assert) {
        const {alignment} = (await api.roundtripWorkbook(FLAGS_OFF)).sheets.S.cells.A1;
        assert.notStrictEqual(alignment?.wrapText, true, 'wrapText must not be spuriously true');
        assert.notStrictEqual(
          alignment?.shrinkToFit,
          true,
          'shrinkToFit must not be spuriously true',
        );
      },
    },
    {
      name: 'a cell with no alignment does not report the flags as true',
      baseline: 'pass',
      async expect(api, assert) {
        const {alignment} = (await api.roundtripWorkbook(PLAIN)).sheets.S.cells.A1;
        assert.notStrictEqual(alignment?.wrapText, true, 'wrapText must not be spuriously true');
        assert.notStrictEqual(
          alignment?.shrinkToFit,
          true,
          'shrinkToFit must not be spuriously true',
        );
      },
    },
    {
      name: 'a numeric indent level survives the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {alignment} = (await api.roundtripWorkbook(INDENT)).sheets.S.cells.A1;
        assert.strictEqual(alignment?.indent, 3, 'the indent level should survive as set');
      },
    },
  ],
};
