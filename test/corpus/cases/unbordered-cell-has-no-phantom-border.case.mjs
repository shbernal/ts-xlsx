// Cluster: styles
//
// Real-world scenario: a consumer renders cell borders (e.g. to HTML) by checking, for each side,
// whether the cell reports a border on that side. A cell that was never given a border must therefore
// come back with no border on any side — otherwise the consumer draws borders the spreadsheet does
// not have. The reported failure is exactly that: cells with no borders in the file appear to have
// borders when read. A read/round-trip must not fabricate border sides: an unbordered cell reports no
// border, and a cell bordered on only one side does not sprout the other three.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const NO_BORDER = {sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x'}]}]};
const TOP_ONLY = {
  sheets: [{name: 'S', cells: [{ref: 'A1', value: 'x', border: {top: {style: 'thin'}}}]}],
};

export default {
  id: 'unbordered-cell-has-no-phantom-border',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A cell with no border reads back with no border sides, and a cell bordered on a single side ' +
    'does not fabricate the other sides — so a consumer never renders borders the file does not have.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'an unbordered cell reports no border sides after a round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {border} = (await api.roundtripWorkbook(NO_BORDER)).sheets.S.cells.A1;
        const sides = border
          ? ['top', 'left', 'right', 'bottom'].filter((s) => border[s]?.style)
          : [];
        assert.deepStrictEqual(
          sides,
          [],
          `no border side should be present; got ${JSON.stringify(border)}`,
        );
      },
    },
    {
      name: 'a cell bordered on one side does not sprout the other three',
      baseline: 'pass',
      async expect(api, assert) {
        const {border} = (await api.roundtripWorkbook(TOP_ONLY)).sheets.S.cells.A1;
        assert.strictEqual(border?.top?.style, 'thin', 'the declared top border survives');
        const others = ['left', 'right', 'bottom'].filter((s) => border?.[s]?.style);
        assert.deepStrictEqual(
          others,
          [],
          `only the top side should be present; got ${JSON.stringify(border)}`,
        );
      },
    },
  ],
};
