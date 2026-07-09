// Cluster: styles
//
// Real-world scenario: an Excel-authored workbook uses cell background (fill) colors and border
// colors that reference the theme and/or the indexed palette rather than plain RGB. Loading the
// file and writing it straight back — with no edits — must preserve the visually-rendered colors
// exactly. A naive style reader/writer that drops or misinterprets a theme/tint or indexed
// reference changes the appearance of a pure open-then-save.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'fill-border-color-survives-roundtrip/source.xlsx';

export default {
  id: 'fill-border-color-survives-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A pure open-then-save round-trip preserves cells\' visible fill (solid/patterned foreground) ' +
    'and border-edge colors — including theme+tint and indexed-palette references — so the sheet ' +
    'renders identically after re-saving.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the round-trip actually exercises colored cells (guard)',
      baseline: 'pass',
      async expect(api, assert) {
        const {checked} = await api.roundtripFixtureColorFidelity(FIXTURE);
        assert.ok(checked > 0, 'the fixture has styled cells with fill and/or border colors');
      },
    },
    {
      name: 'no cell\'s visible fill color changes across the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {fillMismatches, fillSample} = await api.roundtripFixtureColorFidelity(FIXTURE);
        assert.strictEqual(fillMismatches, 0, `fill colors must survive; first divergence: ${JSON.stringify(fillSample)}`);
      },
    },
    {
      name: 'no cell\'s border-edge color changes across the round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {borderMismatches, borderSample} = await api.roundtripFixtureColorFidelity(FIXTURE);
        assert.strictEqual(borderMismatches, 0, `border colors must survive; first divergence: ${JSON.stringify(borderSample)}`);
      },
    },
  ],
};
