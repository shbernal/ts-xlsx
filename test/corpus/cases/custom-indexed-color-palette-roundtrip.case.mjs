// Cluster: styles
//
// Real-world scenario: a workbook overrides the legacy indexed-color palette with a custom
// <indexedColors> block in styles.xml, and cells/fonts/borders reference colors by index into
// that palette. A no-op round-trip drops the custom palette entirely, so every indexed color
// silently resolves to a different RGB (the default palette entry) — colors change on save. The
// custom indexed-color palette must survive a round-trip so indexed color references keep their
// intended RGB.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'custom-indexed-color-palette-roundtrip/sample.xlsx';

export default {
  id: 'custom-indexed-color-palette-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1224},
  cluster: 'styles',
  description:
    'A workbook’s custom <indexedColors> palette survives a no-op round-trip so indexed color ' +
    'references keep their intended RGB, instead of the palette being dropped and colors shifting.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the custom indexed-color palette is re-emitted on round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixtureStyleFacts(FIXTURE);
        assert.ok(source.hasIndexedColors, 'precondition: the source declares a custom palette');
        assert.ok(rewritten.hasIndexedColors, 'the written styles part must re-emit the custom indexedColors block');
      },
    },
    {
      name: 'the palette slot-to-RGB mapping is preserved',
      baseline: 'pass',
      async expect(api, assert) {
        const {source, rewritten} = await api.roundtripFixtureStyleFacts(FIXTURE);
        assert.deepStrictEqual(rewritten.indexedColorSample, source.indexedColorSample, 'the leading palette entries match');
      },
    },
  ],
};
