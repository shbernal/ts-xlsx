// Cluster: styles
//
// Real-world scenario: an author picks custom colours while building a workbook, and the application
// records them in `<colors><mruColors>` — the "Recent Colors" row it offers next time the file is
// opened. A round-trip that regenerates styles.xml drops the block, quietly resetting the author's
// working palette on every save.
//
// The block sits beside the custom indexed palette under the same `<colors>` parent, and CT_Colors
// fixes their order: `indexedColors` then `mruColors`. Emitting them the other way round produces a
// schema-invalid part, so the order is asserted here rather than left to chance.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'mru-color-swatches-survive-roundtrip/recent-colors.xlsx';

export default {
  id: 'mru-color-swatches-survive-roundtrip',
  provenance: {source: 'round-trip-fidelity-audit'},
  cluster: 'styles',
  description:
    'A workbook’s recent-colour swatches (`<colors><mruColors>`) survive a no-op round-trip, in ' +
    'the schema’s child order beside the custom indexed palette, instead of being reset on save.',

  behavior: [
    {
      name: 'the recent-colour swatches survive a round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureStylesTailFacts(FIXTURE);
        assert.deepStrictEqual(
          source.mruColors,
          ['FFBB2649', 'FF0F7173', 'FFD4A017', 'FF5B3758'],
          'precondition: the source records four recent colours',
        );
        assert.deepStrictEqual(
          rewritten.mruColors,
          source.mruColors,
          'every swatch is re-emitted, in order',
        );
      },
    },
    {
      name: 'the swatches follow the indexed palette, as CT_Colors requires',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureStylesTailFacts(FIXTURE);
        assert.deepStrictEqual(
          source.colorsChildren,
          ['indexedColors', 'mruColors'],
          'precondition: the source carries both children, in schema order',
        );
        assert.deepStrictEqual(
          rewritten.colorsChildren,
          ['indexedColors', 'mruColors'],
          'the re-emitted <colors> keeps both, in that order',
        );
        assert.strictEqual(
          rewritten.indexedColorCount,
          source.indexedColorCount,
          'and the indexed palette beside them is not truncated',
        );
      },
    },
  ],
} satisfies Case;
