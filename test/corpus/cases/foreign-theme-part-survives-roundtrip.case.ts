// Cluster: styles
//
// Real-world scenario: a workbook carries a branded theme — a custom <clrScheme> and <fontScheme> in
// the theme part. That part is what every `theme="n"` colour reference and every
// `scheme="major|minor"` font in the file resolves against, so a no-op round-trip that replaces it
// with the default Office theme leaves the cells untouched yet silently re-renders the whole
// workbook in the wrong brand colours and typefaces. The theme a source package declares must
// survive a read→write unchanged.
//
// Two traps ride along. The part is reached through the workbook's `.../theme` relationship, whose
// target is rel-relative — `theme1.xml` is a convention, not a rule, so one fixture names it
// `theme2.xml`. And a theme can carry relationships of its own: a picture used as a themed fill is
// wired by an `r:embed` into the theme's rels part. Re-emitting the theme body without that closure
// is worse than dropping the theme — it leaves a dangling reference, which Excel reports as a
// package needing repair — so the closure travels with it.

import type {Assert, Case, CorpusApi} from '../case.ts';

const BRANDED = 'foreign-theme-part-survives-roundtrip/branded-theme.xlsx';
const THEMED_IMAGE = 'foreign-theme-part-survives-roundtrip/themed-image-fill.xlsx';

export default {
  id: 'foreign-theme-part-survives-roundtrip',
  provenance: {source: 'round-trip-fidelity-audit'},
  cluster: 'styles',
  description:
    'A workbook’s theme part — its colour scheme and font scheme, plus any parts it references — ' +
    'survives a no-op round-trip, instead of being overwritten by the default Office theme and ' +
    'silently re-rendering every themed colour and font in the file.',

  behavior: [
    {
      name: 'the source colour scheme survives a round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureThemeFacts(BRANDED);
        assert.strictEqual(
          source.colors.accent1,
          'BB2649',
          'precondition: the source declares a non-default accent1',
        );
        assert.deepStrictEqual(
          rewritten.colors,
          source.colors,
          'every colour slot the source theme declared is re-emitted unchanged',
        );
      },
    },
    {
      name: 'the source font scheme survives a round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureThemeFacts(BRANDED);
        assert.strictEqual(
          source.majorFont,
          'Bodoni MT',
          'precondition: the source declares a non-default major typeface',
        );
        assert.strictEqual(rewritten.majorFont, source.majorFont, 'the major typeface survives');
        assert.strictEqual(rewritten.minorFont, source.minorFont, 'the minor typeface survives');
      },
    },
    {
      name: 'a theme whose fill references a picture keeps that picture reachable',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {source, rewritten} = await api.roundtripFixtureThemeFacts(THEMED_IMAGE);
        assert.strictEqual(
          source.relTargets.length,
          1,
          'precondition: the source theme references one part of its own',
        );
        assert.strictEqual(
          rewritten.relTargets.length,
          1,
          'the re-emitted theme still declares its relationship',
        );
        assert.ok(
          rewritten.relTargetsResolve,
          'the relationship target names a part the written package actually holds',
        );
      },
    },
    {
      name: 'a theme part not named theme1.xml is still found, and still lands somewhere wired',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        // The relationship target is rel-relative: the conventional `theme1.xml` is a convention,
        // not a rule. The reader must follow the relationship, and the writer must re-emit the part
        // at whatever path its own regenerated theme relationship names — a theme written anywhere
        // else leaves that relationship, and the content-type override, pointing at nothing.
        const {source, rewritten} = await api.roundtripFixtureThemeFacts(BRANDED);
        assert.strictEqual(
          source.path,
          'xl/theme/theme2.xml',
          'precondition: the source names its theme part unconventionally',
        );
        assert.ok(
          rewritten.present,
          'the written theme relationship resolves to a part that exists',
        );
        assert.strictEqual(
          rewritten.name,
          'Harbour',
          'and it is the source theme, not the default',
        );
      },
    },
  ],
} satisfies Case;
