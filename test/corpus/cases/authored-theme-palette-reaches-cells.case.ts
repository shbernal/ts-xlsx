// Cluster: styles
//
// Real-world scenario: a team wants their workbooks in their own brand colours. In a spreadsheet that
// means the *theme* palette, not per-cell fills: Excel's colour picker offers the theme row first, and
// a colour chosen from it is written as `theme="4"` — a reference, resolved at render time. Setting
// `accent1` therefore restyles every cell, chart and table style that follows the theme at once, and
// is the only way to recolour a workbook without touching a single cell.
//
// Authoring generates *over* the existing theme rather than replacing it, which is what the assertions
// here are mostly about:
//
//  • The format scheme — the gradient, line and effect styles that give a theme its texture — is a
//    designer's work that no spreadsheet API hand-authors. It rides through untouched.
//  • A slot the caller did not name keeps its source **encoding**, not just its value. `dk1`/`lt1` are
//    `<a:sysClr>` so they follow the viewer's window colours; rewriting them as `<a:srgbClr>` would
//    pin them to whatever one machine resolved them to.
//  • A theme that carries its own relationships (a picture used as a themed fill) keeps them.

import type {Assert, Case, CorpusApi} from '../case.ts';

const PRESERVED_THEME = 'foreign-theme-part-survives-roundtrip/themed-image-fill.xlsx';

export default {
  id: 'authored-theme-palette-reaches-cells',
  provenance: {source: 'round-trip-fidelity-audit'},
  cluster: 'styles',
  description:
    'Authoring the workbook theme’s colour scheme and typefaces reaches the emitted theme part and ' +
    'every cell that references a slot by `theme="n"`, while the format scheme, the unauthored ' +
    'slots’ encoding, and the theme’s own relationships ride through untouched.',

  behavior: [
    {
      name: 'an authored accent reaches the theme part and the cells that reference it',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.authorThemeReport({colors: {accent1: 'BB2649'}});
        assert.strictEqual(
          report.scheme.accent1,
          'BB2649',
          'the theme part carries the new accent',
        );
        // The point of authoring a palette rather than a fill: a cell that says `theme="4"` follows.
        assert.strictEqual(report.resolvedThemeColor, 'FFBB2649');
        // …and the written package says so, not just the in-memory model.
        assert.strictEqual(report.reReadScheme.accent1, 'BB2649');
      },
    },
    {
      name: 'unauthored slots are left exactly as they were',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.authorThemeReport({colors: {accent1: 'BB2649'}});
        // Office defaults, untouched by an author who named only accent1.
        assert.strictEqual(report.scheme.accent2, 'ED7D31');
        assert.strictEqual(report.scheme.accent6, '70AD47');
        assert.strictEqual(report.scheme.hlink, '0563C1');
        // And, crucially, their *encoding*: dk1/lt1 stay system colours.
        assert.deepStrictEqual(report.encodings, {
          dk1: 'sysClr',
          lt1: 'sysClr',
          accent1: 'srgbClr',
        });
      },
    },
    {
      name: 'the format scheme survives an authored palette',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.authorThemeReport({colors: {accent1: 'BB2649'}});
        assert.strictEqual(report.keptFmtScheme, true, 'the <a:fmtScheme> block is still there');
        // Not merely present — still populated. A regenerated theme would flatten the gradients.
        assert.ok(
          report.fmtSchemeGradientStops > 0,
          `the format scheme still carries its gradient stops (found ${report.fmtSchemeGradientStops})`,
        );
      },
    },
    {
      name: 'the major and minor typefaces are authorable',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        // A font that says `scheme="minor"` names no typeface of its own, so this is the only place
        // the workbook's body face is decided.
        const report = await api.authorThemeReport({fonts: {major: 'Bodoni MT', minor: 'Cambria'}});
        assert.deepStrictEqual(report.fonts, {major: 'Bodoni MT', minor: 'Cambria'});
      },
    },
    {
      name: 'authoring over a source theme keeps the rest of that theme, relationships included',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.authorThemeReport({
          fixture: PRESERVED_THEME,
          colors: {accent3: '112233'},
        });
        assert.strictEqual(report.scheme.accent3, '112233', 'the authored slot changed');
        // The source theme's own values, not the Office defaults — the base is the file's theme.
        assert.strictEqual(report.schemeName, 'Harbour');
        assert.strictEqual(report.scheme.accent1, 'BB2649');
        assert.strictEqual(report.scheme.dk2, '1B3A4B');
        assert.strictEqual(report.fonts.major, 'Bodoni MT');
        // The picture this theme uses as a fill is still reachable: authoring must not sever the
        // theme's relationships, which would leave an `r:embed` dangling.
        assert.strictEqual(report.hasThemeRels, true);
        assert.strictEqual(report.mediaParts, 1);
      },
    },
    {
      name: 'a malformed theme colour is refused at the call that supplied it',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        // Excel does not report a malformed colour value — it renders the slot as flat black — so the
        // library has to, and at the setter rather than at write time far from the cause.
        assert.match(await api.authorInvalidThemeColor('not-a-colour'), /Invalid theme colour/);
        assert.match(await api.authorInvalidThemeColor('#12345'), /Invalid theme colour/);
        // The two conveniences the rest of the library accepts are accepted here too.
        assert.strictEqual(await api.authorInvalidThemeColor('#BB2649'), null);
        assert.strictEqual(await api.authorInvalidThemeColor('FFBB2649'), null);
      },
    },
  ],
} satisfies Case;
