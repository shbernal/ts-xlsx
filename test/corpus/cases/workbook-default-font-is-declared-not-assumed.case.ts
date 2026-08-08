// Cluster: styles
//
// Font id 0 of the styles part is the workbook's default font: the face every cell that names no font
// of its own renders in — **empty cells included** — and the Maximum Digit Width every character-unit
// `<col width>` is expressed in. It is not a cell format and it is not optional; `<fonts>` always has
// a first entry, so a writer that splices in a constant is *declaring* a default rather than omitting
// one.
//
// That makes it a claim spanning two parts. `<scheme val="minor"/>` says "I am the theme's body face";
// `<name>` says which face that is. A file where the two disagree is individually well-formed in both
// parts and renders wrong: Excel resolves the explicit name, so a workbook whose theme nominates Aptos
// but whose font 0 names Calibri shows Calibri in every unstyled cell, and no amount of theme
// authoring reaches them.
//
// The round-trip half is the same fact seen from the other side. A file that declares its own font 0
// — Aptos Narrow, Arial 8, 等线 — must get it back unchanged. Replacing it with an assumed Calibri
// re-faces every empty cell and silently changes what every column width *means*, while populated
// cells keep the real face through a redundant custom entry, so the damage hides.
//
// Note 等线 in particular: its theme's latin body face is Calibri, and Excel still wrote 等线 as font
// 0, because it resolved the body face through the same `<a:minorFont>`'s `script="Hans"` entry. A
// producer resolves that face by script; we do not. So a declared face has to outrank a derived one,
// or a CJK workbook loses its typography on every save.

import type {Assert, Case, CorpusApi} from '../case.ts';

// Excel-authored packages, each declaring a font 0 the library must not overwrite.
const APTOS = 'formula-string-result-under-date-format-roundtrip/source.xlsx';
const CJK = 'builtin-cjk-date-numfmt-ids-resolve-to-date-format/source.xlsx';
const ARIAL = 'column-width-and-pagesetup-roundtrip-fidelity/sample.xlsx';

export default {
  id: 'workbook-default-font-is-declared-not-assumed',
  provenance: {source: 'round-trip-fidelity-audit'},
  cluster: 'styles',
  description:
    'The styles part’s font 0 is the workbook’s default font — what every unstyled cell renders in ' +
    'and what column widths are measured against. A file that declares one keeps it across a ' +
    'round-trip, and a workbook that authors a theme body face (or a default font outright) has it ' +
    'reach font 0 with the `scheme="minor"` claim left truthful rather than contradicting the name ' +
    'beside it.',

  behavior: [
    {
      name: 'a workbook that authors nothing declares a complete, self-consistent default font',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.defaultFontReport({});
        // Complete: a font 0 stating no size or colour is the "missing default font" foreign readers
        // (Apple Numbers, and Excel in some cases) warn about on open, because empty cells fall back
        // to a default the file never properly defines.
        assert.strictEqual(
          report.font0,
          '<font><sz val="11"/><color theme="1"/><name val="Calibri"/>' +
            '<family val="2"/><scheme val="minor"/></font>',
        );
        // Self-consistent: it claims to be the theme's body face and it is.
        assert.strictEqual(report.agreesWithTheme, true);
        assert.strictEqual(report.fontCount, 1);
      },
    },
    {
      name: 'an authored theme body face reaches font 0, so it reaches every unstyled cell',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.defaultFontReport({
          themeFonts: {major: 'Aptos Display', minor: 'Aptos'},
        });
        assert.strictEqual(report.themeMinor, 'Aptos', 'the theme part carries the authored face');
        // The whole point: the styles part follows. Without this the theme is written correctly and
        // changes nothing an unstyled cell renders, which is what forces a caller to set the font on
        // every column and name the face in every rich-text run.
        assert.strictEqual(report.font0Name, 'Aptos');
        assert.strictEqual(report.agreesWithTheme, true);
        assert.strictEqual(report.font0Scheme, 'minor');
      },
    },
    {
      name: 'a default font authored outright reaches font 0 and drops the theme claim it would falsify',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.defaultFontReport({defaultFont: {name: 'Georgia', size: 12}});
        assert.strictEqual(report.font0Name, 'Georgia');
        assert.strictEqual(report.resolved.size, 12);
        // Georgia is not the theme's body face, so claiming `scheme="minor"` would be the exact
        // contradiction this case exists to forbid — and Excel writes no `<scheme>` on such a font 0.
        assert.strictEqual(report.font0Scheme, null);
        assert.strictEqual(report.agreesWithTheme, false);
      },
    },
    {
      name: 'authoring only a size keeps the resolved face rather than blanking it',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.defaultFontReport({
          themeFonts: {minor: 'Aptos'},
          defaultFont: {size: 14},
        });
        assert.strictEqual(report.font0Name, 'Aptos');
        assert.strictEqual(report.resolved.size, 14);
      },
    },
    {
      name: 'a package’s own font 0 survives a round-trip unchanged and gains no duplicate',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.defaultFontReport({fixture: APTOS});
        assert.strictEqual(report.declared!.name, 'Aptos Narrow', 'the declaration is surfaced');
        assert.strictEqual(
          report.font0,
          '<font><sz val="11"/><color theme="1"/><name val="Aptos Narrow"/>' +
            '<family val="2"/><scheme val="minor"/></font>',
        );
        // One entry. The failure being locked out is a font 0 replaced by an assumed Calibri with the
        // real face re-added beside it — which hides, because populated cells still render right.
        assert.strictEqual(report.fontCount, 1);
        assert.strictEqual(report.reReadDeclared!.name, 'Aptos Narrow', 'and it reads back');
      },
    },
    {
      name: 'a declared face outranks the latin body face its own theme nominates',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.defaultFontReport({fixture: CJK});
        // The producer resolved the body face through the theme's `script="Hans"` entry rather than
        // its `<a:latin>` one. Deriving font 0 from the latin face would rewrite this to Calibri and
        // re-face the whole workbook on every save.
        assert.strictEqual(report.themeMinor, 'Calibri');
        assert.strictEqual(report.font0Name, '等线');
        // Its `scheme="minor"` claim rides through too: it is the producer's own resolution, not a
        // contradiction we may "fix".
        assert.strictEqual(report.font0Scheme, 'minor');
      },
    },
    {
      name: 'a default font with neither family nor scheme round-trips without gaining either',
      expect(api: CorpusApi, assert: Assert) {
        // Excel itself writes bare font 0 entries; inventing metadata for one would be a guess about
        // a face we cannot classify, and the size is what the column widths are measured against.
        const report = api.defaultFontReport({fixture: ARIAL});
        assert.strictEqual(report.font0, '<font><sz val="8"/><name val="Arial"/></font>');
      },
    },
    {
      name: 'an authored default font moves cells that only inherited the file’s, not those that named a face',
      expect(api: CorpusApi, assert: Assert) {
        // A cell carrying only a fill still names font 0, so reading resolves the declared face onto
        // it even though the source file said nothing about its font. Treating that as an authored
        // intent would strand the cell on the old face while its unstyled neighbours moved.
        const report = api.defaultFontReport({
          defaultFont: {name: 'Georgia'},
          cells: [
            {address: 'B1'},
            {address: 'B2', fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFEEEEEE'}}},
            {address: 'B3', font: {name: 'Courier New', size: 9}},
          ],
        });
        assert.strictEqual(report.font0Name, 'Georgia');
        assert.strictEqual(
          report.cellFonts.B2,
          'Georgia',
          'the inherited face follows the default',
        );
        assert.strictEqual(
          report.cellFonts.B3,
          'Courier New',
          'a named face is an intent, and stays',
        );
      },
    },
  ],
} satisfies Case;
