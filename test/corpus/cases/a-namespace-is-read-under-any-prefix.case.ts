// Cluster: security
//
// Real-world scenario: a namespace prefix is a local nickname, not an identity. A file may bind the
// relationships namespace to `rel:` instead of `r:`, or the main SpreadsheetML namespace to `x:`
// instead of leaving it as the default, or DrawingML to nothing at all. All three are legal and real
// toolchains emit all three.
//
// The reader is prefix-agnostic almost everywhere, which is exactly what made the handful of places
// testing a prefix STRING so damaging: the rest of the file reads perfectly and one whole feature
// disappears, with nothing thrown and nothing logged.
//
//   - `attrs['r:id']` found nothing in a workbook that bound relationships to another prefix, so every
//     sheet loaded permanently empty;
//   - `name.includes(':')` as a stand-in for "this is an extension element" is true of EVERY element
//     in a worksheet that prefixes the main namespace, so every data validation and conditional format
//     was discarded as unknown;
//   - the theme reader hardcoded `a:`, so a theme under any other prefix read as no theme at all, the
//     workbook silently fell back to the default Office palette, and every theme-indexed colour in the
//     document resolved to the wrong RGB.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'a-namespace-is-read-under-any-prefix',
  provenance: {source: 'namespace-audit'},
  cluster: 'security',
  description:
    'A package is read the same whichever prefix it binds a namespace to: the relationships ' +
    'namespace under any prefix still wires each sheet to its part, a worksheet that prefixes the ' +
    'main namespace keeps its data validations and conditional formats, and a theme under any ' +
    'prefix (or none) resolves its own colours rather than falling back to the default palette.',

  behavior: [
    {
      name: 'a sheet relationship is found whatever prefix names the relationships namespace',
      expect(api: CorpusApi, assert: Assert) {
        for (const prefix of ['r', 'rel', 'q1']) {
          const report = api.relationshipPrefixReport(prefix);
          assert.deepEqual(report.sheetNames, ['S1'], `prefix ${prefix}: the sheet is declared`);
          assert.equal(report.a1, 'hello', `prefix ${prefix}: and its part was actually loaded`);
        }
      },
    },
    {
      name: 'a worksheet that prefixes the main namespace keeps its data validations',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.mainNamespacePrefixReport('x');
        assert.equal(report.a1, 1, 'the cells read, which they always did');
        assert.deepEqual(report.validations, ['A1']);
      },
    },
    {
      name: 'and it keeps its conditional formats',
      expect(api: CorpusApi, assert: Assert) {
        assert.deepEqual(api.mainNamespacePrefixReport('x').conditionalFormats, [
          'duplicateValues',
        ]);
      },
    },
    {
      name: "a theme's colours resolve under any prefix, including the default namespace",
      expect(api: CorpusApi, assert: Assert) {
        for (const prefix of ['a', '', 'dml']) {
          const {colors} = api.themePrefixReport(prefix);
          assert.equal(colors.accent1, 'AABBCC', `prefix ${JSON.stringify(prefix)}: accent1`);
          assert.equal(colors.dk1, '111111', 'a sysClr slot resolves through its lastClr');
          assert.equal(
            colors.lt1,
            '222222',
            'and a slot with an XML comment before its colour is not lost',
          );
        }
      },
    },
    {
      name: "a theme's fonts resolve under any prefix too",
      expect(api: CorpusApi, assert: Assert) {
        for (const prefix of ['a', '', 'dml']) {
          const {fonts} = api.themePrefixReport(prefix);
          assert.equal(fonts.major, 'Major & Co', `prefix ${JSON.stringify(prefix)}`);
          assert.equal(fonts.minor, 'Minor');
        }
      },
    },
    {
      name: "a theme-indexed cell colour resolves against the file's own theme, not the default",
      expect(api: CorpusApi, assert: Assert) {
        // The whole-document consequence: falling back to the library's theme silently repaints every
        // theme-coloured cell in the workbook.
        for (const prefix of ['a', '', 'dml']) {
          const {cellColor} = api.themePrefixReport(prefix);
          assert.equal(cellColor?.theme, 4, `prefix ${JSON.stringify(prefix)}: the link survives`);
        }
      },
    },
  ],
} satisfies Case;
