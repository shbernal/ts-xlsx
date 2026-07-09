// Cluster: xlsx-io
//
// Real-world scenario: `.xlsx` files are produced by many tools other than Excel —
// server-side generators, LibreOffice/Calc, reporting libraries. These files are valid
// OOXML but exercise corners Excel itself never emits, and a reader that matches on
// literal tag names rather than (namespace, local-name) chokes on them. Two distinct
// real-world shapes appear here:
//
//   1. A *namespace-prefixed* root: the workbook element and its children carry an
//      explicit prefix — `<x:workbook …><x:sheets><x:sheet …/></x:sheets></x:workbook>`
//      — instead of the default-namespace form Excel writes. A tag-literal reader never
//      recognizes `<sheets>`, so the workbook model ends up with no sheets and the next
//      access throws "Cannot read properties of undefined (reading 'sheets')". Some of
//      these files also carry a leading UTF-8 byte-order mark before `<?xml`.
//
//   2. Unusual *zip entry ordering*: `xl/workbook.xml` appears in the archive after a
//      worksheet part. A reader that depends on parts arriving in a fixed order parses a
//      sheet before the workbook model exists.
//
// In every case, opening and re-saving the file in Excel repairs it — proof the data is
// valid and the defect is on the read side. The durable target: OOXML parsing must be
// namespace-agnostic, BOM-tolerant, and package-order-independent.
//
// Fixtures (authored by real foreign generators, promoted from bug reports):
//   prefixed-root-bom.xlsx            — `x:`-prefixed root + leading BOM, sheet "Sheet1"
//   prefixed-root-cyrillic-sheet.xlsx — `x:`-prefixed root, non-ASCII sheet "Нотатки"
//   workbook-part-after-worksheet.xlsx— workbook.xml ordered after the worksheet part,
//                                       sheet "my fancy title"
//   libreoffice-calc.xlsx             — LibreOffice/Calc output, unprefixed (control)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const DIR = 'foreign-generator-workbooks-read-without-crashing';

export default {
  id: 'foreign-generator-workbooks-read-without-crashing',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 962},
  cluster: 'xlsx-io',
  description:
    'Workbooks produced by non-Excel generators — namespace-prefixed OOXML roots, a ' +
    'leading byte-order mark, non-ASCII sheet names, or parts ordered unusually within ' +
    'the zip — must be read without crashing and expose their real sheet names, rather ' +
    'than throwing because the reader matched literal tag names or a fixed part order.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a namespace-prefixed workbook root with a leading BOM reads and reports its sheet',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, error, sheetNames} = await api.readFixtureReport(`${DIR}/prefixed-root-bom.xlsx`);
        assert.ok(ok, `reader must not crash on a prefixed root; got error ${JSON.stringify(error)}`);
        assert.deepStrictEqual(sheetNames, ['Sheet1'], 'the real sheet name must be read');
      },
    },
    {
      name: 'a namespace-prefixed root preserves a non-ASCII sheet name',
      baseline: 'fail',
      async expect(api, assert) {
        const {ok, sheetNames} = await api.readFixtureReport(`${DIR}/prefixed-root-cyrillic-sheet.xlsx`);
        assert.ok(ok, 'reader must not crash on a prefixed root with a non-ASCII sheet name');
        assert.deepStrictEqual(sheetNames, ['Нотатки'], 'the non-ASCII sheet name must survive');
      },
    },
    {
      name: 'a workbook whose workbook.xml is ordered after the worksheet part still reads',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, sheetNames} = await api.readFixtureReport(`${DIR}/workbook-part-after-worksheet.xlsx`);
        assert.ok(ok, 'reading must not depend on zip entry order');
        assert.deepStrictEqual(
          sheetNames,
          ['my fancy title'],
          'the sheet name must come from the (later-ordered) workbook part, not a default'
        );
      },
    },
    {
      name: 'a plain foreign-generator (LibreOffice) workbook reads normally',
      baseline: 'pass',
      async expect(api, assert) {
        const {ok, sheetNames} = await api.readFixtureReport(`${DIR}/libreoffice-calc.xlsx`);
        assert.ok(ok, 'an unprefixed foreign-generator file must read');
        assert.deepStrictEqual(sheetNames, ['Sheet1'], 'sheet name read from a control file');
      },
    },
  ],
};
