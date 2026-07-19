// Cluster: formulas
//
// Real-world scenario: cell A1 holds text ("test") but is styled with a date number format
// (mmm-yy). Cell A2 is the formula =A1, so its cached result is the string "test" and A2 inherits
// a date/number format. In the source file A2 is a string-typed formula cell (t="str") with the
// cached value "test". Loading the workbook and writing it straight back out — with no user edits
// — must produce a file that still opens cleanly. The observed corruption: the writer drops A2's
// string type and, because the cell's style carries a numeric/date format, coerces the string
// cached result toward a number and emits the literal token "NaN" as the value. A numeric cell
// containing "NaN" is invalid content, so Excel shows a "we found a problem… recover?" prompt.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'formula-string-result-under-date-format-roundtrip/source.xlsx';

export default {
  id: 'formula-string-result-under-date-format-roundtrip',
  provenance: {source: 'upstream-issue'},
  cluster: 'formulas',
  description:
    'A string-typed formula cell (=A1 whose cached result is the text "test") that is styled with ' +
    'a date number format round-trips as a string formula cell — the writer must not drop the ' +
    'string type and emit the invalid token "NaN" as the cell value (which corrupts the file).',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'no cell in the re-written package carries the invalid literal "NaN" as its value',
      baseline: 'pass',
      async expect(api, assert) {
        const {hasNaNToken} = await api.roundtripFixtureCellXml(FIXTURE, ['A2']);
        assert.strictEqual(
          hasNaNToken,
          false,
          'a round-trip must never emit "NaN" as a cell value',
        );
      },
    },
    {
      name: 'the string-typed formula cell keeps its formula and is not written as a numeric cell',
      baseline: 'pass',
      async expect(api, assert) {
        const {cells} = await api.roundtripFixtureCellXml(FIXTURE, ['A2']);
        assert.ok(cells.A2, 'A2 must survive the round-trip');
        assert.strictEqual(cells.A2.formula, 'A1', 'A2 must keep its =A1 formula');
        assert.strictEqual(
          cells.A2.value,
          'test',
          `A2's cached string result must round-trip as "test", not be coerced; got ${JSON.stringify(cells.A2)}`,
        );
        assert.strictEqual(
          cells.A2.t,
          'str',
          'a string-typed formula cell must stay t="str", not become numeric',
        );
      },
    },
  ],
};
