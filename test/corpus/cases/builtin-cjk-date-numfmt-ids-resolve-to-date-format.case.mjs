// Cluster: types
//
// Real-world scenario: a workbook authored in a CJK (Chinese) Excel locale stores date cells as
// serial numbers styled with a locale-specific *built-in* number-format id. Excel reserves format
// ids in the 27..58 range for locale-specific built-in date/time formats (e.g. id 57 renders like
// "yyyy年m月", id 31 like "yyyy年m月d日"); because they are built-ins, Excel does NOT emit an
// explicit custom <numFmt> for them in styles.xml. A reader whose built-in numFmt table only covers
// the standard Western ids (0..26, 37..49…) does not recognize such a cell as a date: the value
// comes back as a bare number and the effective number-format code is empty, so downstream date
// detection and display both break. The fixture has date columns styled with built-in ids 57 and 31
// holding serials such as 45809 and 45819.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIXTURE = 'builtin-cjk-date-numfmt-ids-resolve-to-date-format/source.xlsx';

export default {
  id: 'builtin-cjk-date-numfmt-ids-resolve-to-date-format',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    'A cell styled with a locale-specific built-in date number-format id (e.g. 57 or 31) resolves ' +
    'to a non-empty date format code and reads as a date, not a bare number — the built-in numFmt ' +
    'table must cover the 27..58 locale-specific date ids, not only the Western ids.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a cell with a built-in CJK date format id exposes a non-empty number-format code',
      baseline: 'fail',
      async expect(api, assert) {
        const cells = await api.readFixtureCells(FIXTURE, ['A2', 'B2']);
        assert.ok(
          cells.A2.numFmt,
          `a built-in locale date id must resolve to a format code, not empty; got ${JSON.stringify(cells.A2)}`
        );
      },
    },
    {
      name: 'a cell with a built-in CJK date format id reads as a date, not a bare number',
      baseline: 'fail',
      async expect(api, assert) {
        const cells = await api.readFixtureCells(FIXTURE, ['A2', 'B2']);
        assert.strictEqual(
          cells.A2.type,
          'date',
          `a serial styled with a built-in date id must read as a date; got ${JSON.stringify(cells.A2)}`
        );
      },
    },
  ],
};
