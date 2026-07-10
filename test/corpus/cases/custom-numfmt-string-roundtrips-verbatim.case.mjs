// Cluster: styles
//
// Real-world scenario: an author assigns a complex custom Excel number-format code to a cell — a
// four-section accounting/currency format with quoted currency literals, alignment placeholders, and
// group/decimal separators, e.g. `_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)`. Whatever
// format code the library is given must be stored and reproduced byte-for-byte: it must not drop the
// group-separator commas, inject stray escape characters, or otherwise rewrite the author's numFmt,
// or the cell renders differently than intended.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FMT = '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)';
const SPEC = {sheets: [{name: 'S', cells: [{ref: 'A1', value: 1234.5, numFmt: FMT}]}]};

export default {
  id: 'custom-numfmt-string-roundtrips-verbatim',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A complex multi-section custom number-format code assigned to a cell round-trips through ' +
    'write-then-read character-for-character — quoted literals, alignment placeholders, and ' +
    'group/decimal separators intact — with no dropped commas or injected escapes.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the custom format code survives a round-trip identical to the supplied string',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(model.sheets.S.cells.A1.numFmt, FMT, 'the numFmt reads back byte-for-byte');
      },
    },
    {
      name: 'the group-separator commas are not dropped from the format code',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const got = model.sheets.S.cells.A1.numFmt || '';
        assert.strictEqual((got.match(/#,##0/g) || []).length, 2, 'both #,##0 group patterns survive with their commas');
      },
    },
    {
      name: 'the cell keeps its assigned custom number format (numFmt survives)',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.ok(model.sheets.S.cells.A1.numFmt, 'the cell still carries a custom numFmt after the round-trip');
      },
    },
  ],
};
