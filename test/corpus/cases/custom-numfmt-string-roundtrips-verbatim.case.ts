// Cluster: styles
//
// Real-world scenario: an author assigns a complex custom Excel number-format code to a cell — a
// four-section accounting/currency format with quoted currency literals, alignment placeholders, and
// group/decimal separators, e.g. `_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)`. Whatever
// format code the library is given must be stored and reproduced byte-for-byte: it must not drop the
// group-separator commas, inject stray escape characters, or otherwise rewrite the author's numFmt,
// or the cell renders differently than intended.

import type {Assert, Case, CorpusApi} from '../case.ts';

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

  behavior: [
    {
      name: 'the custom format code survives a round-trip identical to the supplied string',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(
          model.sheets.S.cells.A1.numFmt,
          FMT,
          'the numFmt reads back byte-for-byte',
        );
      },
    },
    {
      name: 'the group-separator commas are not dropped from the format code',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const got = model.sheets.S.cells.A1.numFmt || '';
        assert.strictEqual(
          (got.match(/#,##0/g) || []).length,
          2,
          'both #,##0 group patterns survive with their commas',
        );
      },
    },
    {
      name: 'the cell keeps its assigned custom number format (numFmt survives)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.ok(
          model.sheets.S.cells.A1.numFmt,
          'the cell still carries a custom numFmt after the round-trip',
        );
      },
    },
    {
      // Number-format codes are stored in an invariant form where "." is always the decimal and ","
      // the grouping separator (and "/" the date separator); the viewer localizes at display time.
      // The library must persist the user's separators verbatim — never swap "." <-> "," or rewrite
      // "/" to "-" — so a comma-decimal locale renders faithfully from the invariant code.
      name: 'invariant separators in a percentage and a date format code survive verbatim',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook({
          sheets: [
            {
              name: 'S',
              cells: [
                {ref: 'A1', value: 0.5, numFmt: '0.00%'},
                {ref: 'A2', value: {date: '2020-03-04T00:00:00.000Z'}, numFmt: 'DD/MM/YYYY'},
              ],
            },
          ],
        });
        assert.strictEqual(
          model.sheets.S.cells.A1.numFmt,
          '0.00%',
          'the percentage code keeps its "." decimal, not a "," swap',
        );
        assert.strictEqual(
          model.sheets.S.cells.A2.numFmt,
          'DD/MM/YYYY',
          'the date code keeps its "/" separators, not "-"',
        );
      },
    },
  ],
} satisfies Case;
