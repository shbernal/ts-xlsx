// Cluster: types
//
// Real-world scenario: a user assigns a string cell value containing a newline to create an
// in-cell line break ("My Text with\na line break"). Results were reported as inconsistent
// depending on whether the source used a bare line-feed (\n) or a carriage-return+line-feed
// pair (\r\n), with breaks sometimes collapsing to a space or being dropped, and readers
// disagreeing. A newline in a cell value must be preserved as a genuine in-cell line break,
// survive a round-trip, and be normalized so the stored result is the same regardless of the
// caller's newline convention. wrapText alignment governs only default *rendering* of the
// break — it must not change whether the break is *stored*.

import type {Assert, Case, CorpusApi} from '../case.ts';

const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'line1\nline2'},
        {ref: 'A2', value: 'crlf1\r\ncrlf2'},
        {ref: 'A3', value: 'a\n\nb'},
        {ref: 'A4', value: 'wrap\nped', alignment: {wrapText: true}},
      ],
    },
  ],
};

export default {
  id: 'cell-value-newline-line-break-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 369},
  cluster: 'types',
  description:
    'Newlines in a string cell value are preserved as in-cell line breaks across a ' +
    'write→read round-trip: a bare \\n survives, a \\r\\n normalizes to the same \\n form, ' +
    'consecutive newlines are all kept, and wrapText does not change what is stored.',

  behavior: [
    {
      name: 'a bare line-feed survives the round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(
          model.sheets.S.cells.A1.value,
          'line1\nline2',
          'the \\n break is preserved',
        );
      },
    },
    {
      name: 'a carriage-return+line-feed normalizes to the same stored break',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(
          model.sheets.S.cells.A2.value,
          'crlf1\ncrlf2',
          '\\r\\n is normalized to \\n so all readers agree',
        );
      },
    },
    {
      name: 'consecutive newlines are all preserved, not merged',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(
          model.sheets.S.cells.A3.value,
          'a\n\nb',
          'a double break is kept intact',
        );
      },
    },
    {
      name: 'wrapText does not change whether the newline is stored',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const model = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(
          model.sheets.S.cells.A4.value,
          'wrap\nped',
          'the break is stored regardless of wrapText',
        );
      },
    },
  ],
} satisfies Case;
