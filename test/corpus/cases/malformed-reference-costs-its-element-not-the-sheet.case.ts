// Cluster: security
//
// Real-world scenario: a `.xlsx` this library did not write carries a reference no cell, area or
// region can have -- `<c r="A0">`, `<c r="ZZZZ1">`, a `sqref` off the grid, a `ref` that is not a
// reference at all. Every foreign scalar has a tolerant reader (a bad `count` falls back, an
// unrecognised enum token is dropped), but a reference had none: it went through the caller-facing
// decoders, so one corrupt attribute aborted the entire read with a native RangeError or
// SyntaxError -- outside the XlsxError taxonomy a caller's single `catch` clause is supposed to
// answer, and from a parser path that faces untrusted input.
//
// The rule this locks: a malformed reference costs the element that carried it and nothing else.
// The cell is skipped, the validation, hyperlink, table, filter or note is dropped, and every other
// feature on the sheet is read exactly as it would have been.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'malformed-reference-costs-its-element-not-the-sheet',
  provenance: {source: 'hostile-input-probe'},
  cluster: 'security',
  description:
    'A reference naming no cell, area or region that can exist is dropped along with the element ' +
    'carrying it: the read still succeeds, the rest of the sheet is intact, and no native ' +
    'RangeError/SyntaxError escapes the reader.',

  behavior: [
    {
      name: 'no malformed reference aborts the read',
      expect(api: CorpusApi, assert: Assert) {
        for (const row of api.malformedReferenceReport()) {
          assert.equal(
            row.threw,
            false,
            `${row.mutation}: the read survives it (${row.errorName})`,
          );
          assert.equal(row.keptSiblingCell, 'keep', `${row.mutation}: A1 is still readable`);
          assert.equal(row.sheetsRead, 1, `${row.mutation}: the sheet is still in the workbook`);
        }
      },
    },
    {
      name: 'a cell whose r names no possible position is skipped, its siblings are not',
      expect(api: CorpusApi, assert: Assert) {
        const rows = api
          .malformedReferenceReport()
          .filter((row) => row.mutation.startsWith('cell r '));
        assert.equal(rows.length, 3, 'off the grid on each axis, and unparseable');
        for (const row of rows) {
          assert.equal(row.model?.targetCell, null, `${row.mutation}: the bad cell is absent`);
          assert.equal(row.model?.validations, 1, `${row.mutation}: the validation survives`);
          assert.equal(row.model?.tables, 1, `${row.mutation}: the table survives`);
          assert.equal(row.model?.note, 'a note', `${row.mutation}: the note survives`);
        }
      },
    },
    {
      name: 'each feature-level reference is dropped alone',
      expect(api: CorpusApi, assert: Assert) {
        const byMutation = new Map(
          api.malformedReferenceReport().map((row) => [row.mutation, row.model]),
        );
        const validation = byMutation.get('dataValidation sqref off the grid');
        assert.equal(validation?.validations, 0, 'the validation covering nothing is dropped');
        assert.equal(validation?.targetCell, 'target', 'and the cells it named are untouched');

        const filter = byMutation.get('autoFilter ref off the grid');
        assert.equal(filter?.autoFilter, null, 'the filter is dropped');
        assert.equal(filter?.tables, 1, 'and the table beside it is not');

        const link = byMutation.get('hyperlink ref unparseable');
        assert.equal(link?.hyperlink, null, 'the hyperlink is dropped');
        assert.equal(link?.note, 'a note', 'and the note is not');

        const table = byMutation.get('table ref unparseable');
        assert.equal(table?.tables, 0, 'a table whose anchor is unreadable is dropped whole');
        assert.equal(table?.validations, 1, 'and the validation beside it is not');

        const note = byMutation.get('comment ref off the grid');
        assert.equal(note?.note, null, 'the note is dropped');
        assert.equal(note?.targetCell, 'target', 'and the cells beside it are not');
      },
    },
  ],
} satisfies Case;
