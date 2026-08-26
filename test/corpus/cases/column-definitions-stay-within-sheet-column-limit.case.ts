// Cluster: address-decoding
//
// Real-world scenario: a worksheet's column definitions (widths/styles) address a range that
// runs past the spreadsheet format's hard limit of 16384 columns (the last legal column is
// XFD = 16384). Emitting a <col> element whose `max` exceeds 16384 produces a file strict
// spreadsheet applications reject as corrupt. A written worksheet must never carry such a
// definition: the last legal column serializes, and an index past it never reaches the file --
// refused when it is authored, so the corruption has no path into the package at all.

import type {Assert, Case, CorpusApi} from '../case.ts';

const MAX_COLUMNS = 16384;
const LAST_LEGAL = {sheets: [{name: 'S', columns: [{index: MAX_COLUMNS, width: 10}]}]};
const PAST_THE_LIMIT = {sheets: [{name: 'S', columns: [{index: MAX_COLUMNS + 1, width: 10}]}]};

export default {
  id: 'column-definitions-stay-within-sheet-column-limit',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1366},
  cluster: 'address-decoding',
  description:
    'A written worksheet never emits a <col> definition whose column index exceeds the ' +
    'spreadsheet limit of 16384 columns; a column addressed beyond the limit is refused at ' +
    'authoring time rather than serialized into a range strict Excel treats as corrupt.',

  behavior: [
    {
      name: 'the last legal column (XFD) serializes, and no emitted group runs past it',
      async expect(api: CorpusApi, assert: Assert) {
        const {sheets} = await api.inspectPackage(LAST_LEGAL);
        assert.ok(
          sheets.S!.maxColumnIndex <= MAX_COLUMNS,
          `no <col> max may exceed ${MAX_COLUMNS}; got groups ${JSON.stringify(sheets.S!.columnGroups)}`,
        );
      },
    },
    {
      name: 'a column addressed past XFD is refused, so no out-of-range group can be written',
      expect(api: CorpusApi, assert: Assert) {
        assert.throws(
          () => api.inspectPackage(PAST_THE_LIMIT),
          RangeError,
          `column ${MAX_COLUMNS + 1} is outside the grid and must not be accepted onto a sheet`,
        );
      },
    },
  ],
} satisfies Case;
