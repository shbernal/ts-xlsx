// Cluster: address-decoding
//
// Real-world scenario: a worksheet's column definitions (widths/styles) address a range that
// runs past the spreadsheet format's hard limit of 16384 columns (the last legal column is
// XFD = 16384). Today setting a column at an index beyond the limit emits a <col> element whose
// `max` exceeds 16384, which strict spreadsheet applications reject as a corrupt file. A written
// worksheet must never emit a column definition whose min or max column index exceeds the sheet
// maximum — an out-of-range column group must be clamped or dropped, keeping the column table
// internally consistent.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const MAX_COLUMNS = 16384;
// A column at the last legal index (control) and one past it (the corruption trigger).
const SPEC = {sheets: [{name: 'S', columns: [{index: MAX_COLUMNS, width: 10}, {index: MAX_COLUMNS + 1, width: 10}]}]};

export default {
  id: 'column-definitions-stay-within-sheet-column-limit',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1366},
  cluster: 'address-decoding',
  description:
    'A written worksheet never emits a <col> definition whose column index exceeds the ' +
    'spreadsheet limit of 16384 columns; a column addressed beyond the limit is clamped or ' +
    'dropped rather than serialized into a range strict Excel treats as corrupt.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'no emitted column group runs past the 16384-column limit',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        assert.ok(
          sheets.S.maxColumnIndex <= MAX_COLUMNS,
          `no <col> max may exceed ${MAX_COLUMNS}; got groups ${JSON.stringify(sheets.S.columnGroups)}`
        );
      },
    },
  ],
};
