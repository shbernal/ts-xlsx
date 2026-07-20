// Cluster: xlsx-io
//
// Real-world scenario: a program sets document metadata on a workbook — author
// (creator), last editor, and the created/modified timestamps that populate
// docProps/core.xml — then writes the file. Opening and reading that file back
// must return the same metadata; it is what shows in Excel's "Info" pane and what
// document-management systems index on. Losing it on write/read is silent data loss.

import type {Assert, Case, CorpusApi} from '../case.ts';

const SPEC = {
  properties: {
    creator: 'Ada Lovelace',
    lastModifiedBy: 'Grace Hopper',
    created: '2020-01-02T03:04:05.000Z',
    modified: '2021-06-07T08:09:10.000Z',
  },
  sheets: [{name: 'Sheet1', cells: [{ref: 'A1', value: 'hi'}]}],
};

export default {
  id: 'workbook-core-properties-roundtrip',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 130},
  cluster: 'xlsx-io',
  description:
    'Workbook core properties (creator, lastModifiedBy, created, modified) must ' +
    'survive a write→read round-trip unchanged.',

  behavior: [
    {
      name: 'the author (creator) set before write is read back unchanged',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {properties} = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(properties.creator, 'Ada Lovelace');
      },
    },
    {
      name: 'the last-modified-by set before write is read back unchanged',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {properties} = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(properties.lastModifiedBy, 'Grace Hopper');
      },
    },
    {
      name: 'the created timestamp survives the round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {properties} = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(properties.created, '2020-01-02T03:04:05.000Z');
      },
    },
    {
      name: 'the modified timestamp survives the round-trip',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {properties} = await api.roundtripWorkbook(SPEC);
        assert.strictEqual(properties.modified, '2021-06-07T08:09:10.000Z');
      },
    },
  ],
} satisfies Case;
