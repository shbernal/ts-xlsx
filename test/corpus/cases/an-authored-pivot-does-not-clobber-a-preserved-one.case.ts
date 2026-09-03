// Cluster: pivot
//
// Real-world scenario: open a workbook that already has a pivot table, add another one, save. The
// existing pivot is content this library does not model, so it rides through verbatim on its original
// path; the new one is generated, numbered globally from 1. Both wanted `xl/pivotTables/pivotTable1.xml`.
//
// The preserved parts are emitted last, so they won. The authored pivot's part was overwritten by the
// old pivot's bytes, its sheet relationship then pointed at data it was never built from, and the
// content types declared the same PartName three times over, which violates OPC M2.5 and makes Excel
// offer to repair the file. Nothing errored: the package simply came out wrong.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'pivot-table-parts-survive-roundtrip/sample.xlsx';

export default {
  id: 'an-authored-pivot-does-not-clobber-a-preserved-one',
  provenance: {source: 'package-assembly-audit'},
  cluster: 'pivot',
  description:
    'Authoring a pivot table onto a workbook that already carries preserved pivot parts emits both: ' +
    'the preserved parts are renumbered past the generated ones, every PartName is declared once, ' +
    'and the authored pivot keeps its own part rather than being overwritten.',

  behavior: [
    {
      name: 'no part name is declared twice in the content types',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotPreservedAndAuthoredCoexist(FIXTURE);
        assert.deepEqual(report.duplicateOverrides, []);
      },
    },
    {
      name: 'the preserved pivots and the authored one all have a part of their own',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotPreservedAndAuthoredCoexist(FIXTURE);
        assert.equal(
          report.pivotTableParts.length,
          report.sourcePivotTableParts.length + 1,
          `expected one more pivot part than the ${report.sourcePivotTableParts.length} the source ` +
            `carried, got ${report.pivotTableParts.join(', ')}`,
        );
      },
    },
    {
      name: 'each pivot has its own cache definition and records part',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotPreservedAndAuthoredCoexist(FIXTURE);
        assert.equal(
          new Set(report.cacheDefinitionParts).size,
          report.cacheDefinitionParts.length,
          'no cache definition path is claimed twice',
        );
        assert.equal(new Set(report.cacheRecordsParts).size, report.cacheRecordsParts.length);
      },
    },
    {
      name: 'the authored pivot is what the generated part holds, not the preserved one',
      expect(api: CorpusApi, assert: Assert) {
        const report = api.pivotPreservedAndAuthoredCoexist(FIXTURE);
        assert.equal(report.authoredPartChanged, true);
      },
    },
  ],
} satisfies Case;
