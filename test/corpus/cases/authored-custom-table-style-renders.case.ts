// Cluster: styles
//
// Real-world scenario: a team wants their tables in their own look — branded header row, their own
// stripe colour — reused across every table in every workbook they generate. Excel calls this a table
// style, and a workbook can define its own beside the built-in gallery: a `<tableStyle>` in styles.xml
// whose elements each name a region of the table and the differential formatting to lay over it. A
// table then reaches it by name through `tableStyleInfo/@name`, exactly as it would name
// "TableStyleMedium2".
//
// This is a *cross-part* claim, and that is what makes it hard to check. The table part names a
// style, the styles part defines it, and the dxf table backs each of its elements — every one of
// those parts can be individually schema-valid while the whole says nothing to Excel and the table
// renders unstyled. So the assertions here resolve the references rather than counting them, and the
// question the corpus structurally cannot answer — whether Excel *renders* the style or merely opens
// the file without complaint — was put to Excel Desktop directly and recorded in
// `test/corpus/fixtures/excel-oracle/authored-table-style-renders.json`: it registers the style in
// the workbook's own gallery, paints the header row from it (bold in DisplayFormat while the cells
// themselves carry no bold), and honours a stripe's `size` across both data rows.

import type {Assert, Case, CorpusApi} from '../case.ts';

const HARBOUR = {
  name: 'Harbour',
  elements: {
    wholeTable: {border: {top: {style: 'thin'}, bottom: {style: 'thin'}}},
    headerRow: {
      font: {bold: true, color: {argb: 'FFFFFFFF'}},
      fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFBB2649'}},
    },
    firstRowStripe: {
      fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFF6E7EB'}},
      size: 2,
    },
  },
};

// A fixture that already defines a table style called "Harbour Table", referenced by its own table.
const PRESERVED = 'custom-table-style-definition-survives-roundtrip/branded-table-style.xlsx';

export default {
  id: 'authored-custom-table-style-renders',
  provenance: {
    source: 'excel-desktop-verification',
    ref: 'test/corpus/fixtures/excel-oracle/authored-table-style-renders.json',
  },
  cluster: 'styles',
  description:
    'A custom table style authored on the workbook is emitted as a real `<tableStyle>` whose every ' +
    'element resolves to a differential style, and a table naming it resolves to that definition — ' +
    'the cross-part wiring Excel needs to actually paint the table.',

  behavior: [
    {
      name: 'an authored style is emitted and the table naming it resolves',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.authorTableStyleReport({
          styles: [HARBOUR],
          tableStyle: {name: 'Harbour', showRowStripes: true},
        });
        assert.deepStrictEqual(report.definitions, ['Harbour']);
        assert.strictEqual(report.nameOnTable, 'Harbour', 'the table asks for the authored style');
        assert.strictEqual(report.resolves, true, 'and the stylesheet defines it');
        // The container's count must agree with the definitions it holds, or Excel treats the
        // stylesheet as corrupt.
        assert.strictEqual(report.declaredCount, 1);
      },
    },
    {
      name: 'every element resolves to a differential style that carries its formatting',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.authorTableStyleReport({
          styles: [HARBOUR],
          tableStyle: {name: 'Harbour', showRowStripes: true},
        });
        assert.deepStrictEqual(
          report.elements.map((element) => element.type),
          ['wholeTable', 'headerRow', 'firstRowStripe'],
          'the elements are emitted in the order the schema fixes, not authoring order',
        );
        // A dxfId that resolves to nothing is a style that paints nothing — the failure mode this
        // whole feature has to avoid, and one no schema check catches.
        assert.ok(
          report.elementDxfs.every((dxf) => dxf !== null),
          'every element’s dxfId lands on a real <dxf>',
        );
        const header =
          report.elementDxfs[report.elements.findIndex((element) => element.type === 'headerRow')];
        assert.match(header!, /<b\/>/, 'the header row’s dxf carries the bold it was given');
        assert.match(header!, /FFBB2649/, 'and its fill colour');
      },
    },
    {
      name: 'a stripe’s band width is written, and only a stripe may carry one',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.authorTableStyleReport({
          styles: [HARBOUR],
          tableStyle: {name: 'Harbour', showRowStripes: true},
        });
        const stripe = report.elements.find((element) => element.type === 'firstRowStripe');
        assert.strictEqual(stripe!.size, 2, 'the authored band width reaches the file');
        // `size` defaults to 1 and means nothing outside the four stripe types, so a non-stripe
        // element must not carry one at all.
        const wholeTable = report.elements.find((element) => element.type === 'wholeTable');
        assert.strictEqual(wholeTable!.size, null);
      },
    },
    {
      name: 'elements painted alike share one differential style',
      async expect(api: CorpusApi, assert: Assert) {
        // The dxf table is shared and interned, so a style whose header row and total row look the
        // same costs one entry, not two — the same interning a conditional-formatting rule uses.
        const fill = {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFBB2649'}};
        const report = await api.authorTableStyleReport({
          styles: [{name: 'Twin', elements: {headerRow: {fill}, totalRow: {fill}}}],
          tableStyle: {name: 'Twin'},
        });
        assert.strictEqual(report.dxfCount, 1, 'one <dxf> backs both elements');
        assert.strictEqual(report.elements[0]!.dxfId, report.elements[1]!.dxfId);
      },
    },
    {
      name: 'authoring a name a source file already defined overrides it, rather than duplicating it',
      async expect(api: CorpusApi, assert: Assert) {
        // Two `<tableStyle>` elements sharing a name leave the table's reference ambiguous — a
        // consumer resolves whichever it indexes first — so a second definition is never the answer.
        const report = await api.authorTableStyleReport({
          fixture: PRESERVED,
          styles: [{name: 'Harbour Table', elements: {headerRow: {font: {italic: true}}}}],
        });
        assert.deepStrictEqual(report.definitions, ['Harbour Table']);
        assert.strictEqual(report.declaredCount, 1);
        assert.strictEqual(report.nameOnTable, 'Harbour Table');
        assert.strictEqual(report.resolves, true);
        assert.match(
          report.elementDxfs[0]!,
          /<i\/>/,
          'the emitted definition is the authored one, not the preserved one',
        );
      },
    },
    {
      name: 'a style that could never paint anything is refused at the call that made it',
      async expect(api: CorpusApi, assert: Assert) {
        // Both of these produce a file Excel opens without complaint and then quietly does nothing
        // with — the class of bug that never gets found, so it is refused at the setter instead.
        assert.match(
          await api.authorInvalidTableStyle({name: '', elements: {}})!,
          /needs a name/,
          'an unnamed style is unreachable: a table references its style by name',
        );
        assert.match(
          await api.authorInvalidTableStyle({name: 'X', elements: {headerRow: {size: 2}}})!,
          /cannot carry a size/,
          'band width applies only to the four stripe types; Excel ignores it elsewhere',
        );
        assert.match(
          await api.authorInvalidTableStyle({name: 'X', elements: {firstRowStripe: {size: 0}}})!,
          /positive integer/,
        );
        // A stripe with a real band width is of course fine.
        assert.strictEqual(
          await api.authorInvalidTableStyle({name: 'X', elements: {firstRowStripe: {size: 2}}}),
          null,
        );
      },
    },
  ],
} satisfies Case;
