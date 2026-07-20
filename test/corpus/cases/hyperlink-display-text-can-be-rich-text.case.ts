// Cluster: types
//
// Real-world scenario: a hyperlink cell's display text is not always a plain string — it can
// carry character formatting (rich text runs), e.g. a partly-bold label over a single link
// target. A hyperlink value must therefore model its display text as either a plain string or
// rich text, and both forms must survive a round-trip with the target URL intact. A workbook
// mixing plain and rich hyperlink labels must preserve each kind.

import type {Assert, Case, CorpusApi} from '../case.ts';

const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', hyperlink: 'https://example.com', text: 'plain label'},
        {
          ref: 'A2',
          hyperlink: 'https://example.org',
          text: {richText: [{text: 'bold', font: {bold: true}}, {text: 'plain'}]},
        },
      ],
    },
  ],
};

export default {
  id: 'hyperlink-display-text-can-be-rich-text',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 1988},
  cluster: 'types',
  description:
    'A hyperlink cell’s display text may be a plain string or rich text; both survive a ' +
    'write→read round-trip with the link target intact, and a workbook mixing the two ' +
    'preserves each kind independently.',

  behavior: [
    {
      name: 'a plain-string hyperlink label round-trips with its target',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {A1} = (await api.roundtripWorkbook(SPEC)).sheets.S.cells;
        assert.strictEqual(A1.hyperlink, 'https://example.com', 'the target survives');
        assert.strictEqual(A1.text, 'plain label', 'the plain display text survives');
      },
    },
    {
      name: 'a rich-text hyperlink label round-trips as rich text with its runs',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {A2} = (await api.roundtripWorkbook(SPEC)).sheets.S.cells;
        assert.strictEqual(A2.hyperlink, 'https://example.org', 'the target survives');
        assert.ok(
          A2.text && Array.isArray(A2.text.richText),
          'the display text is rich text, not flattened to a string',
        );
        assert.strictEqual(A2.text.richText[0].text, 'bold', 'the first run text survives');
        assert.strictEqual(
          A2.text.richText[0].font.bold,
          true,
          'the first run keeps its bold formatting',
        );
        assert.strictEqual(A2.text.richText[1].text, 'plain', 'the second run text survives');
      },
    },
  ],
} satisfies Case;
