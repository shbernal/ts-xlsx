// Cluster: streaming
//
// Real-world scenario: streaming a workbook to disk with shared strings enabled, several
// cells each hold a *distinct* rich-text value (different runs, different formatting). Each
// cell must serialize to its own shared-string entry and read back as its own rich text.
// With shared strings enabled today, every rich-text cell collapses onto the first one —
// all cells end up displaying the first cell's text and formatting. With shared strings
// disabled the text is distinct and run formatting is retained; that path is the control.

import type {Assert, Case, CorpusApi} from '../case.ts';

const RT_ALPHA = {richText: [{text: 'Alpha', bold: true}, {text: 'One'}]};
const RT_BETA = {richText: [{text: 'Beta', italic: true}, {text: 'Two'}]};

const text = (cell: CorpusApi) =>
  cell?.richText ? cell.richText.map((r: CorpusApi) => r.text).join('') : cell;

export default {
  id: 'streaming-write-richtext-shared-strings-distinct',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2267},
  cluster: 'streaming',
  description:
    'Streaming a workbook with shared strings enabled, two cells with distinct rich-text ' +
    'values read back as their own distinct rich text — they must not be deduplicated onto ' +
    'the first cell’s value.',

  behavior: [
    {
      name: 'with shared strings enabled, distinct rich-text cells stay distinct',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, cells} = await api.streamWriteSheet({
          useSharedStrings: true,
          ops: [
            {op: 'addRow', value: [RT_ALPHA]},
            {op: 'addRow', value: [RT_BETA]},
          ],
          read: ['A1', 'A2'],
        });
        assert.ok(ok, 'the streaming write must succeed');
        assert.strictEqual(text(cells.A1), 'AlphaOne', 'first cell keeps its own text');
        assert.strictEqual(
          text(cells.A2),
          'BetaTwo',
          'second cell must keep its own rich text, not collapse onto the first',
        );
      },
    },
    {
      name: 'with shared strings disabled, distinct rich-text cells stay distinct (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {ok, cells} = await api.streamWriteSheet({
          useSharedStrings: false,
          ops: [
            {op: 'addRow', value: [RT_ALPHA]},
            {op: 'addRow', value: [RT_BETA]},
          ],
          read: ['A1', 'A2'],
        });
        assert.ok(ok, 'the streaming write must succeed');
        assert.strictEqual(text(cells.A1), 'AlphaOne', 'first cell text');
        assert.strictEqual(text(cells.A2), 'BetaTwo', 'second cell text is distinct');
      },
    },
    {
      name: 'with shared strings disabled, rich-text run formatting is retained (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.streamWriteSheet({
          useSharedStrings: false,
          ops: [{op: 'addRow', value: [RT_ALPHA]}],
          read: ['A1'],
        });
        const runs = cells.A1?.richText;
        assert.ok(runs, 'the cell reads back as rich text');
        assert.strictEqual(runs[0].text, 'Alpha', 'first run text');
        assert.strictEqual(runs[0].bold, true, 'first run stays bold');
      },
    },
  ],
} satisfies Case;
