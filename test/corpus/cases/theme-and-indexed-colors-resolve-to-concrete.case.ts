// Cluster: styles
//
// Real-world scenario: a caller inspects a cell's fill or font colour and gets `{theme: 4}` or
// `{indexed: 2}` — a reference into a table they have no access to, and no colour at all. OOXML lets a
// colour state itself three ways, and two of them are indirections: `theme="n"` points into the
// workbook theme's colour scheme, `indexed="n"` into a legacy 64-entry palette a workbook may also
// override. Either may carry a `tint` that lightens or darkens the result. Without resolution the
// model can report the colour of an explicit-RGB cell and nothing else — which, in files produced by
// Excel, is most cells.
//
// Two traps this pins down:
//
//  • `theme="n"` does NOT index the order the slots appear in the theme part. Index 0 is `lt1` and 1
//    is `dk1` — each dark/light pair swapped relative to the `<a:clrScheme>` child sequence the spec
//    tabulates. Reading the sequence order instead inverts text against background on every workbook.
//    Settled against Excel Desktop; see the recorded observation in
//    `test/corpus/fixtures/excel-oracle/theme-color-index-order.json`.
//  • `indexed="64"` is not a colour. It is the system-foreground sentinel, and it sits on the
//    background of essentially every solid fill Excel writes — resolving it to black would repaint
//    them all.
//
// Resolution is a *derived* view. The model keeps the encoding the file used, so a round-trip
// re-emits `theme="4" tint="0.4"` rather than a literal ARGB; resolving into the model would sever
// every cell's link to the theme, so recolouring the workbook would stop working.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'theme-and-indexed-colors-resolve-to-concrete/branded-colors.xlsx';
const CELLS = ['A1', 'B1', 'C1', 'D1', 'E1'];

export default {
  id: 'theme-and-indexed-colors-resolve-to-concrete',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'A colour stated as a theme slot or a palette index resolves to a concrete ARGB the caller can ' +
    'render — through the workbook’s own theme and its own custom palette, with any tint applied — ' +
    'while the model keeps the original reference so a round-trip re-emits it unchanged.',

  behavior: [
    {
      name: 'a theme-backed fill resolves through the workbook’s own colour scheme',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.fixtureColorResolution(FIXTURE, CELLS);
        assert.deepStrictEqual(
          cells.A1.fill,
          {theme: 4},
          'precondition: the cell states its fill as a theme slot and nothing else',
        );
        // accent1 in this file's theme, not the Office default 4472C4 — so a resolver that ignored
        // the workbook's theme could not pass by luck.
        assert.strictEqual(cells.A1.fillResolved, 'FFBB2649');
      },
    },
    {
      name: 'theme index 1 is dk1, so default body text resolves dark, not light',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells, themeColors} = await api.fixtureColorResolution(FIXTURE, CELLS);
        assert.strictEqual(themeColors.dk1, '1A1A1A', 'precondition: an off-black dk1');
        assert.strictEqual(themeColors.lt1, 'FAFAFA', 'precondition: an off-white lt1');
        assert.deepStrictEqual(cells.A1.font, {theme: 1});
        // Reading the clrScheme child order instead would give FFFAFAFA — white text on a coloured
        // fill, on every workbook ever written.
        assert.strictEqual(cells.A1.fontResolved, 'FF1A1A1A');
      },
    },
    {
      name: 'a tint is applied on top of the resolved theme colour',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.fixtureColorResolution(FIXTURE, CELLS);
        assert.strictEqual(cells.B1.fill.theme, 4, 'precondition: same slot as the untinted cell');
        assert.ok(cells.B1.fill.tint > 0, 'precondition: a lightening tint');
        assert.notStrictEqual(
          cells.B1.fillResolved,
          cells.A1.fillResolved,
          'the tinted cell must not resolve to the untinted colour',
        );
        // Lightening raises every channel; the check is on the effect, not on one Excel build's
        // rounding of the luminance shift.
        const channels = (argb: CorpusApi) =>
          [2, 4, 6].map((at: number) => Number.parseInt(argb.slice(at, at + 2), 16));
        const base = channels(cells.A1.fillResolved);
        const tinted = channels(cells.B1.fillResolved);
        assert.ok(
          tinted.every((value: number, i: number) => value > (base[i] as number)),
          `a positive tint lightens every channel (base ${cells.A1.fillResolved}, tinted ${cells.B1.fillResolved})`,
        );
      },
    },
    {
      name: 'the workbook’s custom palette wins over the built-in one',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.fixtureColorResolution(FIXTURE, CELLS);
        assert.deepStrictEqual(cells.C1.fill, {indexed: 2});
        // Slot 2 is pure red (00FF0000) in the built-in palette; this workbook overrode it.
        assert.strictEqual(cells.C1.fillResolved, 'FF123456');
        // A slot the workbook left alone still resolves — the override is per entry, not a reset.
        assert.deepStrictEqual(cells.D1.fill, {indexed: 10});
        assert.strictEqual(cells.D1.fillResolved, 'FFFF0000');
      },
    },
    {
      name: 'the system foreground index resolves to nothing rather than to black',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {cells} = await api.fixtureColorResolution(FIXTURE, CELLS);
        assert.deepStrictEqual(cells.E1.fill, {indexed: 64});
        assert.strictEqual(
          cells.E1.fillResolved,
          null,
          'indexed 64 names the system foreground, which has no fixed value',
        );
      },
    },
    {
      name: 'the built-in palette backs a workbook that overrides nothing',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        // Entries 0-7 duplicate 8-15 in the legacy palette; both must be known.
        assert.strictEqual(await api.resolveColorOnEmptyWorkbook({indexed: 2}), 'FFFF0000');
        assert.strictEqual(await api.resolveColorOnEmptyWorkbook({indexed: 10}), 'FFFF0000');
        assert.strictEqual(await api.resolveColorOnEmptyWorkbook({indexed: 22}), 'FFC0C0C0');
        assert.strictEqual(await api.resolveColorOnEmptyWorkbook({indexed: 65}), null);
        // A workbook with no theme of its own resolves against the Office default scheme.
        assert.strictEqual(await api.resolveColorOnEmptyWorkbook({theme: 4}), 'FF4472C4');
        assert.strictEqual(await api.resolveColorOnEmptyWorkbook({theme: 1}), 'FF000000');
      },
    },
    {
      name: 'resolving does not rewrite the model’s stored encoding',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        // The reference must survive being resolved: the writer re-emits what the model holds, so a
        // resolver that wrote back would turn every themed cell into a literal ARGB and break the
        // file's link to its theme.
        const {cells} = await api.fixtureColorResolution(FIXTURE, CELLS);
        assert.strictEqual(
          cells.A1.fill.argb,
          undefined,
          'the theme reference is still a reference',
        );
        assert.strictEqual(cells.C1.fill.argb, undefined, 'the palette index is still an index');
      },
    },
  ],
} satisfies Case;
