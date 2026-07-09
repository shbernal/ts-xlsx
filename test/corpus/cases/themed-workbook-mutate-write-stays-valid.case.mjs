// Cluster: styles
//
// Real-world scenario: a user opens a workbook that carries a theme (fonts referencing colors
// by theme index, the usual state of any Excel-authored file), edits a few cells — e.g. sets
// solid fills to highlight differences — and writes it back to a buffer. A recurring report
// claims the result is corrupt and only writing succeeds after stripping the theme first. The
// durable requirement: reading a themed workbook, mutating cells, and writing it back must
// produce a valid, re-readable package with the theme intact and the edits applied. Clearing
// the theme must never be a precondition for a non-corrupt write. (The original corruption in
// that report traced to concurrent read/write of the same file — a caller error, not a library
// defect — so the library path itself is expected to be sound; this case locks that in.)

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// A theme-color font (color by theme index) plus a solid fill applied to a different cell —
// exactly the "themed workbook with a highlight edit" shape.
const SPEC = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: 'header', font: {name: 'Calibri', color: {theme: 1}}},
        {ref: 'B2', value: 'edited', fill: {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FFFFFF00'}}},
      ],
    },
  ],
};

export default {
  id: 'themed-workbook-mutate-write-stays-valid',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 745},
  cluster: 'styles',
  description:
    'A themed workbook — a theme-color font plus a solid highlight fill — writes back to a ' +
    'valid, re-readable package with the theme part intact and the edited fill preserved; ' +
    'stripping the theme is never required to avoid corruption.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'the edited fill survives the write→read round-trip',
      baseline: 'pass',
      async expect(api, assert) {
        const model = await api.roundtripWorkbook(SPEC);
        const fill = model.sheets.S.cells.B2.fill;
        assert.strictEqual(fill && fill.pattern, 'solid', 'the solid fill is preserved');
        assert.strictEqual(fill.fgColor.argb, 'FFFFFF00', 'the fill color survives');
      },
    },
    {
      name: 'a theme part backs the theme-color font (no unresolved theme reference)',
      baseline: 'pass',
      async expect(api, assert) {
        const {styles} = await api.inspectPackage(SPEC);
        assert.ok(styles.hasThemePart, 'the package ships a theme part');
        assert.ok(styles.themeColorResolvable, 'every theme-color reference is backed by a theme part');
      },
    },
    {
      name: 'the written package is a structurally valid, self-consistent OOXML zip',
      baseline: 'pass',
      async expect(api, assert) {
        const {consistency, sheets} = await api.inspectPackage(SPEC);
        assert.ok(consistency.declaredConsistent, 'every worksheet part is declared and related consistently');
        assert.ok(sheets.S.xmlWellFormed, 'the worksheet XML is well-formed');
      },
    },
  ],
};
