// Cluster: styles
//
// Real-world scenario: a caller sets explicit widths on several columns, one of which is given a width
// that happens to coincide with the format's conventional default column width (numerically 9). On
// write and re-read, the column explicitly set to 9 collapses to the default — no <col> element is
// emitted for it, and it reads back with no width — while every other explicit width (8, 10, …) is
// preserved with customWidth="1". The writer treats "width equals the magic default" as "same as
// default, skip", discarding a width the caller set on purpose. An explicitly-set width must be
// preserved regardless of whether its value coincides with the default.

import type {Assert, Case, CorpusApi} from '../case.ts';

export default {
  id: 'explicit-column-width-equal-to-default-magic-value-survives',
  provenance: {source: 'upstream-issue'},
  cluster: 'styles',
  description:
    'An explicitly-set column width that equals the conventional default (9) is preserved on write ' +
    'and re-read — it emits an explicit <col> width and reads back as 9 — rather than being silently ' +
    'dropped as "same as default".',

  behavior: [
    {
      name: 'a column explicitly set to a non-default width round-trips (control)',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {readBack, emitted} = await api.columnWidthDefaultCollisionReport([8, 9, 10]);
        assert.strictEqual(readBack.c1, 8, 'width 8 reads back');
        assert.ok(emitted.c1, 'width 8 emits an explicit <col>');
        assert.strictEqual(readBack.c3, 10, 'width 10 reads back');
        assert.ok(emitted.c3, 'width 10 emits an explicit <col>');
      },
    },
    {
      name: 'a column explicitly set to the default magic width (9) emits an explicit <col>',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {emitted} = await api.columnWidthDefaultCollisionReport([8, 9, 10]);
        assert.strictEqual(
          emitted.c2,
          true,
          'an explicit width of 9 must emit a <col>, not defer to the default',
        );
      },
    },
    {
      name: 'a column explicitly set to 9 reads back as 9, not undefined',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const {readBack} = await api.columnWidthDefaultCollisionReport([8, 9, 10]);
        assert.strictEqual(readBack.c2, 9, 'an explicit width of 9 must survive the round-trip');
      },
    },
  ],
} satisfies Case;
