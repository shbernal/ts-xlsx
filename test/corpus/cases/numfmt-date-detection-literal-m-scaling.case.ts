// Cluster: types
//
// Real-world scenario: a cell stores a plain large integer (here 84311311) and applies a custom
// number format used to display big numbers compactly — a comma-scaling format that divides by
// millions and appends a literal unit letter, written `0.0,, \m` (the `m` is backslash-escaped as
// a literal character, and the trailing commas are scaling, not time). The cell is a number and
// must read back as a number. The bug this guards against: the reader's date-detection heuristic
// scans the format code for date/time letters (d/m/y/h/s) and, seeing the escaped literal `m`
// and/or the comma scaling, wrongly classifies the format as a date and coerces the serial into a
// JavaScript Date. Correct behavior honors OOXML escaping/quoting — escaped and quoted characters
// are never format tokens — so a format counts as a date only when it holds genuine, unescaped
// date/time tokens.

import type {Assert, Case, CorpusApi} from '../case.ts';

const FIXTURE = 'numfmt-date-detection-literal-m-scaling/source.xlsx';

export default {
  id: 'numfmt-date-detection-literal-m-scaling',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    'A cell holding a plain integer with the compact-millions number format "0.0,, \\m" (an ' +
    'escaped literal "m", plus comma scaling) reads back as a number, not a Date — escaped/quoted ' +
    'characters and comma-scaling groups are never treated as date/time tokens.',

  behavior: [
    {
      name: 'a plain integer with the escaped-literal-m scaling format reads back as a number, not a Date',
      baseline: 'pass',
      async expect(api: CorpusApi, assert: Assert) {
        const cells = await api.readFixtureCells(FIXTURE, ['A84']);
        assert.strictEqual(
          cells.A84.type,
          'number',
          `an escaped literal "m" must not make the format a date; A84 must be a number, got ${JSON.stringify(cells.A84)}`,
        );
      },
    },
  ],
} satisfies Case;
