// Cluster: types
//
// Real-world scenario: an application builds a worksheet from mixed data where a logically-numeric
// column arrives partly as JS numbers and partly as JS strings — a quantity field holding both 15
// and "10", or an identifier column of zero-padded codes like "007". The library must preserve the
// caller's DECLARED type: a JS string is written as a text cell, a JS number as a numeric cell. It
// must NOT silently coerce a digit-only string into a number — that would destroy zero-padded codes,
// account numbers, and identifiers, and is exactly the corruption a "number stored as text" advisory
// in a spreadsheet application is warning about (correctly). Type fidelity through a write→read
// round-trip is the durable contract; the advisory is the honest consequence, not a bug.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const spec = {
  sheets: [
    {
      name: 'S',
      cells: [
        {ref: 'A1', value: '10'},
        {ref: 'A2', value: 15},
        {ref: 'A3', value: '007'},
      ],
    },
  ],
};

const readBack = async (api) => (await api.roundtripWorkbook(spec)).sheets.S.cells;

export default {
  id: 'numeric-looking-string-stays-text-cell',
  provenance: {source: 'upstream-issue'},
  cluster: 'types',
  description:
    'A cell assigned a digit-only JS string is persisted and read back as a string (text cell), and ' +
    'a cell assigned a JS number is persisted and read back as a number — the library preserves the ' +
    'declared type and never coerces a numeric-looking string into a numeric cell, so zero-padded ' +
    'codes survive intact.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'a digit-only string value ("10") round-trips as a string, not coerced to a number',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = await readBack(api);
        assert.strictEqual(
          typeof cells.A1.value,
          'string',
          `"10" must stay a string; got ${typeof cells.A1.value} ${JSON.stringify(cells.A1.value)}`,
        );
        assert.strictEqual(cells.A1.value, '10');
      },
    },
    {
      name: 'a numeric value (15) round-trips as a number',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = await readBack(api);
        assert.strictEqual(
          typeof cells.A2.value,
          'number',
          `15 must stay a number; got ${typeof cells.A2.value}`,
        );
        assert.strictEqual(cells.A2.value, 15);
      },
    },
    {
      name: 'a zero-padded code ("007") survives without losing its leading zeros to numeric coercion',
      baseline: 'pass',
      async expect(api, assert) {
        const cells = await readBack(api);
        assert.strictEqual(
          cells.A3.value,
          '007',
          `zero-padded code must not be coerced to 7; got ${JSON.stringify(cells.A3.value)}`,
        );
      },
    },
  ],
};
