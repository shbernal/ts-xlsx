// Cluster: xlsx-io
//
// Real-world scenario: a user sets a couple of print margins on a worksheet (say
// just left and right) and saves. OOXML's <pageMargins> element requires all six
// attributes — left, right, top, bottom, header, footer — to be present; a partial
// element is invalid and can make Excel repair the file. A faithful writer must
// emit all six (filling the untouched ones with valid finite defaults) while
// preserving the values the user did set.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => void }} Behavior */

const SPEC = {sheets: [{name: 'Sheet1', cells: [{ref: 'A1', value: 1}], pageMargins: {left: 0.1, right: 0.1}}]};
const SIX = ['left', 'right', 'top', 'bottom', 'header', 'footer'];

export default {
  id: 'page-margins-must-be-complete',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 631},
  cluster: 'xlsx-io',
  description:
    'When any page margin is set, the written <pageMargins> element must contain ' +
    'all six margin attributes with valid numeric values; explicitly-set values ' +
    'are preserved.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'setting a subset of margins still emits all six pageMargins attributes',
      baseline: 'pass',
      async expect(api, assert) {
        const {sheets} = await api.inspectPackage(SPEC);
        const present = sheets.Sheet1.pageMargins.present;
        for (const attr of SIX) {
          assert.ok(present.includes(attr), `pageMargins missing "${attr}"; got [${present}]`);
        }
      },
    },
    {
      name: 'margins not explicitly set are filled with valid finite numbers',
      baseline: 'pass',
      async expect(api, assert) {
        const {values} = (await api.inspectPackage(SPEC)).sheets.Sheet1.pageMargins;
        for (const attr of SIX) {
          const n = Number(values[attr]);
          assert.ok(Number.isFinite(n), `margin "${attr}" is not a finite number: ${values[attr]}`);
        }
      },
    },
    {
      name: 'explicitly-set margin values are preserved, not clobbered by defaults',
      baseline: 'pass',
      async expect(api, assert) {
        const {values} = (await api.inspectPackage(SPEC)).sheets.Sheet1.pageMargins;
        assert.strictEqual(Number(values.left), 0.1);
        assert.strictEqual(Number(values.right), 0.1);
      },
    },
  ],
};
