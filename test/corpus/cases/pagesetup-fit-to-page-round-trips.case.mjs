// Cluster: page-setup
//
// Real-world scenario: a user configures a worksheet's print scaling via fit-to-page — fit all
// columns onto one page wide, or all rows onto one page tall — instead of a fixed zoom. The
// fit-to-page intent (the flag plus the fitToWidth/fitToHeight page counts and any scale) must
// be emitted onto the sheet properties and pageSetup consistently and survive a round-trip, so
// the reopened workbook prints with the same fit behavior the author requested.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const FIT_WIDTH = {
  sheets: [
    {
      name: 'S',
      cells: [{ref: 'A1', value: 1}],
      pageSetup: {fitToPage: true, fitToWidth: 1, fitToHeight: 0, scale: 80},
    },
  ],
};
const FIT_HEIGHT = {
  sheets: [
    {
      name: 'S',
      cells: [{ref: 'A1', value: 1}],
      pageSetup: {fitToPage: true, fitToWidth: 0, fitToHeight: 1},
    },
  ],
};

export default {
  id: 'pagesetup-fit-to-page-round-trips',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 388},
  cluster: 'page-setup',
  description:
    'A worksheet’s fit-to-page print scaling — the fitToPage flag together with the ' +
    'fitToWidth/fitToHeight page counts (and scale) — survives a write→read round-trip so the ' +
    'reopened sheet prints with the requested fit behavior.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'fit-all-columns-on-one-page (fitToWidth=1, fitToHeight=0) round-trips',
      baseline: 'pass',
      async expect(api, assert) {
        const ps = (await api.roundtripWorkbook(FIT_WIDTH)).sheets.S.pageSetup;
        assert.strictEqual(ps.fitToPage, true, 'the fit-to-page flag is set');
        assert.strictEqual(ps.fitToWidth, 1, 'fitToWidth survives');
        assert.strictEqual(ps.fitToHeight, 0, 'fitToHeight survives');
        assert.strictEqual(ps.scale, 80, 'the author scale survives');
      },
    },
    {
      name: 'fit-all-rows-on-one-page (fitToHeight=1, fitToWidth=0) round-trips',
      baseline: 'pass',
      async expect(api, assert) {
        const ps = (await api.roundtripWorkbook(FIT_HEIGHT)).sheets.S.pageSetup;
        assert.strictEqual(ps.fitToPage, true, 'the fit-to-page flag is set');
        assert.strictEqual(ps.fitToWidth, 0, 'fitToWidth survives');
        assert.strictEqual(ps.fitToHeight, 1, 'fitToHeight survives');
      },
    },
  ],
};
