// Cluster: dates
//
// Real-world scenario: a workbook declares which calendar its serials count from. `<workbookPr
// date1904="1"/>` selects the 1904 system, where serial 0 is 1904-01-01 and there is no phantom leap
// day; without it, serial 1 is 1900-01-01 and serial 60 is a February 29th that never happened. The
// two are 1462 days apart. Files carrying the flag are everywhere a Mac Excel ever touched a
// workbook, and Excel still offers the setting on every platform.
//
// A reader that ignores the declaration does not fail: it produces perfectly good dates, four years
// and a day off, with no error anywhere. Writing such a workbook back without the flag is the same
// bug pointed the other way, and worse, because the serials stay as they were and the *next* reader
// applies the other calendar to them: the file silently changes meaning.
//
// The fixtures are Excel Desktop's own, saved with the setting on and off, in both serialisations.
// Excel puts 2023-03-15 at serial 45000 in the 1900 workbook and 43538 in the 1904 one, so the pair
// is an independent oracle: the two files must read back to the same calendar date.

import type {Assert, Case, CorpusApi} from '../case.ts';

const DIR = 'workbook-date-system-1904';
const CELLS = ['A2', 'A4'];
const MARCH_15 = '2023-03-15T00:00:00.000Z';
const MARCH_16 = '2023-03-16T00:00:00.000Z';

export default {
  id: 'workbook-date-system-1904',
  provenance: {source: 'excel-desktop-verification'},
  cluster: 'dates',
  description:
    "A workbook's `date1904` declaration governs every serial in it, in both readers and both " +
    'serialisations, and survives a write: the same calendar date reads back from the 1904 file and ' +
    'its 1900 twin, and re-writing the 1904 file still declares the system its serials are in.',

  behavior: [
    {
      name: 'a 1904-system workbook reads its serials on the calendar dates Excel shows',
      async expect(api: CorpusApi, assert: Assert) {
        const report = await api.dateSystemReport(`${DIR}/date-system-1904.xlsx`, CELLS);
        assert.strictEqual(report.epoch, 1904, 'the declaration must reach the model');
        assert.strictEqual(report.eager.A2.type, 'date');
        assert.strictEqual(report.eager.A2.value.date, MARCH_15);
        assert.strictEqual(report.eager.A4.value.result, MARCH_16, 'a cached result too');
      },
    },
    {
      name: 'the 1904 file and its 1900 twin read back to the same dates, from different serials',
      async expect(api: CorpusApi, assert: Assert) {
        // The whole point of the pair. The two files store 43538 and 45000 for one calendar day, so
        // a reader that agreed with both is a reader that read the declaration rather than assuming.
        const bare = await api.dateSystemReport(`${DIR}/date-system-1900.xlsx`, CELLS);
        assert.strictEqual(bare.epoch, 1900);
        assert.strictEqual(bare.eager.A2.value.date, MARCH_15);
      },
    },
    {
      name: 'the streaming reader applies the same date system as the buffered one',
      async expect(api: CorpusApi, assert: Assert) {
        for (const file of ['date-system-1904.xlsx', 'date-system-1900.xlsx']) {
          const report = await api.dateSystemReport(`${DIR}/${file}`, CELLS);
          assert.deepStrictEqual(
            report.streaming.A2,
            report.eager.A2,
            `${file}: a row-by-row read must not decode a cell differently from a whole read`,
          );
        }
      },
    },
    {
      name: 'the binary .xlsb serialisation carries the same declaration',
      async expect(api: CorpusApi, assert: Assert) {
        // BIFF12 states it in a record rather than an attribute, and the twin files say what the
        // record means: the same workbook saved both ways must read back the same dates.
        const binary = await api.dateSystemReport(`${DIR}/date-system-1904.xlsb`, CELLS);
        assert.strictEqual(binary.epoch, 1904);
        assert.strictEqual(binary.eager.A2.value.date, MARCH_15);
        const bare = await api.dateSystemReport(`${DIR}/date-system-1900.xlsb`, CELLS);
        assert.strictEqual(bare.epoch, 1900);
        assert.strictEqual(bare.eager.A2.value.date, MARCH_15);
      },
    },
    {
      name: 'writing a 1904 workbook back keeps both the declaration and the dates',
      async expect(api: CorpusApi, assert: Assert) {
        // Dropping the flag is the silent half of this bug: the serials are re-emitted unchanged, so
        // the file still says 43538 while no longer saying what 43538 counts from.
        const report = await api.dateSystemReport(`${DIR}/date-system-1904.xlsx`, CELLS);
        assert.strictEqual(report.roundtripEpoch, 1904);
        assert.strictEqual(report.roundtrip.A2.value.date, MARCH_15);
        assert.strictEqual(report.roundtrip.A4.value.result, MARCH_16);
      },
    },
  ],
} satisfies Case;
