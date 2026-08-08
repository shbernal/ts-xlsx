import {messageOf} from '../../thrown.ts';
// The CSV reader and writer, including encoding behaviour.

import type {Untyped} from '../../untyped.ts';
import {
  normalizeCsvValue,
  specCsvValue,
  translateCsvReadOptions,
  translateCsvWriteOptions,
} from './csv-support.ts';
import {readCsv, Workbook, writeCsv, writeCsvText} from './runtime.ts';

export const csv = {
  csvRead({csv, options}: Untyped = {}) {
    try {
      const wb = readCsv(csv, translateCsvReadOptions(options));
      const sheet = wb.worksheets[0];
      const rows = [];
      if (sheet) {
        for (const {cells} of sheet.rows()) {
          let width = 0;
          for (const cell of cells) if (cell.col > width) width = cell.col;
          const fields = new Array(width).fill(null);
          for (const cell of cells) fields[cell.col - 1] = normalizeCsvValue(cell.value);
          rows.push(fields);
        }
      }
      return {ok: true, error: null, rows};
    } catch (e) {
      return {ok: false, error: messageOf(e), rows: []};
    }
  },

  csvWrite({spec = {}, options}: Untyped = {}) {
    try {
      const wb = new Workbook();
      const sheet = wb.addWorksheet('S');
      for (const row of spec.rows || []) sheet.addRow((row || []).map(specCsvValue));
      const text = writeCsvText(wb, translateCsvWriteOptions(options));
      return {ok: true, error: null, text};
    } catch (e) {
      return {ok: false, error: messageOf(e), text: null};
    }
  },

  csvWriteSheetSelection(sheetName: Untyped) {
    const wb = new Workbook();
    wb.addWorksheet('First').addRow(['a', 1]);
    const second = wb.addWorksheet('Second');
    second.addRow(['b', 2]);
    second.addRow(['c', 3]);
    let error = null;
    let text = null;
    try {
      text = writeCsvText(wb, sheetName === undefined ? {} : {sheetName});
    } catch (e) {
      error = messageOf(e);
    }
    return {
      ok: error === null,
      error,
      text,
      rowCount: text ? text.split(/\r?\n/).filter(Boolean).length : 0,
    };
  },

  csvReadMapReport() {
    const csv = 'id,amount\n007,32.5\n008,40';
    const read = (map: Untyped) => {
      const wb = readCsv(csv, map ? {map} : {});
      const sheet = wb.worksheets[0];
      const a = sheet ? sheet.getCell('A2').value : null;
      const b = sheet ? sheet.getCell('B2').value : null;
      return {a, aType: typeof a, b, bType: typeof b};
    };
    return {default: read(null), identity: read((v: Untyped) => v)};
  },

  csvWriteEncodingReport({encoding = 'utf16le', text = 'café'}: Untyped = {}) {
    const EMOJI = '😀🎉';
    const CJK = '日本語テスト';

    const fidelityWb = new Workbook();
    fidelityWb.addWorksheet('S').addRow([EMOJI, CJK]);
    const reread = readCsv(writeCsv(fidelityWb)).worksheets[0];
    const emojiRoundtrips =
      !!reread && reread.getCell('A1').value === EMOJI && reread.getCell('B1').value === CJK;

    const encodedWb = new Workbook();
    encodedWb.addWorksheet('S').addRow([text]);
    const encodedBuffer = Buffer.from(writeCsv(encodedWb, {encoding}));
    const decodesAsRequested = encodedBuffer.toString(encoding).replace(/\r?\n$/, '') === text;
    const decodesAsUtf8 = encodedBuffer.toString('utf8').replace(/\r?\n$/, '') === text;

    return {emojiRoundtrips, requestedEncoding: encoding, decodesAsRequested, decodesAsUtf8};
  },

  csvNonAsciiEncodingReport(text = 'שלום') {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = text;
    sheet.getCell('B1').value = 'world';
    const buffer = Buffer.from(writeCsv(wb));
    const hasBom = buffer[0] === 0xef && buffer[1] === 0xbb && buffer[2] === 0xbf;
    const body = hasBom ? buffer.subarray(3) : buffer;
    return {hasBom, bytesDecodeToText: body.toString('utf8').startsWith(text)};
  },
};
