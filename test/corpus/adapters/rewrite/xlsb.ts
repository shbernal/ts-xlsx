// Reading the binary `.xlsb` serialisation. The headline capability compares the two readings of one
// workbook Excel saved in both forms, which is why the fixture is a *pair*: the XML twin is an
// independent oracle for what the binary must decode to, not something this library produced.

import {strToU8, zipSync} from 'fflate';

import type {CorpusApi} from '../../case.ts';
import {fixtureBytes, readXlsb, readXlsx} from './runtime.ts';

const FIXTURE = 'xlsb-binary-workbook-reads-like-its-xlsx-twin';

export const xlsb = {
  // Read the same workbook from its binary and its XML serialisation and report whether the two
  // models agree — sheet names, order and visibility; every cell's value and style facets; row and
  // column geometry; merges. On disagreement, name the first field that differs so a failure is
  // legible without a debugger.
  //
  // Formula *text* is projected away on both sides: a BIFF12 formula is a Ptg token stream this
  // reader does not decode yet, so the binary side has the cached result where the XML side has
  // `{formula, result}`. Everything else is compared strictly.
  xlsbModelMatchesXlsxTwin() {
    const binary = snapshot(readXlsx(fixtureBytes(`${FIXTURE}/source.xlsb`)));
    const xml = snapshot(readXlsx(fixtureBytes(`${FIXTURE}/source.xlsx`)));
    if (binary === xml) return {identical: true, firstDifference: null};
    const binaryLines = binary.split('\n');
    const xmlLines = xml.split('\n');
    for (let index = 0; index < Math.max(binaryLines.length, xmlLines.length); index++) {
      if (binaryLines[index] !== xmlLines[index]) {
        return {
          identical: false,
          firstDifference: `xlsb: ${binaryLines[index] ?? '(end)'} | xlsx: ${xmlLines[index] ?? '(end)'}`,
        };
      }
    }
    return {identical: false, firstDifference: 'lengths differ'};
  },

  // The sheets a binary workbook declares, in tab order, with their visibility.
  xlsbSheets() {
    return readXlsb(fixtureBytes(`${FIXTURE}/source.xlsb`)).worksheets.map((sheet: CorpusApi) => ({
      name: sheet.name,
      state: sheet.state,
    }));
  },

  // One cell of the binary reading, as JSON-serializable facts: its decoded value (a Date rendered
  // as an ISO string) and the style facets it resolved to.
  xlsbCell(sheetName: CorpusApi, reference: CorpusApi) {
    const workbook = readXlsb(fixtureBytes(`${FIXTURE}/source.xlsb`));
    const sheet = workbook.getWorksheet(sheetName);
    if (sheet === undefined) return {found: false};
    const cell = sheet.getCell(reference);
    return {
      found: true,
      value: normalize(cell.value),
      type: typeof cell.value,
      numFmt: cell.numFmt ?? null,
      font: cell.font ?? null,
      fill: cell.fill ?? null,
      border: cell.border ?? null,
      alignment: cell.alignment ?? null,
      protection: cell.protection ?? null,
    };
  },

  // A sheet's row/column geometry and merged ranges, from the binary reading.
  xlsbGrid(sheetName: CorpusApi) {
    const workbook = readXlsb(fixtureBytes(`${FIXTURE}/source.xlsb`));
    const sheet = workbook.getWorksheet(sheetName);
    if (sheet === undefined) return {found: false};
    const model = sheet.model;
    return {
      found: true,
      columns: model.columns,
      rows: model.rows,
      merges: model.merges,
    };
  },

  // A ZIP whose office document is a binary workbook that does not conform to the record framing:
  // classification succeeded (it *is* an `.xlsb`), so the failure must be a typed parse error rather
  // than a crash, a hang, or a silently empty workbook.
  xlsbMalformedBinaryWorkbook() {
    // A record header claiming a payload far longer than the part holds — the lever a hostile file
    // would pull.
    const workbook = Uint8Array.of(0x81, 0x00, 0x02, 0xff, 0xff, 0xff, 0x7f, 0x01);
    const archive = zipSync({
      '[Content_Types].xml': strToU8('<?xml version="1.0"?><Types/>'),
      'xl/workbook.bin': workbook,
    });
    try {
      readXlsx(archive);
      return {threw: false, errorName: null, message: null};
    } catch (error) {
      const failure = error as CorpusApi;
      return {
        threw: true,
        errorName: failure?.name ?? null,
        message: String(failure?.message ?? ''),
      };
    }
  },
};

// A Date is not JSON-serializable in a way a case can compare, and a formula value is a shape the
// binary reader does not produce yet; both are flattened the same way on either side of a comparison.
function normalize(value: CorpusApi): CorpusApi {
  if (value instanceof Date) return value.toISOString();
  if (value !== null && typeof value === 'object' && 'formula' in value) {
    return normalize(value.result ?? null);
  }
  return value;
}

// A stable, order-independent rendering of everything a worksheet holds. Key order is an artefact of
// which parser filled the object, so it is normalised away; nothing else is.
function snapshot(workbook: CorpusApi): string {
  return JSON.stringify(
    canonical(
      workbook.worksheets.map((sheet: CorpusApi) => {
        const model = sheet.model;
        return {
          name: sheet.name,
          state: model.state,
          columns: model.columns,
          rows: model.rows,
          merges: model.merges,
          cells: model.cells.map((cell: CorpusApi) => ({...cell, value: normalize(cell.value)})),
        };
      }),
    ),
    null,
    1,
  );
}

function canonical(value: CorpusApi): CorpusApi {
  if (Array.isArray(value)) return value.map(canonical);
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value)
        .sort(([left], [right]) => (left < right ? -1 : 1))
        .map(([key, nested]) => [key, canonical(nested)]),
    );
  }
  return value;
}
