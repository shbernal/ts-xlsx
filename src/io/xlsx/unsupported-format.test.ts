import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strToU8, zipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {writeCompoundFile} from '../../vba/cfb-writer.ts';
import {UnsupportedFormatError} from '../opc/errors.ts';
import {sniffContainer} from '../opc/sniff-format.ts';
import {XlsbParseError} from '../xlsb/errors.ts';
import {readXlsx} from './read.ts';
import {readWorkbookStream} from './read-rows.ts';
import {writeXlsx} from './write.ts';

/** A genuine `.xlsx` package — the control that must keep reading after the probe is in front. */
function validXlsx(): Uint8Array {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 42;
  return writeXlsx(wb);
}

/** A genuine OLE2/CFB compound file — the container a legacy `.xls` uses. */
function cfbBlob(): Uint8Array {
  return writeCompoundFile([{name: 'Workbook', data: strToU8('legacy biff bytes')}]);
}

/** A real ZIP/OPC package whose office document is the binary `xl/workbook.bin`, as a `.xlsb` carries. */
function xlsbBlob(): Uint8Array {
  return zipSync({
    '[Content_Types].xml': strToU8(
      '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>' +
        '</Types>',
    ),
    'xl/workbook.bin': Uint8Array.of(0x00, 0x01, 0x02, 0x03),
  });
}

/** Text that is neither a ZIP nor a CFB — a CSV handed to the xlsx reader, say. */
function nonZipBlob(): Uint8Array {
  return strToU8('name,amount\nwidget,10\n');
}

/** A blob that opens with the ZIP magic but is not a parseable archive — corrupt or truncated. */
function corruptZipBlob(): Uint8Array {
  return Uint8Array.of(0x50, 0x4b, 0x03, 0x04, 0xff, 0xff, 0xff, 0xff, 0x00, 0x11, 0x22);
}

const catchError = (fn: () => void): unknown => {
  try {
    fn();
  } catch (err) {
    return err;
  }
  throw new assert.AssertionError({message: 'expected the read to throw'});
};

test('sniffContainer classifies by leading magic bytes', () => {
  assert.equal(sniffContainer(cfbBlob()), 'cfb');
  assert.equal(sniffContainer(validXlsx()), 'zip');
  assert.equal(sniffContainer(xlsbBlob()), 'zip');
  assert.equal(sniffContainer(nonZipBlob()), 'other');
  assert.equal(sniffContainer(new Uint8Array(0)), 'other');
});

test('a valid .xlsx still reads through the format probe', () => {
  const wb = readXlsx(validXlsx());
  assert.equal(wb.getWorksheet('S')?.getCell('A1').value, 42);
});

test('a legacy .xls (CFB) throws UnsupportedFormatError with format "xls"', () => {
  const err = catchError(() => readXlsx(cfbBlob()));
  assert.ok(err instanceof UnsupportedFormatError);
  assert.equal(err.format, 'xls');
  assert.equal(err.name, 'UnsupportedFormatError');
  assert.match(err.message, /\.xls/);
});

test('a binary .xlsb is dispatched to the BIFF12 codec, not rejected as unreadable', () => {
  // The classification stands — this *is* an `.xlsb` — but `readXlsx` now reads one, so what a
  // deliberately malformed binary workbook must produce is a parse failure, not a format failure.
  const err = catchError(() => readXlsx(xlsbBlob()));
  assert.ok(err instanceof XlsbParseError);
  assert.ok(!(err instanceof UnsupportedFormatError));
});

test('the row streamer still reports a binary .xlsb as a format it cannot take', () => {
  // Row streaming is built on the XML worksheet parser, so it has no binary path yet. It must say so
  // in terms of the format — and point at the entry point that does read one.
  const err = catchError(() => {
    for (const _sheet of readWorkbookStream(xlsbBlob())) break;
  });
  assert.ok(err instanceof UnsupportedFormatError);
  assert.equal(err.format, 'xlsb');
  assert.match(err.message, /\.xlsb/);
  assert.match(err.message, /readXlsx|readXlsb/);
});

test('non-ZIP input throws UnsupportedFormatError with format "unknown"', () => {
  const err = catchError(() => readXlsx(nonZipBlob()));
  assert.ok(err instanceof UnsupportedFormatError);
  assert.equal(err.format, 'unknown');
  assert.match(err.message, /not a valid \.xlsx package/);
});

test('a ZIP-headed but corrupt archive fails typed, never leaking raw zip internals', () => {
  const err = catchError(() => readXlsx(corruptZipBlob()));
  assert.ok(err instanceof UnsupportedFormatError);
  assert.equal(err.format, 'unknown');
  // The fflate failure ("… end of central directory …") must never surface, nor any filesystem path.
  assert.doesNotMatch(err.message, /central directory|is this a zip/i);
  assert.doesNotMatch(err.message, /[A-Za-z]:\\|\/(?:Users|home)\//);
});

test('the streaming reader raises the same typed error for a non-.xlsx input', () => {
  const err = catchError(() => {
    for (const _sheet of readWorkbookStream(cfbBlob())) break;
  });
  assert.ok(err instanceof UnsupportedFormatError);
  assert.equal(err.format, 'xls');
});
