import assert from 'node:assert/strict';
import {test} from 'node:test';

import {strToU8, zipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {writeCompoundFile} from '../../vba/cfb-writer.ts';
import {PackageReadError, UnsupportedFormatError} from '../opc/errors.ts';
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

/**
 * A genuine package cut off mid-stream — a half-downloaded or truncated file, which is what a corrupt
 * archive looks like in the wild. The bytes fflate actually chokes on are the point: a few hand-made
 * `PK` bytes are quietly skipped by a streaming unzip rather than rejected (see {@link zipStubBlob}).
 */
function truncatedZipBlob(): Uint8Array {
  const good = validXlsx();
  return good.subarray(0, good.length >> 1);
}

/** Four `PK` bytes and junk: ZIP-headed, no entry a streaming unzip can even see, let alone reject. */
function zipStubBlob(): Uint8Array {
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
  // The sniff refused before any part was looked for, so the message must not blame a missing part.
  assert.doesNotMatch(err.message, /workbook part/);
});

test('a ZIP that inflates but carries no workbook part keeps the "no workbook part" message', () => {
  // The one case where that message is the truth: the archive unpacked, and the part search that ran
  // over it came up empty. A ZIP-headed stub with no visible entry lands here too — a streaming unzip
  // skips it silently rather than reporting a failure, so nothing is left to report but the absence.
  for (const blob of [zipSync({'not-a-workbook.txt': strToU8('hello')}), zipStubBlob()]) {
    const err = catchError(() => readXlsx(blob));
    assert.ok(err instanceof UnsupportedFormatError);
    assert.equal(err.format, 'unknown');
    assert.match(err.message, /no OOXML workbook part was found/);
  }
});

test('a truncated archive is a PackageReadError, not an unsupported format', () => {
  // Nothing inflated, so no part search ever ran: this is the right *kind* of container that cannot
  // be unpacked — `malformed-input`, not `unsupported-format`.
  const err = catchError(() => readXlsx(truncatedZipBlob()));
  assert.ok(err instanceof PackageReadError);
  assert.equal(err.code, 'malformed-input');
  assert.ok(!(err instanceof UnsupportedFormatError));
  assert.match(err.message, /corrupt or truncated/);
});

test('a truncated archive leaks neither raw zip internals nor a filesystem path', () => {
  const err = catchError(() => readXlsx(truncatedZipBlob()));
  assert.ok(err instanceof PackageReadError);
  // The fflate failure ("… end of central directory …", "invalid zip data") must never surface, nor
  // any filesystem path — neither in the message nor through a `cause` chain a logger would print.
  const text = `${err.message} ${String(err.cause ?? '')}`;
  assert.doesNotMatch(text, /central directory|is this a zip|invalid zip|unexpected EOF/i);
  assert.doesNotMatch(text, /[A-Za-z]:\\|\/(?:Users|home)\//);
  assert.equal(err.cause, undefined);
});

test('the row streamer classifies a truncated archive the same way', () => {
  const err = catchError(() => {
    for (const _sheet of readWorkbookStream(truncatedZipBlob())) break;
  });
  assert.ok(err instanceof PackageReadError);
});

test('the streaming reader raises the same typed error for a non-.xlsx input', () => {
  const err = catchError(() => {
    for (const _sheet of readWorkbookStream(cfbBlob())) break;
  });
  assert.ok(err instanceof UnsupportedFormatError);
  assert.equal(err.format, 'xls');
});
