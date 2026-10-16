import assert from 'node:assert/strict';
import {test} from 'node:test';
import {inspect} from 'node:util';

import {strToU8, zipSync} from 'fflate';

import {Workbook} from '../../core/workbook.ts';
import {writeCompoundFile} from '../../vba/cfb-writer.ts';
import {PackageReadError, UnsupportedFormatError} from '../opc/errors.ts';
import {sniffContainer} from '../opc/sniff-format.ts';
import {readXlsx} from './read.ts';
import {writeXlsx} from './write.ts';

/** A genuine `.xlsx` package: the control that must keep reading after the probe is in front. */
function validXlsx(): Uint8Array {
  const wb = new Workbook();
  wb.addWorksheet('S').getCell('A1').value = 42;
  return writeXlsx(wb);
}

/** A genuine OLE2/CFB compound file: the container a legacy `.xls` uses. */
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

/** Text that is neither a ZIP nor a CFB: a CSV handed to the xlsx reader, say. */
function nonZipBlob(): Uint8Array {
  return strToU8('name,amount\nwidget,10\n');
}

/**
 * A genuine package cut off mid-stream: a half-downloaded or truncated file, which is what a corrupt
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

test('a ZIP that inflates but carries no workbook part keeps the "no workbook part" message', () => {
  // The one case where that message is the truth: the archive unpacked, and the part search that ran
  // over it came up empty. A ZIP-headed stub with no visible entry lands here too: a streaming unzip
  // skips it silently rather than reporting a failure, so nothing is left to report but the absence.
  for (const blob of [zipSync({'not-a-workbook.txt': strToU8('hello')}), zipStubBlob()]) {
    const err = catchError(() => readXlsx(blob));
    assert.ok(err instanceof UnsupportedFormatError);
    assert.equal(err.format, 'unknown');
    assert.match(err.message, /no OOXML workbook part was found/);
  }
});

test('a truncated archive leaks neither raw zip internals nor a filesystem path', () => {
  const err = catchError(() => readXlsx(truncatedZipBlob()));
  assert.ok(err instanceof PackageReadError);
  // The fflate failure ("… end of central directory …", "invalid zip data") must never reach the caller, nor
  // any filesystem path, neither in the message nor through a `cause` chain a logger would print.
  // `inspect` rather than `String`: a logger printing an error prints its `cause` chain in full,
  // and that whole rendering is what must not leak; `String(cause)` would show only its top line.
  const text = `${err.message} ${inspect(err.cause)}`;
  assert.doesNotMatch(text, /central directory|is this a zip|invalid zip|unexpected EOF/i);
  assert.doesNotMatch(text, /[A-Za-z]:\\|\/(?:Users|home)\//);
  assert.equal(err.cause, undefined);
});
