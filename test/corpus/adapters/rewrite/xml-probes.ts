// XML- and byte-level probes: small readers for the corners a model-level assertion cannot
// reach (drawing anchors, sqref expansion, patched-part reloads), plus the synthetic reader
// inputs used to prove the reader classifies foreign formats instead of crashing on them.

import {strFromU8, strToU8, unzipSync, zipSync} from 'fflate';
import type {CorpusApi} from '../../case.ts';
import {decodeRange, encodeAddress, readXlsx, writeCompoundFile, writeXlsx} from './runtime.ts';
import {buildFrom} from './spec-model.ts';

export const hexBytes = (hex: string) =>
  Uint8Array.from((hex.match(/../g) ?? []).map((h) => parseInt(h, 16)));

// Translate a corpus image range — a string like "B2:D6", or a {tl, br?/ext?, editAs?} object — into
// the model's typed addImage call. A one-cell anchor is a point plus a fixed pixel extent (editAs is a
// two-cell-only attribute the model drops by construction); a two-cell anchor spans tl..br. A
// fractional grid coordinate (col 3.5) is passed through — the model floors it to the cell and derives
// the sub-cell EMU offset from that cell's real width/height.
export function anchorSpecImage(sheet: CorpusApi, imageId: CorpusApi, range: CorpusApi) {
  if (typeof range === 'string') {
    const {left, top, right, bottom} = decodeRange(range);
    sheet.addImage(imageId, {tl: {col: left! - 1, row: top! - 1}, br: {col: right, row: bottom}});
    return;
  }
  const {tl, br, ext, editAs} = range;
  if (ext !== undefined) {
    sheet.addImage(imageId, {tl, ext: {width: ext.width, height: ext.height}});
  } else {
    sheet.addImage(imageId, editAs !== undefined ? {tl, br, editAs} : {tl, br});
  }
}

// Parse the integer children of an <xdr:from>/<xdr:to> block, mirroring the oracle so a drawing anchor
// reports the same plain-number geometry from either adapter.
export const intAt = (xml: string, tag: string) => {
  const m = xml.match(new RegExp(`<${tag}>(-?\\d+)</${tag}>`));
  return m ? Number(m[1]) : null;
};
export const parseAnchorSide = (block: string | null | undefined) =>
  block
    ? {
        col: intAt(block, 'xdr:col'),
        colOff: intAt(block, 'xdr:colOff'),
        row: intAt(block, 'xdr:row'),
        rowOff: intAt(block, 'xdr:rowOff'),
      }
    : null;
export const imageXmlWellFormed = (xml: string) =>
  !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml);

// Expand an OOXML sqref (space-separated ranges) into its covered cell references, bounded by a cap so
// a whole-column range never balloons — used to check that a range-form validation is reported on
// every covered cell. An unbounded whole-row/column part is skipped rather than expanded.
export function expandSqref(sqref: CorpusApi, cap = 4096) {
  const refs: string[] = [];
  for (const part of String(sqref).split(/\s+/).filter(Boolean)) {
    const {left, right, top, bottom} = decodeRange(part);
    if (left == null || right == null || top == null || bottom == null) continue;
    for (let c = left; c <= right && refs.length < cap; c++) {
      for (let r = top; r <= bottom && refs.length < cap; r++) refs.push(encodeAddress(c, r));
    }
  }
  return refs;
}

// Rewrite named parts of a written package and read the result back — the way to feed the
// reader the hand-edited OOXML forms real producers emit but the writer itself never generates
// (an explicit-false boolean flag `<b val="0"/>`, an alignment element carrying only `wrapText="0"`,
// an injected xf). `edits` maps a part path to a (xml) => xml transform; unlisted parts pass through.
export function reloadPatched(buffer: Uint8Array, edits: Record<string, (xml: string) => string>) {
  const files = unzipSync(buffer);
  for (const [name, transform] of Object.entries(edits)) {
    files[name] = strToU8(transform(strFromU8(files[name]!)));
  }
  return readXlsx(zipSync(files));
}

// Parse an XML tag's attributes into a plain { name: value } map. base64 salt/hash values use
// only XML-safe characters, so a naive quoted-value scan is sufficient here.
export function attrsOf(tag: string) {
  const out: Record<string, string> = {};
  const re = /([\w:]+)="([^"]*)"/g;
  let m = re.exec(tag);
  while (m !== null) {
    out[m[1]!] = m[2]!;
    m = re.exec(tag);
  }
  return out;
}

// ── Macro-enabled fixture builder ──────────────────────────────────────────────────────────────────
// Assemble a genuine, navigable vbaProject.bin (via the production CFB writer + MS-OVBA compressor, so
// its storage tree is walkable exactly as Excel's is) inside a minimal .xlsm package. This is the only
// way to produce an edit-in-place *input* without an interactive VBA editor: the writer cannot author a
// project from a model (no reference support, document-module linkage is host-coupled), but the editor
// splices new module source into an existing bin. The fixture carries a hand-crafted PROJECTREFERENCES
// record and a `document` code-behind module — the two things splice-editing must preserve that
// re-synthesis structurally cannot.

// Build a reader input of a given format family, to probe the reader's typed-error classification: a
// genuine `.xlsx` (the control that must still read), a legacy `.xls` (an OLE2/CFB compound file, via the
// production CFB writer), a binary `.xlsb` (a real ZIP whose office document is `xl/workbook.bin`),
// non-ZIP text (a CSV handed to the wrong reader), and a ZIP-headed-but-corrupt archive.
export function buildReadInput(kind: CorpusApi): Uint8Array {
  switch (kind) {
    case 'xlsx':
      return writeXlsx(buildFrom({sheets: [{name: 'S', cells: [{ref: 'A1', value: 42}]}]}));
    case 'xls':
      return writeCompoundFile([{name: 'Workbook', data: strToU8('legacy biff bytes')}]);
    case 'xlsb':
      return zipSync({
        '[Content_Types].xml': strToU8(
          '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
            '<Override PartName="/xl/workbook.bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/>' +
            '</Types>',
        ),
        'xl/workbook.bin': Uint8Array.of(0, 1, 2, 3),
      });
    case 'garbage':
      return strToU8('name,amount\nwidget,10\n');
    case 'corrupt-zip': {
      // A genuine package cut off mid-stream — a half-downloaded file, which is what a corrupt archive
      // looks like in the wild. The bytes a zip layer actually rejects are the point: a few hand-made
      // `PK` bytes are quietly skipped by a streaming unzip rather than reported as a failure.
      const good = writeXlsx(buildFrom({sheets: [{name: 'S', cells: [{ref: 'A1', value: 42}]}]}));
      return good.subarray(0, good.length >> 1);
    }
    default:
      throw new Error(`unknown read-input kind: ${kind}`);
  }
}

// Turn a reader call into JSON-serializable classification facts: whether it threw, the error's `name`,
// `code` and `format` branch fields (the typed contract a caller keys on — `code` says what kind of
// failure it is, `format` which unsupported input it was), and whether the message leaks the zip
// layer's internals or an absolute filesystem path (the anti-leak contract).
export function classifyReadError(run: () => void): CorpusApi {
  try {
    run();
    return {threw: false, errorName: null, code: null, format: null, message: null};
  } catch (e) {
    const err = e as CorpusApi;
    const message = String(err?.message ?? '');
    return {
      threw: true,
      errorName: err?.name ?? null,
      code: err?.code ?? null,
      format: err?.format ?? null,
      message,
      leaksZipInternals: /central directory|is this a zip|invalid zip data|unexpected EOF/i.test(
        message,
      ),
      leaksAbsolutePath: /[A-Za-z]:\\|\/(?:Users|home)\//.test(message),
    };
  }
}
