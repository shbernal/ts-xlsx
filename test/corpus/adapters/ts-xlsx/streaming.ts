// The streaming writer and the streaming reader, including every case that asserts the
// streamed result matches what the eager path produces.

import {tmpdir} from 'node:os';
import {Duplex, PassThrough} from 'node:stream';
import {strFromU8, unzipSync} from 'fflate';
import {codeOrMessageOf, messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {
  decodeRange,
  detectValueType,
  fixtureBytes,
  JSZip,
  readFixture,
  readWorkbookStream,
  readXlsx,
  Workbook,
  WorkbookStreamWriter,
  writeXlsx,
} from './runtime.ts';
import {
  buildFrom,
  modelValueToSpec,
  normalizeStreamValue,
  ONE_PX_PNG,
  specValueToModel,
  streamedRowValues,
} from './spec-model.ts';
import {buildReadInput, classifyReadError, type ReadInputKind} from './xml-probes.ts';

export const streaming = {
  // The same classification through the STREAMING reader, driven far enough to open the package — for
  // asserting the streaming entry point is wired to the identical typed-error contract.
  classifyStreamReadInput(kind: ReadInputKind) {
    return classifyReadError(() => {
      for (const _sheet of readWorkbookStream(buildReadInput(kind))) break;
    });
  },

  // Drive the streaming writer to produce a package, then treat the bytes as an UNTRUSTED archive:
  // JSZip (an independent zip impl) extracts with CRC checking on, so a mismatched entry throws. Also
  // report part count, empty parts, and a whole-file re-read → { partCount, emptyParts, crcValid,
  // crcError, reloadOk, reloadError, sheetNames, firstCol }. The streamed container must be valid with
  // no repair step.
  async streamWritePackageReport({rows = 50} = {}) {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    for (let i = 1; i <= rows; i++) sheet.addRow([`r${i}`, i]).commit();
    sheet.commit();
    const buffer = Buffer.from(await writer.commit());

    let crcValid = true;
    let crcError = null;
    const emptyParts = [];
    let partCount = 0;
    try {
      const zip = await JSZip.loadAsync(buffer, {checkCRC32: true});
      const names = Object.keys(zip.files).filter((n) => !zip.files[n].dir);
      partCount = names.length;
      for (const n of names) {
        const bytes = await zip.files[n].async('nodebuffer');
        if (bytes.length === 0) emptyParts.push(n);
      }
    } catch (e) {
      crcValid = false;
      crcError = messageOf(e);
    }

    let reloadOk = true;
    let reloadError = null;
    let sheetNames: string[] = [];
    const firstCol = [];
    try {
      const wb = readXlsx(buffer);
      sheetNames = wb.worksheets.map((s: Untyped) => s.name);
      const s = wb.worksheets[0];
      for (let i = 1; i <= Math.min(rows, 3); i++)
        firstCol.push(normalizeStreamValue(s!.getCell(`A${i}`).value));
    } catch (e) {
      reloadOk = false;
      reloadError = messageOf(e);
    }
    return {partCount, emptyParts, crcValid, crcError, reloadOk, reloadError, sheetNames, firstCol};
  },

  // Drive the streaming worksheet writer through row ops (addRow/addRows), commit, and read the
  // requested cells back → { ok, error, cells, rowCount }. Exercises the batch-add convenience and the
  // single-row control on the same path.
  async streamWriteSheet({ops = [], read = [], useSharedStrings = false}: Untyped = {}) {
    const toRow = (values: Untyped) => (values || []).map(specValueToModel);
    const writer = new WorkbookStreamWriter({useSharedStrings});
    let error = null;
    try {
      const sheet = writer.addWorksheet('S');
      for (const op of ops) {
        if (op.op === 'addRow') sheet.addRow(toRow(op.value)).commit();
        else if (op.op === 'addRows') sheet.addRows((op.value || []).map(toRow));
        else throw new Error(`unknown stream op: ${op.op}`);
      }
      sheet.commit();
    } catch (e) {
      error = messageOf(e);
    }
    const buffer = Buffer.from(await writer.commit());
    // Declared before the failure return so both paths report the same `cells` type. A bare `{}` on
    // the error path widened the result to `Record<string, …> | {}`, and a case reading `cells.A1`
    // then had to defeat the union before it could assert anything.
    const cells: Record<string, Untyped> = {};
    if (error) return {ok: false, error, cells, rowCount: 0};

    const s = readXlsx(buffer).worksheets[0];
    for (const ref of read)
      cells[ref] = modelValueToSpec(normalizeStreamValue(s!.getCell(ref).value));
    return {ok: true, error: null, cells, rowCount: s!.rowCount};
  },

  // Commit a streaming worksheet, then add a row → { rejected, error, legibleRejection, internalCrash,
  // reloadOk }. A post-commit add must be rejected with a legible "already committed" error, not an
  // internal null-property crash, and a cleanly-committed workbook still reads back.
  async streamAddRowAfterCommit() {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    sheet.addRow(['a']).commit();
    sheet.commit();
    let error = null;
    try {
      sheet.addRow(['b']).commit();
    } catch (e) {
      error = messageOf(e);
    }
    const buffer = Buffer.from(await writer.commit());
    const legibleRejection = error != null && /commit|committed|finaliz|closed/i.test(error);
    const internalCrash = error != null && /Cannot read propert|of (null|undefined)/i.test(error);
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {rejected: error != null, error, legibleRejection, internalCrash, reloadOk};
  },

  // Stream-write a sheet carrying both an autofilter and sheet protection, then inspect the emitted
  // worksheet XML for CT_Worksheet ordering → { protectThrew, sheetProtectionBeforeAutoFilter,
  // reloadOk }. <sheetProtection> must precede <autoFilter>; the shared serializer guarantees it.
  async streamAutoFilterProtectionOrder() {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    sheet.addRow(['H1', 'H2']).commit();
    sheet.addRow(['a', 'b']).commit();
    sheet.autoFilter = 'A1:B1';
    let protectThrew = false;
    try {
      sheet.protect('pw', {});
    } catch {
      protectThrew = true;
    }
    sheet.commit();
    const buffer = Buffer.from(await writer.commit());
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const posProt = xml.indexOf('<sheetProtection');
    const posAf = xml.indexOf('<autoFilter');
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {
      protectThrew,
      sheetProtectionBeforeAutoFilter: posProt >= 0 && posAf >= 0 && posProt < posAf,
      reloadOk,
    };
  },

  // Probe the streaming writer's image parity → { writerAddImage, sheetAddImage, error, mediaParts,
  // drawingParts }. A registered image anchored on a streamed sheet must embed a media binary and a
  // drawing part, exactly like the in-memory writer. The oracle anchors by range string; the model's
  // addImage takes grid points, so decode the range into a tl/br rectangle here.
  async streamWriterImageSupport(range = 'B2:D6') {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    const surface = {
      writerAddImage: typeof writer.addImage === 'function',
      sheetAddImage: typeof sheet.addImage === 'function',
    };
    let error = null;
    let buffer = null;
    try {
      const imageId = writer.addImage({buffer: ONE_PX_PNG, extension: 'png'});
      const {left, top, right, bottom} = decodeRange(range);
      sheet.addImage(imageId, {
        tl: {col: left! - 1, row: top! - 1},
        br: {col: right!, row: bottom!},
      });
      sheet.addRow(['x']).commit();
      sheet.commit();
      buffer = Buffer.from(await writer.commit());
    } catch (e) {
      error = messageOf(e);
    }
    let mediaParts: string[] = [];
    let drawingParts: string[] = [];
    if (!error && buffer) {
      const parts = Object.keys(partMapOf(buffer));
      mediaParts = parts.filter((n) => /xl\/media\//.test(n));
      drawingParts = parts.filter((n) => /drawing/.test(n));
    }
    return {...surface, error, mediaParts, drawingParts};
  },

  // Stream-read a styled workbook and rebuild it through the streaming writer, copying each cell's
  // value AND resolved style onto the new sheet → { copyError, loadOk, fontBold, fontColor, numFmt,
  // hasFill }. The streaming reader surfaces each cell's style facets, so the per-cell font, fill, and
  // number format survive the streaming read→write copy and the emitted styles part loads cleanly.
  async streamingStyleCopyReport() {
    const src = new Workbook();
    const c = src.addWorksheet('S').getCell('A1');
    c.value = 'hello';
    c.font = {bold: true, color: {argb: 'FFFF0000'}};
    c.numFmt = '0.00%';
    c.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
    const srcBuffer = writeXlsx(src);

    const writer = new WorkbookStreamWriter();
    let copyError = null;
    try {
      for (const sheet of readWorkbookStream(srcBuffer)) {
        const ows = writer.addWorksheet(sheet.name);
        for (const row of sheet.rows()) {
          for (const cell of row.cells) {
            const target = ows.getCell(cell.address);
            target.value = cell.value;
            if (cell.style) {
              if (cell.style.font !== undefined) target.font = cell.style.font;
              if (cell.style.fill !== undefined) target.fill = cell.style.fill;
              if (cell.style.numFmt !== undefined) target.numFmt = cell.style.numFmt;
              if (cell.style.border !== undefined) target.border = cell.style.border;
              if (cell.style.alignment !== undefined) target.alignment = cell.style.alignment;
            }
          }
        }
        ows.commit();
      }
    } catch (e) {
      copyError = messageOf(e);
    }
    const buffer = Buffer.from(await writer.commit());
    if (copyError)
      return {
        copyError,
        loadOk: false,
        fontBold: null,
        fontColor: null,
        numFmt: null,
        hasFill: null,
      };

    let loadOk = true;
    let cell = null;
    try {
      cell = readXlsx(buffer).getWorksheet('S')!.getCell('A1');
    } catch {
      loadOk = false;
    }
    return {
      copyError,
      loadOk,
      fontBold: cell ? !!cell.font?.bold : null,
      fontColor: cell?.font?.color ? (cell.font.color.argb ?? null) : null,
      numFmt: cell ? (cell.numFmt ?? null) : null,
      hasFill: cell ? !!(cell.fill && cell.fill.type === 'pattern' && cell.fill.fgColor) : null,
    };
  },

  // Pipe the streaming writer's output stream to a sink and report { pipeReturnsDestination, bytes,
  // valid }. Node's Readable.pipe(dest) must RETURN dest so `.pipe(out).on('finish', …)` composes,
  // while the piped payload still reconstitutes a valid workbook.
  async streamWriterPipeContract() {
    const writer = new WorkbookStreamWriter();
    const source = writer.stream;
    const sink = new PassThrough();
    const chunks: Untyped[] = [];
    sink.on('data', (c) => chunks.push(c));
    const pipeReturn = source.pipe(sink);
    const pipeReturnsDestination = pipeReturn === sink;
    const ws = writer.addWorksheet('S');
    ws.addRow(['a', 'b']).commit();
    ws.commit();
    await writer.commit();
    await new Promise((res) => setTimeout(res, 20));
    const buffer = Buffer.concat(chunks);
    let valid = false;
    try {
      valid = readXlsx(buffer).worksheets[0]!.getCell('A1').value === 'a';
    } catch {
      valid = false;
    }
    return {pipeReturnsDestination, bytes: buffer.length, valid};
  },

  // Request fullCalcOnLoad on the streaming writer (via calcProperties) and report whether it reaches
  // the output, versus the in-memory writer → { streamSetThrew, streamHasFlag, streamDefaultHasFlag,
  // memoryHasFlag }. Recalc-on-load must work identically on both writers.
  async streamingFullCalcOnLoadReport() {
    const streamCalc = async (setFlag: Untyped) => {
      const writer = new WorkbookStreamWriter();
      let threw = false;
      if (setFlag) {
        try {
          writer.calcProperties.fullCalcOnLoad = true;
        } catch {
          threw = true;
        }
      }
      const sheet = writer.addWorksheet('S');
      sheet.getCell('A1').value = 1;
      sheet.commit();
      const buffer = Buffer.from(await writer.commit());
      const wbXml = strFromU8(unzipSync(buffer)['xl/workbook.xml']!);
      return {threw, hasFlag: /fullCalcOnLoad="1"/.test(wbXml)};
    };
    const set = await streamCalc(true);
    const def = await streamCalc(false);

    const wb = new Workbook();
    wb.fullCalcOnLoad = true;
    wb.addWorksheet('S').getCell('A1').value = 1;
    const memXml = strFromU8(unzipSync(writeXlsx(wb))['xl/workbook.xml']!);

    return {
      streamSetThrew: set.threw,
      streamHasFlag: set.hasFlag,
      streamDefaultHasFlag: def.hasFlag,
      memoryHasFlag: /fullCalcOnLoad="1"/.test(memXml),
    };
  },

  // Build via the streaming writer with a master formula + shared-formula slaves, reload →
  // { masterHasFormula, slaveResolved, slaveValue }. Streamed slaves must not be dropped to empty.
  async streamingSharedFormulaReport(rows = 10) {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('yua');
    for (let i = 1; i <= rows; i++) sheet.getCell(`A${i}`).value = i * 10;
    sheet.getCell('B1').value = {formula: 'A1*2', result: 20};
    for (let j = 2; j <= rows; j++) sheet.getCell(`B${j}`).value = {sharedFormula: 'B1'};
    sheet.commit();
    const buffer = Buffer.from(await writer.commit());

    const rs = readXlsx(buffer).getWorksheet('yua');
    const slave = rs!.getCell('B3').value;
    const slaveIsEmpty =
      slave == null || (typeof slave === 'object' && Object.keys(slave).length === 0);
    const master = rs!.getCell('B1').value;
    return {
      masterHasFormula: !!(master && typeof master === 'object' && 'formula' in master),
      slaveResolved: !slaveIsEmpty,
      slaveValue: normalizeStreamValue(slave ?? null),
    };
  },

  // Stream-write a sheet carrying both a conditional-formatting rule and a hyperlink cell, then report
  // the relative order of the emitted <conditionalFormatting> and <hyperlinks> blocks plus reload
  // success. Both writers share one worksheet serializer, so the streamed sheet emits the blocks in
  // CT_Worksheet order (conditionalFormatting before hyperlinks) rather than the reversed order the
  // upstream streaming writer produced.
  async streamWriteCfHyperlinkOrder() {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
    sheet.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {
          type: 'expression',
          formulae: ['MOD(ROW(),2)=0'],
          style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFEEEEEE'}}},
        },
      ],
    });
    sheet.addRow(['x']).commit();
    sheet.commit();

    const buffer = Buffer.from(await writer.commit());
    const xml = strFromU8(unzipSync(buffer)['xl/worksheets/sheet1.xml']!);
    const posCf = xml.indexOf('<conditionalFormatting');
    const posHl = xml.indexOf('<hyperlinks');
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {
      posConditionalFormatting: posCf,
      posHyperlinks: posHl,
      conditionalFormattingBeforeHyperlinks: posCf >= 0 && posHl >= 0 ? posCf < posHl : null,
      reloadOk,
    };
  },

  // Stream-write a sheet carrying both a data validation and a hyperlink cell, then report the relative
  // order of the emitted <dataValidations> and <hyperlinks> blocks plus reload success. CT_Worksheet
  // requires dataValidations before hyperlinks; the shared serializer emits them in that order on the
  // streaming path too. Companion to streamWriteCfHyperlinkOrder.
  async streamWriteDvHyperlinkOrder() {
    const writer = new WorkbookStreamWriter();
    const sheet = writer.addWorksheet('S');
    sheet.getCell('A1').value = {text: 'link', hyperlink: 'https://example.com'};
    sheet.addDataValidation('B1', {type: 'list', allowBlank: true, formulae: ['"x,y,z"']});
    sheet.addRow(['r']).commit();
    sheet.commit();

    const buffer = Buffer.from(await writer.commit());
    const xml = strFromU8(unzipSync(buffer)['xl/worksheets/sheet1.xml']!);
    const posDv = xml.indexOf('<dataValidations');
    const posHl = xml.indexOf('<hyperlinks');
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {
      posDataValidations: posDv,
      posHyperlinks: posHl,
      dataValidationsBeforeHyperlinks: posDv >= 0 && posHl >= 0 ? posDv < posHl : null,
      reloadOk,
    };
  },

  // Commit a streaming workbook over a caller-supplied writable (a plain PassThrough or a Duplex) and
  // report { settled, timedOut, bytes, valid }. The commit must settle within bounded time and the
  // sink must receive a complete, re-openable package — the library owes this even when it does not own
  // the stream.
  async streamCommitReport({duplex = false, timeoutMs = 4000} = {}) {
    const chunks: Untyped[] = [];
    const stream = duplex
      ? new Duplex({
          read() {},
          write(c, _e, cb) {
            chunks.push(c);
            cb();
          },
        })
      : new PassThrough();
    if (!duplex) stream.on('data', (c) => chunks.push(c));

    const writer = new WorkbookStreamWriter({stream});
    const sheet = writer.addWorksheet('S');
    sheet.addRow(['a', 'b']).commit();
    sheet.commit();

    let settled = 'pending';
    const commit = writer.commit().then(
      () => (settled = 'resolved'),
      (e) => (settled = `rejected:${messageOf(e)}`),
    );
    const timedOut = await Promise.race([
      commit.then(() => false),
      new Promise((res) => setTimeout(() => res(true), timeoutMs)),
    ]);

    let valid = false;
    if (settled === 'resolved') {
      try {
        const back = readXlsx(Buffer.concat(chunks));
        valid = back.worksheets.length === 1 && back.worksheets[0]!.getCell('A1').value === 'a';
      } catch {
        valid = false;
      }
    }
    return {settled, timedOut, bytes: Buffer.concat(chunks).length, valid};
  },

  // Commit a streaming workbook to a destination that cannot be opened for writing and report
  // { outcome, rejected, carriesIoError, error }. The write stream errors on a later tick; commit must
  // reject with that I/O error rather than hanging forever.
  async streamCommitBadDestination() {
    const badPath = `${tmpdir()}/ts-xlsx-no-such-dir-${process.pid}/${'x'.repeat(300)}/out.xlsx`;
    let outcome = 'hung';
    let error = null;
    try {
      const writer = new WorkbookStreamWriter({filename: badPath});
      const sheet = writer.addWorksheet('S');
      sheet.addRow(['a']).commit();
      sheet.commit();
      await Promise.race([
        writer
          .commit()
          .then(() => {
            outcome = 'resolved';
          })
          .catch((e) => {
            outcome = 'rejected';
            error = codeOrMessageOf(e);
          }),
        new Promise((res) => setTimeout(res, 5000)),
      ]);
    } catch (e) {
      outcome = 'threw-sync';
      error = codeOrMessageOf(e);
    }
    return {
      outcome,
      rejected: outcome === 'rejected',
      carriesIoError: error != null && /ENOENT|ENAMETOOLONG|ENOTDIR|open|write/i.test(error),
      error,
    };
  },

  // Read a fixture both eagerly and through the streaming reader, reporting the sheet names each
  // surfaces → { eager, streaming }. The streaming reader joins each worksheet part to the
  // workbook-level declaration, so it exposes the real names, not positional placeholders.
  streamVsEagerSheetNames(rel: string) {
    const eager = readFixture(rel).worksheets.map((s) => s.name);
    const streaming = [...readWorkbookStream(fixtureBytes(rel))].map((s) => s.name);
    return {eager, streaming};
  },

  // Report the first sheet's populated row numbers from both paths → { eager, streaming }. Both skip
  // fully-empty rows (the eager `includeEmpty:false` intent) so a gap between data rows is preserved
  // as a jump in the numbers, never resequenced.
  streamVsEagerRowNumbers(rel: string) {
    const es = readFixture(rel).worksheets[0];
    const eager: Untyped[] = [];
    if (es) for (const row of es.rows()) if (row.cells.length) eager.push(row.number);
    const streaming: Untyped[] = [];
    for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
      for (const row of sheet.rows()) if (row.cells.length) streaming.push(row.number);
      break; // first worksheet only
    }
    return {eager, streaming};
  },

  // Report each populated first-sheet row's { number, hidden } from both paths → { eager, streaming }.
  // The streaming reader must surface a row's hidden flag (in the string form "true"/"false" some
  // generators write), agreeing with the eager read rather than reporting every row visible.
  streamVsEagerRowHidden(rel: string) {
    const es = readFixture(rel).worksheets[0];
    const eager: Untyped[] = [];
    if (es)
      for (const row of es.rows())
        if (row.cells.length) eager.push({number: row.number, hidden: !!row.properties?.hidden});
    const streaming: Untyped[] = [];
    for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
      for (const row of sheet.rows())
        if (row.cells.length) streaming.push({number: row.number, hidden: !!row.hidden});
      break; // first worksheet only
    }
    return {eager, streaming};
  },

  // Write a sheet with a hidden column, then read it eagerly and through the streaming reader,
  // reporting each path's per-column hidden flags → { eager, stream, error }. The streaming reader
  // parses <col hidden> and surfaces it after the rows are drained, matching the eager oracle.
  streamVsEagerColumnHidden() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getColumn(2).hidden = true;
    s.getCell('A1').value = 'a';
    s.getCell('B1').value = 'b';
    s.getCell('C1').value = 'c';
    const buffer = writeXlsx(wb);

    const es = readXlsx(buffer).getWorksheet('S')!;
    const eager = {
      col1: !!es.getColumn(1).hidden,
      col2: !!es.getColumn(2).hidden,
      col3: !!es.getColumn(3).hidden,
    };

    const stream: Record<string, Untyped> = {};
    let error = null;
    try {
      for (const sheet of readWorkbookStream(buffer)) {
        for (const _row of sheet.rows()) void _row;
        const hidden = new Set(sheet.hiddenColumns);
        stream.col1 = hidden.has(1);
        stream.col2 = hidden.has(2);
        stream.col3 = hidden.has(3);
        break; // first worksheet only
      }
    } catch (e) {
      error = messageOf(e);
    }
    return {eager, stream, error};
  },

  // Build a sheet with two merged ranges, then report the merge geometry from both the eager and the
  // streaming path → { eagerMerges, streamedMerges, error }. The streaming reader collects
  // <mergeCells> (which follows <sheetData>) during the same pass and exposes it after the rows.
  streamReadMergesReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'm';
    ws.mergeCells('A1:B2');
    ws.getCell('D1').value = 'n';
    ws.mergeCells('D1:D3');
    const buffer = writeXlsx(wb);

    const eagerMerges = [...readXlsx(buffer).getWorksheet('S')!.merges].sort();

    let streamedMerges = null;
    let error = null;
    try {
      for (const sheet of readWorkbookStream(buffer)) {
        for (const _row of sheet.rows()) void _row;
        streamedMerges = [...sheet.merges].sort();
        break; // first worksheet only
      }
    } catch (e) {
      error = messageOf(e);
    }
    return {eagerMerges, streamedMerges, error};
  },

  // Report the 1-based row-values array for the requested rows from both paths → { eager, streamed }.
  // A streamed row indexes exactly as a full-load row does (empty slot at 0, column A at 1), so a
  // caller can switch readers without re-indexing.
  streamVsEagerRowValues(spec: Untyped, rowNumbers = [1]) {
    const buffer = writeXlsx(buildFrom(spec));
    const wanted = new Set(rowNumbers);

    const es = readXlsx(buffer).worksheets[0];
    const eager: Record<string, Untyped> = {};
    if (es)
      for (const row of es.rows())
        if (wanted.has(row.number)) eager[row.number] = streamedRowValues(row.cells as Untyped[]);
    for (const n of rowNumbers) eager[n] ??= [null];

    const streamed: Record<string, Untyped> = {};
    for (const sheet of readWorkbookStream(buffer)) {
      for (const row of sheet.rows())
        if (wanted.has(row.number))
          streamed[row.number] = streamedRowValues(row.cells as Untyped[]);
      break; // first worksheet only
    }
    for (const n of rowNumbers) streamed[n] ??= [null];
    return {eager, streamed};
  },

  // Write `count` single-cell worksheets, then stream them back, reporting { written, emitted, error,
  // first, last }. Exercises the reader across far more than 100 sheets and a package whose worksheet
  // parts may precede the workbook part — every sheet must be emitted exactly once.
  streamReadManySheets(count = 180) {
    const wb = new Workbook();
    for (let i = 0; i < count; i++) wb.addWorksheet(`Sheet${i + 1}`).getCell('A1').value = i;
    const buffer = writeXlsx(wb);
    const names: Untyped[] = [];
    let error = null;
    try {
      for (const sheet of readWorkbookStream(buffer)) names.push(sheet.name);
    } catch (e) {
      error = messageOf(e);
    }
    return {
      written: count,
      emitted: names.length,
      error,
      first: names[0] ?? null,
      last: names[names.length - 1] ?? null,
    };
  },

  // Write a shared-strings workbook, then read it through the streaming reader once and again
  // concurrently, reporting whether every shared-string cell resolved → { singleComplete, singleLength,
  // concurrentAllComplete, concurrentLengths }. The rewrite's reader is a pure synchronous generator,
  // so concurrent reads cannot race over a shared shared-strings table.
  async streamingSharedStringsRead(rowCount = 20, concurrency = 8) {
    const build = new Workbook();
    const bs = build.addWorksheet('S');
    for (let r = 1; r <= rowCount; r++) {
      bs.getCell(`A${r}`).value = `str${r % 3}`;
      bs.getCell(`B${r}`).value = r;
    }
    const buffer = writeXlsx(build, {useSharedStrings: true});

    const readOne = () => {
      const strings = [];
      for (const sheet of readWorkbookStream(buffer)) {
        for (const row of sheet.rows()) {
          const first = row.cells.find((cell) => cell.col === 1);
          strings.push(first ? first.value : undefined);
        }
      }
      return strings;
    };

    const single = readOne();
    const singleComplete = single.length === rowCount && single.every((v) => typeof v === 'string');
    const many = await Promise.all(Array.from({length: concurrency}, async () => readOne()));
    const allComplete = many.every(
      (v) => v.length === rowCount && v.every((x) => typeof x === 'string'),
    );
    return {
      singleComplete,
      singleLength: single.length,
      concurrentAllComplete: allComplete,
      concurrentLengths: many.map((v) => v.length),
    };
  },

  // Stream-read a fixture end-to-end, reporting { ok, error, sheetNames, totalRows } — the read either
  // completes (with every sheet name and the total rows delivered) or its error is captured as data.
  // Locks that the reader tolerates a package whose ZIP places a worksheet part before xl/workbook.xml
  // (the inflate builds a path→bytes map, so entry order is irrelevant).
  streamReadReport(rel: string) {
    const sheetNames: Untyped[] = [];
    let totalRows = 0;
    try {
      for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
        sheetNames.push(sheet.name);
        for (const _row of sheet.rows()) totalRows += 1;
      }
      return {ok: true, error: null, sheetNames, totalRows};
    } catch (e) {
      return {ok: false, error: messageOf(e), sheetNames, totalRows};
    }
  },

  // Stream-read a fixture's first sheet, reporting each requested cell's { type, value } | null. The
  // type is the model's stable label; a date-formatted numeric cell surfaces as a Date because the
  // streaming reader applies the cell's number format when decoding, exactly as the eager read does.
  streamReadFixture(rel: string, cells: Untyped = []) {
    const wanted = new Map(cells.map((a: Untyped) => [a, null]));
    for (const sheet of readWorkbookStream(fixtureBytes(rel))) {
      for (const row of sheet.rows()) {
        for (const cell of row.cells) {
          if (wanted.has(cell.address))
            wanted.set(cell.address, {
              type: detectValueType(cell.value),
              value: normalizeStreamValue(cell.value),
            });
        }
      }
      break; // first worksheet only
    }
    const out: Record<string, Untyped> = {};
    for (const [k, v] of wanted) out[k as string] = v;
    return out;
  },

  // Write a spec, then read the requested cells through both paths → { streamed, eager }. Proves the
  // streaming reader returns multi-byte UTF-8 text (CJK, emoji) byte-exact and identical to the eager
  // read — the whole-package inflate decodes UTF-8 as one unit, so no character is split.
  streamReadSpec(spec: Untyped, cells: Untyped = []) {
    const buffer = writeXlsx(buildFrom(spec));
    const wanted = new Set(cells);
    const streamed: Record<string, Untyped> = {};
    for (const sheet of readWorkbookStream(buffer)) {
      for (const row of sheet.rows()) {
        for (const cell of row.cells)
          if (wanted.has(cell.address)) streamed[cell.address] = normalizeStreamValue(cell.value);
      }
      break; // first worksheet only
    }
    const es = readXlsx(buffer).worksheets[0];
    const eager: Record<string, Untyped> = {};
    for (const ref of cells) eager[ref] = normalizeStreamValue(es ? es.getCell(ref).value : null);
    return {streamed, eager};
  },
};
