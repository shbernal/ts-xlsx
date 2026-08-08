// Core model and whole-package behaviour: address decoding, reader input classification,
// package inspection, and the workbook-level round-trips that are not about one feature.

import {canonicalJson} from '../../canonical-json.ts';
import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {packageFacts} from '../ooxml-facts.ts';
import {packagePartFacts, partMapOf} from './package-facts.ts';
import {
  decodeAddress,
  decodeRange,
  fixtureBytes,
  readFixture,
  readXlsx,
  Workbook,
  writeXlsx,
} from './runtime.ts';
import {
  buildFrom,
  isoOrNull,
  normalizeRewriteCell,
  normalizeStreamValue,
  UnsupportedSpecError,
} from './spec-model.ts';
import {buildReadInput, classifyReadError, type ReadInputKind} from './xml-probes.ts';

export const core = {
  // Classify a reader input by format family and report the typed error (or success) it produces —
  // `{threw, errorName, code, format, message, leaksZipInternals, leaksAbsolutePath}` — for asserting a
  // non-`.xlsx` blob fails with a clear, catchable, typed error rather than a raw zip crash.
  classifyReadInput(kind: ReadInputKind) {
    return classifyReadError(() => {
      readXlsx(buildReadInput(kind));
    });
  },

  decodeAddress(reference: string) {
    return decodeAddress(reference);
  },

  decodeRange(reference: string) {
    return decodeRange(reference);
  },

  inspectPackage(spec: Untyped) {
    return packageFacts(spec, partMapOf(writeXlsx(buildFrom(spec))));
  },

  // Same package facts as `inspectPackage`, but after a full write→read→write cycle: the spec is
  // written, loaded back into a fresh model, and re-emitted. Lets a case assert that content the
  // writer materializes (e.g. a table's totals row) survives a round-trip unchanged — neither dropped
  // on read nor duplicated/clobbered when the reloaded model is written again.
  roundtripInspectPackage(spec: Untyped) {
    return packageFacts(spec, partMapOf(writeXlsx(readXlsx(writeXlsx(buildFrom(spec))))));
  },

  // Author a pivot table over source data containing XML-special characters (& < > " ') and a
  // missing field value, write, and report whether the emitted pivotCacheDefinition is well-formed
  // and free of raw unescaped ampersands. Mirrors the oracle's shape → { ok, writeError,
  // cacheWellFormed, hasRawUnescapedAmp }.
  pivotCacheSpecialCharsReport() {
    try {
      const wb = new Workbook();
      const src = wb.addWorksheet('Data');
      src.getCell('A1').value = 'Name';
      src.getCell('B1').value = 'Region';
      src.getCell('C1').value = 'Amount';
      src.getCell('A2').value = 'Smith & Co';
      src.getCell('B2').value = '<West>';
      src.getCell('C2').value = 10;
      src.getCell('B3').value = 'East';
      src.getCell('C3').value = 20;
      src.getCell('A4').value = 'It\'s "best"';
      src.getCell('B4').value = 'West';
      src.getCell('C4').value = 30;
      wb.addWorksheet('Pivot').addPivotTable({
        source: src,
        rows: ['Name'],
        columns: ['Region'],
        values: ['Amount'],
        metric: 'sum',
      });
      const parts = partMapOf(writeXlsx(wb));
      const key = Object.keys(parts).find((n) => /pivotCacheDefinition\d*\.xml$/.test(n));
      const cacheXml = key ? parts[key]! : '';
      const hasRawUnescapedAmp = /&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/.test(
        cacheXml,
      );
      return {
        ok: true,
        writeError: null,
        cacheWellFormed: cacheXml ? !hasRawUnescapedAmp : false,
        hasRawUnescapedAmp,
      };
    } catch (e) {
      return {
        ok: false,
        writeError: messageOf(e),
        cacheWellFormed: null,
        hasRawUnescapedAmp: null,
      };
    }
  },

  // Set an autofilter over a range, write, and report the sheet's autoFilter ref plus whether the
  // workbook declares the hidden, sheet-scoped `_xlnm._FilterDatabase` defined name portable consumers
  // (LibreOffice) rely on → { autoFilterRef, hasFilterDatabase, filterDatabaseHidden,
  // filterDatabaseFormula }. Mirrors the oracle's shape.
  autoFilterDefinedNameReport(ref = 'A1:B2') {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'H1';
    sheet.getCell('B1').value = 'H2';
    sheet.getCell('A2').value = 1;
    sheet.getCell('B2').value = 2;
    sheet.autoFilter = ref;
    const parts = partMapOf(writeXlsx(workbook));
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const wbXml = parts['xl/workbook.xml'] || '';
    const autoFilterRef = (sheetXml.match(/<autoFilter\b[^>]*ref="([^"]*)"/) || [])[1] ?? null;
    const filterDb = wbXml.match(
      /<definedName\b([^>]*)name="_xlnm._FilterDatabase"([^>]*)>([\s\S]*?)<\/definedName>/,
    );
    return {
      autoFilterRef,
      hasFilterDatabase: !!filterDb,
      filterDatabaseHidden: filterDb ? /hidden="1"/.test(filterDb[1]! + filterDb[2]!) : false,
      filterDatabaseFormula: filterDb ? filterDb[3] : null,
    };
  },

  // Write one string cell with the buffered writer's useSharedStrings option and report how it was
  // stored → { hasSharedStringsPart, isSharedRef, isInline }. The option must actually control
  // storage: enabled emits a sharedStrings part and a t="s" cell reference; disabled keeps the string
  // inline with no such part.
  sharedStringsOption(useSharedStrings: boolean) {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'shared-me';
    const parts = partMapOf(writeXlsx(wb, {useSharedStrings}));
    const sheet = parts['xl/worksheets/sheet1.xml'] || '';
    return {
      hasSharedStringsPart: 'xl/sharedStrings.xml' in parts,
      isSharedRef: /t="s"><v>\d+<\/v>/.test(sheet),
      isInline: /t="inlineStr"><is><t>shared-me<\/t>/.test(sheet),
    };
  },

  // Read a fixture, write it back unchanged, and report package-part facts before/after →
  // { source, rewritten } — for asserting a no-op round-trip PRESERVES parts the reader does not
  // model (a vector-shape drawing, a header/footer image and its VML) instead of dropping them.
  roundtripFixturePackageParts(rel: string) {
    const source = packagePartFacts(partMapOf(fixtureBytes(rel)));
    const rewritten = packagePartFacts(partMapOf(writeXlsx(readXlsx(fixtureBytes(rel)))));
    return {source, rewritten};
  },

  // Write a non-finite numeric cell (NaN / Infinity / -Infinity) and report whether the sheet XML
  // carries a bare token in a <v> → { hasNonFiniteToken, token }. A non-finite value has no OOXML
  // representation, so it must serialize as a valueless cell, never a literal "NaN"/"Infinity".
  nonFiniteCellReport(kind: 'NaN' | 'Infinity' | '-Infinity') {
    const value =
      kind === 'NaN'
        ? Number.NaN
        : kind === '-Infinity'
          ? Number.NEGATIVE_INFINITY
          : Number.POSITIVE_INFINITY;
    const workbook = new Workbook();
    workbook.addWorksheet('S').getCell('A1').value = value;
    const sheetXml = partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '';
    const token = (sheetXml.match(/<c r="A1"[^>]*>\s*<v>([\s\S]*?)<\/v>/) || [])[1] ?? null;
    return {hasNonFiniteToken: /<v>[^<]*(NaN|Infinity)[^<]*<\/v>/.test(sheetXml), token};
  },

  // Read a fixture, write it back, and parse the requested cells straight from the re-emitted sheet
  // XML → { hasNaNToken, cells }. Each cell is { t, formula, value } read off the raw `<c>`. Guards
  // that a string-typed formula result under a date format is not coerced to a numeric/NaN cell.
  roundtripFixtureCellXml(rel: string, refs: string[] = []) {
    const parts = partMapOf(writeXlsx(readXlsx(fixtureBytes(rel))));
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const cells: Record<string, Untyped> = {};
    for (const ref of refs) {
      const match = sheetXml.match(new RegExp(`<c r="${ref}"([^>]*)>([\\s\\S]*?)</c>`));
      if (!match) continue;
      const t = (match[1]!.match(/\bt="([^"]*)"/) || [])[1] ?? null;
      const formula = (match[2]!.match(/<f[^>]*>([\s\S]*?)<\/f>/) || [])[1] ?? null;
      const rawValue = (match[2]!.match(/<v>([\s\S]*?)<\/v>/) || [])[1] ?? null;
      // A t="str" (or shared-string) cell holds text; anything else with a bare <v> is numeric.
      const value = rawValue === null ? null : t === 'str' ? rawValue : Number(rawValue);
      cells[ref] = {t, formula, value};
    }
    return {hasNaNToken: /<v>[^<]*NaN[^<]*<\/v>/.test(sheetXml), cells};
  },

  // Add a sheet, then probe name lookup and uniqueness for case-consistency → { foundExact,
  // foundVariant, addVariantThrew }. Lookup and add must agree on identity: a case-variant name is
  // found by getWorksheet AND rejected by addWorksheet (both case-insensitive), so no absent-yet-
  // unaddable surprise exists.
  worksheetNameLookupReport() {
    const workbook = new Workbook();
    workbook.addWorksheet('Sheet');
    const foundExact = workbook.getWorksheet('Sheet') !== undefined;
    const foundVariant = workbook.getWorksheet('sheet') !== undefined;
    let addVariantThrew = false;
    try {
      workbook.addWorksheet('sheet');
    } catch {
      addVariantThrew = true;
    }
    return {foundExact, foundVariant, addVariantThrew};
  },

  // Read a real fixture `.xlsx` and report only whether it loaded, any error, its sheet names, and
  // a couple of core properties → { ok, error, sheetNames, lastModifiedBy, creator }. The read error
  // is captured and returned as data (never propagated) so a case asserts on a crash rather than the
  // runner blowing up. Exercises the reader against foreign generators and schema-valid corners Excel
  // never emits (namespace-prefixed roots, a leading BOM, unusual part order, missing optional parts).
  readFixtureReport(rel: string) {
    try {
      const wb = readFixture(rel);
      return {
        ok: true,
        error: null,
        sheetNames: wb.worksheets.map((s) => s.name),
        lastModifiedBy: wb.properties.lastModifiedBy ?? null,
        creator: wb.properties.creator ?? null,
      };
    } catch (e) {
      return {ok: false, error: messageOf(e), sheetNames: null};
    }
  },

  // Author a sheet named "History" (a name Excel blocks in its UI but that is a valid OOXML name)
  // through the public API, round-trip it, and separately confirm a genuinely-illegal name is still
  // rejected → { addThrew, addError, roundtripName, invalidRejected }. The reserved-name UI nicety
  // is not a document-model rule, so construction must not throw; the character-legality guard is a
  // real rule and must stay.
  addReservedSheetNameReport() {
    const wb = new Workbook();
    let addThrew = false;
    let addError = null;
    try {
      wb.addWorksheet('History');
    } catch (e) {
      addThrew = true;
      addError = messageOf(e);
    }
    const roundtrip = readXlsx(writeXlsx(wb));
    const roundtripName =
      roundtrip.worksheets.map((s) => s.name).find((n) => n === 'History') ?? null;
    let invalidRejected = false;
    try {
      wb.addWorksheet('a/b');
    } catch {
      invalidRejected = true;
    }
    return {addThrew, addError, roundtripName, invalidRejected};
  },

  // Author a workbook whose sheets carry each visibility state (a valid workbook keeps one visible),
  // then report the state read back after a round-trip and the state attribute the workbook.xml
  // sheet-list entry carries → { readStates, xmlStates }, each keyed by sheet name. veryHidden — a
  // format-only state — must survive both as the model state and as the sheet-list attribute, never
  // degrading to hidden or visible.
  worksheetStateReport() {
    const wb = new Workbook();
    wb.addWorksheet('Visible');
    wb.addWorksheet('Hid', {state: 'hidden'});
    wb.addWorksheet('VeryHid', {state: 'veryHidden'});
    const buffer = writeXlsx(wb);
    const readStates: Record<string, Untyped> = {};
    for (const sheet of readXlsx(buffer).worksheets) readStates[sheet.name] = sheet.state;
    const workbookXml = partMapOf(buffer)['xl/workbook.xml'] ?? '';
    const xmlStates: Record<string, Untyped> = {};
    for (const m of workbookXml.matchAll(/<sheet\b([^>]*)\/?>/g)) {
      const attrs = m[1]!;
      const name = attrs.match(/\bname="([^"]*)"/)?.[1];
      if (name === undefined) continue;
      xmlStates[name] = attrs.match(/\bstate="([^"]*)"/)?.[1] ?? 'visible';
    }
    return {readStates, xmlStates};
  },

  // Read a fixture and report its defined names as { <name>: [refersTo…] }, mirroring the oracle.
  // The model retains every name as its own entry rather than keying by name, so two same-named
  // names scoped to different sheets both survive — the scope collision that drops one on the
  // oracle's name-keyed reader.
  readFixtureDefinedNames(rel: string) {
    const wb = readFixture(rel);
    const names: Record<string, Untyped> = {};
    for (const dn of wb.definedNames) {
      names[dn.name] ||= [];
      names[dn.name].push(dn.refersTo);
    }
    for (const k of Object.keys(names)) names[k]!.sort();
    return {names, count: Object.keys(names).length, modelCount: wb.definedNames.length};
  },

  // Read a real fixture `.xlsx` and report the fill and font colour the reader surfaces for each
  // requested `<sheet>!<address>` cell → { [key]: { fill, fontColor } | null }. Mirrors the oracle:
  // a solid-pattern fill's visible colour lives on fgColor while bgColor is the automatic indexed
  // placeholder, and the font colour is a wholly separate facet — the two are never conflated.
  // Read a real fixture `.xlsx` and report each requested cell's observable type, value, number
  // format, and note → { <addr>: {type, value, numFmt, note} | null }, on the first sheet. Mirrors
  // the oracle: a date-formatted numeric serial surfaces as a Date (value { date: iso }), not a raw
  // number, honouring the 1900 date-system leap-year quirk. `type` is a stable label.
  readFixtureCells(rel: string, cells: Untyped = []) {
    const wb = readFixture(rel);
    const sheet = wb.worksheets[0];
    const out: Record<string, Untyped> = {};
    for (const addr of cells) {
      const cell = sheet ? sheet.getCell(addr) : null;
      out[addr] = cell
        ? {
            type: cell.type,
            value: normalizeStreamValue(cell.value),
            numFmt: cell.numFmt ?? null,
            note: cell.note !== undefined ? cell.note : undefined,
          }
        : null;
    }
    return out;
  },

  // ── Streaming reader (readWorkbookStream) ──────────────────────────────────────────────────────
  // The corpus's streaming-read contract, bound to the rewrite's generator-based reader. Each method
  // mirrors its oracle sibling in workbook-io.mjs; where a case compares the streaming path to the
  // eager one, BOTH come from the rewrite, so the assertion checks that streaming and buffered reads
  // agree cell-for-cell. The rewrite's reader is a synchronous generator, so the "without race" and
  // "chunk boundary" hazards the ExcelJS stream reader faces are structurally absent.

  // Read a real styled template, write it straight back out, read the result, and report whether
  // its sheet names, custom column widths, and per-cell fill/font/numFmt/alignment/border survived
  // that no-op read→write→read → { sheetNames(Before), columns(Before), styleSurvival }. This is the
  // mainstream "open a styled template, fill it in, save it" path, which must be format-preserving.
  // Style comparison is key-order-insensitive so a case asserts on content survival, not
  // serialization incidentals. In the rewrite's model a column stores a width only when it is a
  // custom width, so "has a width" is exactly "is a custom width".
  roundtripFixture(rel: string) {
    const before = readFixture(rel);
    const after = readXlsx(writeXlsx(before));

    const hasStyle = (cell: Untyped) =>
      !!(cell.numFmt || cell.fill?.type || cell.font || cell.alignment || cell.border);
    const styleKey = (cell: Untyped) =>
      JSON.stringify(
        canonicalJson({
          numFmt: cell.numFmt || null,
          fill: cell.fill?.type ? cell.fill : null,
          font: cell.font || null,
          alignment: cell.alignment || null,
          border: cell.border || null,
        }),
      );
    const columnsWithWidth = (wb: Untyped) => {
      const out: Record<string, Untyped> = {};
      for (const sheet of wb.worksheets) {
        const cols: Record<string, Untyped> = {};
        for (const {index, properties} of sheet.columns()) {
          if (properties.width !== undefined)
            cols[index] = {width: properties.width, customWidth: true};
        }
        out[sheet.name] = cols;
      }
      return out;
    };

    let checked = 0;
    let mismatches = 0;
    let sample = null;
    for (const sheet of before.worksheets) {
      const other = after.getWorksheet(sheet.name);
      for (const {cells} of sheet.rows()) {
        for (const cell of cells) {
          // Resolve both sides through getCell so a merged-range slave redirects to its master on
          // each — comparing the row-iterated slave (its own style) against getCell (the master) would
          // report a phantom drift that is only an access asymmetry, not a lost style.
          const beforeCell = sheet.getCell(cell.address);
          if (!hasStyle(beforeCell)) continue;
          checked += 1;
          const beforeKey = styleKey(beforeCell);
          const afterKey = other ? styleKey(other.getCell(cell.address)) : '(sheet missing)';
          if (beforeKey !== afterKey) {
            mismatches += 1;
            if (!sample)
              sample = {cell: `${sheet.name}!${cell.address}`, before: beforeKey, after: afterKey};
          }
        }
      }
    }

    return {
      sheetNamesBefore: before.worksheets.map((s) => s.name),
      sheetNames: after.worksheets.map((s) => s.name),
      columnsBefore: columnsWithWidth(before),
      columns: columnsWithWidth(after),
      styleSurvival: {checked, mismatches, sample},
    };
  },

  // Build → write → read back through the rewrite's own reader, then normalize to the
  // same JSON model current.mjs reports, so every write→read round-trip case runs
  // unchanged. Facets the writer/reader do not materialize yet come back empty/null;
  // the writer's feature-gate keeps a case whose spec needs those from ever running here.
  roundtripWorkbook(spec: Untyped) {
    const reloaded = readXlsx(writeXlsx(buildFrom(spec)));
    const sheets: Record<string, Untyped> = {};
    for (const s of spec.sheets || []) {
      const sheet = reloaded.getWorksheet(s.name);
      if (!sheet) {
        sheets[s.name] = null;
        continue;
      }
      const cells: Record<string, Untyped> = {};
      for (const c of s.cells || []) cells[c.ref] = normalizeRewriteCell(sheet.getCell(c.ref));
      const columns: Record<string, Untyped> = {};
      for (const col of s.columns || []) {
        const p = sheet.getColumn(col.index);
        columns[col.index] = {width: p.width ?? null, hidden: !!p.hidden, numFmt: p.numFmt ?? null};
      }
      const rows: Record<string, Untyped> = {};
      for (const row of s.rows || []) {
        const p = sheet.getRow(row.index);
        rows[row.index] = {height: p.height ?? null, hidden: !!p.hidden};
      }
      const margins = Object.keys(sheet.pageMargins).length > 0 ? {...sheet.pageMargins} : null;
      const ps = sheet.pageSetup;
      sheets[s.name] = {
        cells,
        columns,
        rows,
        margins,
        pageSetup: {
          fitToPage: !!ps.fitToPage,
          fitToWidth: ps.fitToWidth ?? null,
          fitToHeight: ps.fitToHeight ?? null,
          scale: ps.scale ?? null,
          paperSize: ps.paperSize ?? null,
        },
        autoFilter: sheet.autoFilter?.ref ?? null,
        merges: [...sheet.merges],
        rowCount: sheet.rowCount,
        actualRowCount: sheet.actualRowCount,
      };
    }
    const props = reloaded.properties;
    const definedNames: Record<string, Untyped> = {};
    for (const dn of reloaded.definedNames) {
      definedNames[dn.name] ||= [];
      definedNames[dn.name].push(dn.refersTo);
    }
    for (const k of Object.keys(definedNames)) definedNames[k]!.sort();
    return {
      properties: {
        creator: props.creator ?? null,
        lastModifiedBy: props.lastModifiedBy ?? null,
        created: isoOrNull(props.created),
        modified: isoOrNull(props.modified),
      },
      sheets,
      definedNames,
    };
  },

  // Load a fixture and try to write it back → { loadOk, loadError, writeOk, writeError, sheetNames } —
  // for asserting a foreign construct round-trips without the writer crashing.
  roundtripFixtureWriteReport(rel: string) {
    let workbook: Untyped;
    try {
      workbook = readFixture(rel);
    } catch (e) {
      return {
        loadOk: false,
        loadError: messageOf(e),
        writeOk: false,
        writeError: null,
        sheetNames: [],
      };
    }
    let writeOk = false;
    let writeError = null;
    try {
      writeXlsx(workbook);
      writeOk = true;
    } catch (e) {
      writeError = messageOf(e);
    }
    return {
      loadOk: true,
      loadError: null,
      writeOk,
      writeError,
      sheetNames: workbook.worksheets.map((w: Untyped) => w.name),
    };
  },

  // Clone a worksheet through its model export/import: build a source sheet with cells and merges,
  // read its model, and assign that model onto a fresh sheet. Reports the merges the source model
  // exposed and the merges the destination carries afterwards → { srcMerges, dstMerges, error }.
  // The historical bug this measures is an asymmetric model contract that dropped merges on import;
  // the rewrite's getter and setter cover the same fields, so the round-trip is lossless.
  copyWorksheetModel({merges = ['A1:C1'], cells = [{ref: 'A1', value: 'merged'}]}: Untyped = {}) {
    const workbook = new Workbook();
    const src = workbook.addWorksheet('Src');
    for (const c of cells) src.getCell(c.ref).value = c.value;
    for (const m of merges) src.mergeCells(m);
    const dst = workbook.addWorksheet('Dst');

    let error = null;
    let dstMerges: Untyped[] = [];
    const srcMerges = [...src.model.merges];
    try {
      dst.model = {...src.model, name: 'Dst'} as Untyped;
      dstMerges = [...dst.model.merges];
    } catch (e) {
      error = messageOf(e);
    }
    return {srcMerges: srcMerges.sort(), dstMerges: dstMerges.sort(), error};
  },

  mutateWorksheet({cells = [], ops = [], read = [], readStyles = []}: Untyped = {}) {
    const sheet = new Workbook().addWorksheet('S');
    for (const c of cells) {
      const cell = sheet.getCell(c.ref);
      cell.value = c.value;
      // Optional per-cell style so a case can assert a structural edit carries a cell's style to its
      // shifted position rather than blanking it.
      if (c.font) cell.font = c.font;
      if (c.fill) cell.fill = c.fill;
      if (c.numFmt) cell.numFmt = c.numFmt;
    }

    let error = null;
    try {
      for (const op of ops) {
        const inserts = op.inserts || [];
        if (op.op === 'spliceRows') sheet.spliceRows(op.start, op.count, ...inserts);
        else if (op.op === 'spliceColumns') sheet.spliceColumns(op.start, op.count, ...inserts);
        else if (op.op === 'mergeCells') sheet.mergeCells(op.range);
        else if (op.op === 'insertRow') sheet.insertRow(op.pos, op.value || []);
        else if (op.op === 'duplicateRow')
          sheet.duplicateRow(op.start, {count: op.count ?? 1, insert: op.insert !== false});
        else throw new Error(`unknown mutation op: ${op.op}`);
      }
    } catch (e) {
      error = messageOf(e);
    }

    const readCells: Record<string, Untyped> = {};
    for (const ref of read) readCells[ref] = sheet.getCell(ref).value ?? null;

    // Per-cell style facets after the mutations — for asserting the style a cell carried before a
    // splice still describes the (possibly shifted) cell afterward, rather than being lost.
    const styles: Record<string, Untyped> = {};
    for (const ref of readStyles) {
      const cell = sheet.getCell(ref);
      styles[ref] = {
        value: cell.value ?? null,
        font: cell.font ? JSON.parse(JSON.stringify(cell.font)) : null,
        fill: cell.fill?.type ? JSON.parse(JSON.stringify(cell.fill)) : null,
        numFmt: cell.numFmt ?? null,
      };
    }

    // The last POPULATED row and its column-1 value, derived from the row iterator (ascending, so
    // the final populated row wins) — a delete-splice must leave this on the true last row, never a
    // trailing empty slot.
    let lastRow = null;
    for (const {number, cells: rowCells} of sheet.rows()) {
      if (rowCells.some((c) => c.value !== null && c.value !== undefined)) {
        const first = rowCells.find((c) => c.col === 1);
        lastRow = {number, value: first?.value ?? null};
      }
    }

    return {
      rowCount: sheet.rowCount,
      columnCount: sheet.columnCount,
      cells: readCells,
      styles,
      merges: [...sheet.merges],
      lastRow,
      error,
    };
  },

  tryWriteWorkbook(spec: Untyped) {
    let workbook: Untyped;
    try {
      workbook = buildFrom(spec);
    } catch (error) {
      if (error instanceof UnsupportedSpecError) throw error;
      return {ok: false, phase: 'build', error: messageOf(error)};
    }
    let buffer: Uint8Array;
    try {
      buffer = writeXlsx(workbook);
    } catch (error) {
      if (error instanceof UnsupportedSpecError) throw error;
      return {ok: false, phase: 'write', error: messageOf(error)};
    }
    // Report which cells survived the round-trip, so a case can prove a bad cell (e.g. an Invalid
    // Date, written value-less) did not drop its siblings.
    const reread = readXlsx(buffer);
    const survivingCells: Record<string, Untyped> = {};
    for (const s of spec.sheets || []) {
      const sheet = reread.getWorksheet(s.name);
      survivingCells[s.name] = (s.cells || [])
        .filter((c: Untyped) => sheet && sheet.getCell(c.ref).value !== null)
        .map((c: Untyped) => c.ref);
    }
    return {ok: true, byteLength: buffer.byteLength ?? buffer.length, survivingCells};
  },
};
