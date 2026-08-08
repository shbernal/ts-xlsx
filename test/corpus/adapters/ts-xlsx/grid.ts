import {messageOf} from '../../thrown.ts';
// Rows, columns, merges and the sheet geometry around them — insertion, splicing, outline
// levels, freeze panes, print areas and page breaks.

import type {RowInput} from '../../../../src/core/worksheet.ts';
import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {
  decodeRange,
  encodeAddress,
  fixtureBytes,
  readFixture,
  readXlsx,
  Workbook,
  writeXlsx,
} from './runtime.ts';
import {buildFrom, normalizeStreamValue, ONE_PX_PNG, printAreaRefersTo} from './spec-model.ts';
import {reloadPatched} from './xml-probes.ts';

export const grid = {
  cellColRowTypes(ref = 'B3') {
    const sheet = new Workbook().addWorksheet('S');
    const cell = sheet.getCell(ref);
    cell.value = 'x';
    return {col: cell.col, row: cell.row, colType: typeof cell.col, rowType: typeof cell.row};
  },

  // Author a sheet, round-trip, load, append more rows after the last populated row, round-trip again →
  // { loadedRowCount, finalRowCount, rows }. The load-bearing fact: a reloaded sheet reports its last
  // populated row so addRow lands new content at N+1 with no gap or overwrite.
  appendRowsAfterReload(initial: RowInput[] = [], append: RowInput[] = []) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const row of initial) sheet.addRow(row);

    const loaded = readXlsx(writeXlsx(workbook));
    const s = loaded.getWorksheet('S');
    const loadedRowCount = s!.rowCount;
    for (const row of append) s!.addRow(row);

    const final = readXlsx(writeXlsx(loaded));
    const f = final.getWorksheet('S');
    // Mirror the oracle's `row.values.slice(1)` per-row array: each row is sized to its own populated
    // extent, holes are null, and an empty row is an empty array — indexed by row number so a gap shows.
    const rows: Untyped[] = Array.from({length: f!.rowCount}, () => []);
    for (const {number, cells} of f!.rows()) {
      const maxCol = cells.reduce((m: number, c: Untyped) => Math.max(m, c.col), 0);
      const arr = new Array(maxCol).fill(null);
      for (const cell of cells) arr[cell.col - 1] = normalizeStreamValue(cell.value);
      rows[number - 1] = arr;
    }
    return {loadedRowCount, finalRowCount: f!.rowCount, rows};
  },

  // Read a fixture's single _xlnm.Print_Area name (a comma-separated range list), re-write it, and read
  // it again → { sourceRangeCount, readPrintArea, rewrittenRangeCount }. Both disjoint ranges must
  // survive read and re-serialization, never truncated to the first.
  roundtripFixturePrintAreas(rel: Untyped) {
    const printAreaOf = (wb: Untyped) =>
      wb.definedNames.find((n: Untyped) => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const source = readXlsx(fixtureBytes(rel));
    const readPrintArea = printAreaOf(source);
    const sourceRangeCount = readPrintArea.split(',').filter(Boolean).length;
    const rewritten = readXlsx(writeXlsx(source));
    const rewrittenRangeCount = printAreaOf(rewritten).split(',').filter(Boolean).length;
    return {sourceRangeCount, readPrintArea, rewrittenRangeCount};
  },

  // Author a sheet-scoped _xlnm.Print_Area over a comma-separated area, round-trip, and report the
  // emitted ranges (sheet prefix stripped) → { ranges }. Two disjoint areas must emit two proper
  // rectangular ranges in one name, not a truncated single range.
  writePrintAreaDefinedName(area: Untyped) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    workbook.defineName({
      name: '_xlnm.Print_Area',
      scope: sheet.name,
      refersTo: printAreaRefersTo(sheet.name, area),
    });
    const back = readXlsx(writeXlsx(workbook));
    const refersTo = back.definedNames.find((n) => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const ranges = refersTo.split(',').map((r) => r.split('!').pop());
    return {ranges};
  },

  // Author a sheet-scoped _xlnm.Print_Area over one area (whole-column or bounded), round-trip, and
  // report the written and recovered forms → { writtenDefinedName, reReadPrintArea, reloadOk }. A
  // column-only reference ($A:$D) must recover intact, never decoded to a NaN-mangled address.
  printAreaRoundtrip(area: Untyped) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    workbook.defineName({
      name: '_xlnm.Print_Area',
      scope: sheet.name,
      refersTo: printAreaRefersTo(sheet.name, area),
    });
    let reloadOk = true;
    let back: Untyped;
    try {
      back = readXlsx(writeXlsx(workbook));
    } catch {
      reloadOk = false;
    }
    const refersTo =
      back?.definedNames.find((n: Untyped) => n.name === '_xlnm.Print_Area')?.refersTo ?? '';
    const reReadPrintArea = refersTo.split('!').pop()?.replace(/\$/g, '') ?? '';
    return {writtenDefinedName: refersTo, reReadPrintArea, reloadOk};
  },

  // Freeze the first row, write, and report the emitted pane plus a round-trip → { paneEmitted,
  // reReadState, reReadYSplit, reReadXSplit }. A frozen-header view serializes a <pane> and reloads
  // as a frozen split of one row and no columns.
  frozenTopRowRoundtrip() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.freeze(1);
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const paneEmitted = /<pane\b[^>]*ySplit="1"[^>]*state="frozen"/.test(sheetXml);
    const view = readXlsx(buffer).getWorksheet('S')!.view;
    return {
      paneEmitted,
      reReadState: view.state ?? 'normal',
      reReadYSplit: view.ySplit ?? 0,
      reReadXSplit: view.xSplit ?? 0,
    };
  },

  // Author the shape a generated report has — a frozen top row above grouped, hidden columns, on the
  // first of two sheets — write it, and report the view-initialisation facts of the written package →
  // { bookViewCount, bookViewsBeforeSheets, windowWidth, windowHeight, selectedSheets,
  // outlineLevelCol, paneEmitted }. A consumer restores the document window from the workbook view
  // and lays the frozen pane and the column-outline bar out inside it; omitting those facts leaves
  // the top pane measured against an uninitialised window.
  frozenHeaderViewStateReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Report');
    sheet.getCell('A1').value = 'header';
    sheet.getCell('A2').value = 'body';
    for (const index of [2, 3]) {
      sheet.getColumn(index).outlineLevel = 1;
      sheet.getColumn(index).hidden = true;
    }
    sheet.freeze(1);
    wb.addWorksheet('Notes').getCell('A1').value = 'notes';

    const parts = partMapOf(writeXlsx(wb));
    const workbookXml = parts['xl/workbook.xml'] || '';
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const bookView = (workbookXml.match(/<workbookView\b[^>]*\/>/) || [''])[0] || '';
    const attrNumber = (tag: string, name: string): number | null => {
      const found = tag.match(new RegExp(`\\b${name}="(-?\\d+(?:\\.\\d+)?)"`));
      return found ? Number(found[1]) : null;
    };
    const selectedSheets = ['Report', 'Notes'].filter((_, i) =>
      /<sheetView\b[^>]*\btabSelected="1"/.test(parts[`xl/worksheets/sheet${i + 1}.xml`] || ''),
    );
    const sheetFormat = (sheetXml.match(/<sheetFormatPr\b[^>]*\/>/) || [''])[0] || '';
    return {
      bookViewCount: (workbookXml.match(/<workbookView\b/g) || []).length,
      bookViewsBeforeSheets:
        workbookXml.indexOf('<bookViews>') >= 0 &&
        workbookXml.indexOf('<bookViews>') < workbookXml.indexOf('<sheets>'),
      windowWidth: attrNumber(bookView, 'windowWidth'),
      windowHeight: attrNumber(bookView, 'windowHeight'),
      selectedSheets,
      outlineLevelCol: attrNumber(sheetFormat, 'outlineLevelCol'),
      outlineLevelRow: attrNumber(sheetFormat, 'outlineLevelRow'),
      paneEmitted: /<pane\b[^>]*ySplit="1"[^>]*state="frozen"/.test(sheetXml),
    };
  },

  // Freeze a view, then unfreeze it, and report the pane presence in each written form plus a reload
  // → { frozenHasPane, normalHasPane, reloadedState, reloadedHasSplit }. Unfreezing must leave no
  // leftover <pane> (which triggers Excel's repair prompt) and reload as a normal, unsplit view.
  unfreezeViewRoundtrip() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.freeze(1);
    const frozenHasPane = /<pane\b/.test(
      partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '',
    );

    sheet.unfreeze();
    const normalBuffer = writeXlsx(workbook);
    const normalHasPane = /<pane\b/.test(partMapOf(normalBuffer)['xl/worksheets/sheet1.xml'] || '');
    const view = readXlsx(normalBuffer).getWorksheet('S')!.view;
    return {
      frozenHasPane,
      normalHasPane,
      reloadedState: view.state ?? 'normal',
      reloadedHasSplit: (view.xSplit ?? 0) > 0 || (view.ySplit ?? 0) > 0,
    };
  },

  // Author three columns with distinct widths (one hidden), write, then REVERSE the order of the
  // emitted `<col>` tags — the shape foreign generators (excelize, jxls-poi) produce — and read the
  // patched package back → { w1, w2, w3, hidden2 }. Each column's width and hidden flag must bind to
  // the column its min/max names, regardless of document order.
  outOfOrderColumnsReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getColumn(1).width = 10;
    sheet.getColumn(2).width = 20;
    sheet.getColumn(2).hidden = true;
    sheet.getColumn(3).width = 30;
    const back = reloadPatched(writeXlsx(wb), {
      'xl/worksheets/sheet1.xml': (xml) =>
        xml.replace(/<cols>([\s\S]*?)<\/cols>/, (_, inner) => {
          const tags = inner.match(/<col\b[^>]*\/>/g) || [];
          return `<cols>${tags.reverse().join('')}</cols>`;
        }),
    }).getWorksheet('S')!;
    return {
      w1: back.getColumn(1).width,
      w2: back.getColumn(2).width,
      w3: back.getColumn(3).width,
      hidden2: back.getColumn(2).hidden ?? false,
    };
  },

  // Assign an outline (grouping) level to a row and a column, write, and read back → { rowOutline,
  // colOutline }. The OOXML outlineLevel attribute on <row>/<col> must survive the round-trip on both
  // axes so a collapsible grouping is preserved on reopen.
  rowColumnOutlineLevelRoundtrip() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'x';
    sheet.getRow(2).outlineLevel = 1;
    sheet.getColumn(3).outlineLevel = 1;
    const back = readXlsx(writeXlsx(wb)).getWorksheet('S')!;
    return {
      rowOutline: back.getRow(2).outlineLevel ?? 0,
      colOutline: back.getColumn(3).outlineLevel ?? 0,
    };
  },

  // Set explicit widths on three columns (one of which coincides with the format's conventional
  // default width, 9), write, and report per-column whether a `<col>` with a customWidth flag was
  // emitted and what width each reads back → { emitted: {c1,c2,c3}, readBack: {c1,c2,c3} }. An
  // explicitly-set width must survive even when its value equals the magic default.
  columnWidthDefaultCollisionReport(widths = [8, 9, 10]) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    widths.forEach((w, i) => {
      sheet.getColumn(i + 1).width = w;
    });
    const buffer = writeXlsx(wb);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    // A column emits an explicit width when its `<col>` carries a customWidth flag over its index.
    const emittedAt = (index: Untyped) =>
      new RegExp(`<col\\b[^>]*\\bmin="${index}"[^>]*\\bmax="${index}"[^>]*\\bcustomWidth="1"`).test(
        sheetXml,
      );
    const back = readXlsx(buffer).getWorksheet('S')!;
    return {
      emitted: {c1: emittedAt(1), c2: emittedAt(2), c3: emittedAt(3)},
      readBack: {
        c1: back.getColumn(1).width,
        c2: back.getColumn(2).width,
        c3: back.getColumn(3).width,
      },
    };
  },

  // Append rows in every shape (dense array, sparse array, keyed object, mixed batch), round-trip, and
  // read them back letter-keyed by row number → { rows }. Column keys bind object values to columns;
  // dense/sparse arrays map positionally with holes left empty; a numeric/date value survives typed.
  appendRowShapes() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getColumn(1).key = 'k1';
    sheet.getColumn(2).key = 'k2';
    sheet.addRow(['header']); // row 1 — keeps the checked rows at their stated numbers
    sheet.addRow(['a', 'b', 'c']); // row 2 — dense positional array
    // biome-ignore lint/suspicious/noSparseArray: a genuine array hole (not undefined) is the point — the gap at column B must be skipped positionally
    sheet.addRow(['x', , 'z']); // row 3 — sparse array, gap at column B
    sheet.addRow({k1: 'o1', k2: 'o2'}); // row 4 — keyed object
    sheet.addRow([7, new Date(Date.UTC(2021, 0, 2))]); // row 5 — number + date
    sheet.addRows([['m1', 'm2'], {k1: 'n1'}]); // rows 6, 7 — mixed batch

    const loaded = readXlsx(writeXlsx(workbook));
    const s = loaded.getWorksheet('S');
    const rows: Record<string, Record<string, Untyped>> = {};
    for (const {number, cells} of s!.rows()) {
      const row: Record<string, Untyped> = {};
      for (const cell of cells)
        row[encodeAddress(cell.col, number).match(/^[A-Z]+/)![0]] = normalizeStreamValue(
          cell.value,
        );
      rows[number] = row;
    }
    // Every checked column reads as null when the round-trip left it empty, so a gap is visible.
    for (const n of Object.keys(rows)) for (const col of ['A', 'B', 'C']) rows[n]![col] ??= null;
    return {rows};
  },

  // Feed addRow an array built in another realm (a vm context): Array.isArray must recognize it so its
  // elements fill columns → { isArrayCrossRealm, a, b, c }. `instanceof Array` would miss it and place
  // nothing, walking it as a keyed object instead.
  async crossRealmArrayRow() {
    const vm = await import('node:vm');
    const arr = vm.runInNewContext('[10, 20, 30]');
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.addRow(arr);
    return {
      isArrayCrossRealm: Array.isArray(arr),
      a: sheet.getCell('A1').value ?? null,
      b: sheet.getCell('B1').value ?? null,
      c: sheet.getCell('C1').value ?? null,
    };
  },

  // Set row-level properties on rows that carry NO cell value and report what survives a round-trip
  // → { row3Hidden, row4Hidden, row4Height, row5Hidden }. A content-less row bearing a hidden flag,
  // a height, or an outline level must still be written (its <row> element materialised) so the
  // property is not lost — the failure mode is a blank hidden/grouped spacer row coming back visible.
  hiddenEmptyRowReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getRow(3).hidden = true;
    ws.getRow(4).hidden = true;
    ws.getRow(4).height = 25;
    ws.getRow(5).hidden = true;
    ws.getRow(5).outlineLevel = 1;
    const rt = readXlsx(writeXlsx(wb)).getWorksheet('S')!;
    return {
      row3Hidden: rt.getRow(3).hidden ?? false,
      row4Hidden: rt.getRow(4).hidden ?? false,
      row4Height: rt.getRow(4).height ?? null,
      row5Hidden: rt.getRow(5).hidden ?? false,
    };
  },

  // Prove manual horizontal page breaks survive load and a load→save round-trip → { sourceBreaks,
  // loadedBreaks, rewrittenBreaks }, each the ascending list of break row ids. sourceBreaks reads the
  // raw fixture XML (the precondition); loadedBreaks/rewrittenBreaks come off the model after read and
  // after write→re-read, so a dropped-on-read or dropped-on-write regression shows as an empty list.
  roundtripFixtureRowBreaks(rel: Untyped) {
    const rowBreakIds = (xml: Untyped) => {
      const section = xml.match(/<rowBreaks[\s\S]*?<\/rowBreaks>/);
      if (section === null) return [];
      return [...section[0].matchAll(/<brk\b[^>]*\bid="(\d+)"/g)]
        .map((m) => Number(m[1]))
        .sort((a, b) => a - b);
    };
    const sheet1 = (parts: Untyped) => parts['xl/worksheets/sheet1.xml'] ?? '';
    const modelBreaks = (wb: Untyped) =>
      wb.worksheets[0].rowBreaks
        .map((brk: Untyped) => brk.id)
        .sort((a: Untyped, b: Untyped) => a - b);

    const sourceBreaks = rowBreakIds(sheet1(partMapOf(fixtureBytes(rel))));
    const loaded = readFixture(rel);
    const loadedBreaks = modelBreaks(loaded);
    const rewrittenBreaks = modelBreaks(readXlsx(writeXlsx(loaded)));
    return {sourceBreaks, loadedBreaks, rewrittenBreaks};
  },

  // Merge A1:B3 with values only in A1/A2, round-trip, then iterate every used-range position
  // (include-empty) and report → { rowCount, visited, a3: { visited, isMerged, master } }. A merge
  // reaching into an otherwise-empty trailing row keeps that row within the bounds, so A3 is visited
  // and resolves to its master A1 rather than being skipped.
  trailingMergedRowIterationReport() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getCell('A1').value = 'top';
    s.getCell('A2').value = 'data';
    s.mergeCells('A1:B3');
    const rs = readXlsx(writeXlsx(wb)).worksheets[0]!;
    const rects = rs.merges.map((range: Untyped) => {
      const {top, left, bottom, right} = decodeRange(range);
      return {top, left, bottom, right, masterRef: encodeAddress(left!, top!)};
    });
    const masterOf = (row: Untyped, col: Untyped) =>
      rects.find((r) => row >= r.top! && row <= r.bottom! && col >= r.left! && col <= r.right!);
    const visited: Untyped[] = [];
    for (let row = 1; row <= rs.rowCount; row++) {
      for (let col = 1; col <= rs.columnCount; col++) visited.push(encodeAddress(col, row));
    }
    const a3Rect = masterOf(3, 1);
    return {
      rowCount: rs.rowCount,
      visited,
      a3: {
        visited: visited.includes('A3'),
        isMerged: a3Rect !== undefined,
        master: a3Rect ? a3Rect.masterRef : null,
      },
    };
  },

  // Build a workbook from a spec, round-trip it, and for each requested row report the column indices
  // an include-empty iteration yields → { rows: { <n>: { cols } }, columnCount }. Positional iteration
  // walks 1..columnCount (the sheet's declared width), so interior *and* trailing empties are surfaced
  // and every row reconstructs to the header width — the alignment invariant a positional consumer needs.
  async readRowCellPresence(spec: Untyped, rowNumbers: Untyped = []) {
    const sheet = readXlsx(writeXlsx(buildFrom(spec))).worksheets[0]!;
    const columnCount = sheet.columnCount;
    const rows: Record<string, Untyped> = {};
    for (const rn of rowNumbers) {
      const cols: Untyped[] = [];
      for (let col = 1; col <= columnCount; col++) {
        sheet.getCell(encodeAddress(col, rn));
        cols.push(col);
      }
      rows[rn] = {cols};
    }
    return {rows, columnCount};
  },

  // Merge A1:B1 with the value in the master (A1), then read the display text of the master and of a
  // merged child (B1) → { masterText, childText, childThrew }. Addressing a covered cell resolves to
  // its region's master, so a merged child's text mirrors the master and never throws.
  mergedCellDisplayTextReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'Group';
    ws.mergeCells('A1:B1');
    const textOf = (ref: Untyped) => {
      const v = ws.getCell(ref).value;
      return v === null || v === undefined ? '' : String(v);
    };
    let masterText = null;
    let childText = null;
    let childThrew = false;
    try {
      masterText = textOf('A1');
    } catch {
      masterText = null;
    }
    try {
      childText = textOf('B1');
    } catch {
      childThrew = true;
    }
    return {masterText, childText, childThrew};
  },

  // Merge a horizontal span with a value + alignment on the anchor, write, then read back →
  // { mergeCount, merges, populatedCoveredCells, anchorValue, anchorAlignment }. A clean merge
  // declares the range exactly once and emits a value only on the anchor, so the covered cells
  // carry no conflicting <v> — the shape that opens without Excel's repair prompt — and the
  // anchor's value and alignment survive the round-trip.
  mergeCleanReport({anchor = 'B1', range = 'B1:G1', value = 'Group Title'}: Untyped = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell(anchor);
    cell.value = value;
    cell.alignment = {horizontal: 'center'};
    sheet.mergeCells(range);
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const merges = [...sheetXml.matchAll(/<mergeCell\b[^>]*ref="([^"]*)"/g)].map((m) => m[1]);
    const {left, right, top, bottom} = decodeRange(range);
    const populatedCoveredCells: Untyped[] = [];
    for (let r = top!; r <= bottom!; r++) {
      for (let c = left!; c <= right!; c++) {
        const ref = encodeAddress(c, r);
        if (ref === anchor) continue;
        if (new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<v>`).test(sheetXml))
          populatedCoveredCells.push(ref);
      }
    }
    const a = readXlsx(buffer).getWorksheet('S')!.getCell(anchor);
    return {
      mergeCount: merges.length,
      merges,
      populatedCoveredCells,
      anchorValue: a.value ?? null,
      anchorAlignment: a.alignment ? {...a.alignment} : null,
    };
  },

  // Populate covered non-anchor cells FIRST, then merge over them — the order that leaves stray
  // values behind. Write and read back → { anchorValue, populatedCoveredCells, coveredValuesOnRead }.
  // Excel keeps only the anchor's value on merge; a covered cell that still carries a <v> under the
  // <mergeCell> ref is the geometry that trips Excel's repair prompt.
  mergeOverPopulatedReport({anchor = 'B1', range = 'B1:G1', value = 'Group Title'}: Untyped = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const {left, right, top, bottom} = decodeRange(range);
    // Fill every cell of the range, anchor included, before the merge collapses it.
    for (let r = top!; r <= bottom!; r++) {
      for (let c = left!; c <= right!; c++) {
        sheet.getCell(encodeAddress(c, r)).value = c === left && r === top ? value : `covered-${c}`;
      }
    }
    sheet.mergeCells(range);
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const populatedCoveredCells: Untyped[] = [];
    for (let r = top!; r <= bottom!; r++) {
      for (let c = left!; c <= right!; c++) {
        const ref = encodeAddress(c, r);
        if (ref === anchor) continue;
        if (new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*>[\\s\\S]*?<v>`).test(sheetXml))
          populatedCoveredCells.push(ref);
      }
    }
    const reread = readXlsx(buffer).getWorksheet('S')!;
    return {
      anchorValue: reread.getCell(anchor).value ?? null,
      populatedCoveredCells,
    };
  },

  // Give the top-left/master cell a border (+ numFmt + font), merge it into a region, round-trip,
  // then report the master's border/numFmt/font and the declared merges → for asserting a merge
  // does not strip the style the master needs to render the merged region's outline.
  mergeMasterBorderReport() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell('A1');
    cell.value = 'x';
    cell.border = {top: {style: 'thin'}, bottom: {style: 'medium'}};
    cell.numFmt = '0.00';
    cell.font = {bold: true};
    sheet.mergeCells('A1:B2');
    const reread = readXlsx(writeXlsx(workbook)).getWorksheet('S')!;
    const m = reread.getCell('A1');
    const b = m.border || {};
    return {
      hasTopBorder: !!b.top?.style,
      hasBottomBorder: !!b.bottom?.style,
      topStyle: b.top ? (b.top.style ?? null) : null,
      bottomStyle: b.bottom ? (b.bottom.style ?? null) : null,
      numFmt: m.numFmt ?? null,
      fontBold: !!m.font?.bold,
      merges: [...reread.merges],
    };
  },

  // Merge a rectangular region (master = top-left), set a value by addressing a NON-master
  // (slave) cell inside it, write, and report which cells carry an independent <v> in the sheet
  // XML, the declared merges, and the re-read master/slave values → for asserting the slave write
  // resolves to the master (only the master carries a value; reading either address returns it).
  mergeSlaveWrite({range = 'A1:B2', slave = 'B2', value = 'slave-write'}: Untyped = {}) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.mergeCells(range);
    sheet.getCell(slave).value = value;
    const buffer = writeXlsx(workbook);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    // A cell "carries a value" if its element has value content — a number/bool/formula (<v>),
    // an inline string (<is>), or a formula (<f>). The writer serialises strings as inlineStr,
    // so keying on <v> alone would miss them; an empty covered cell is never emitted at all.
    const cellsWithValue = [
      ...sheetXml.matchAll(/<c\b[^>]*\br="([A-Z]+\d+)"[^>]*>([\s\S]*?)<\/c>/g),
    ]
      .filter((m) => /<(?:v|is|f)\b/.test(m[2]!))
      .map((m) => m[1]);
    const merges = [...sheetXml.matchAll(/<mergeCell\b[^>]*ref="([^"]*)"/g)].map((m) => m[1]);
    const master = range.split(':')[0]!;
    const s = readXlsx(buffer).getWorksheet('S')!;
    return {
      cellsWithValue,
      merges,
      masterValue: s.getCell(master).value ?? null,
      slaveValue: s.getCell(slave).value ?? null,
    };
  },

  // Duplicate a populated row with default args, then merge a range on the copy — for asserting the
  // copy is faithful (values, not empty/NaN) and carries no phantom merge that would reject the merge.
  duplicateRowReport() {
    const sheet = new Workbook().addWorksheet('S');
    sheet.getCell('A1').value = 'a';
    sheet.getCell('B1').value = 'b';
    sheet.getCell('C1').value = 'c';
    let dupError = null;
    try {
      sheet.duplicateRow(1, {count: 1, insert: true});
    } catch (e) {
      dupError = messageOf(e);
    }
    const val = (ref: Untyped) => sheet.getCell(ref).value ?? null;
    const row1 = [val('A1'), val('B1'), val('C1')];
    const row2 = [val('A2'), val('B2'), val('C2')];
    let mergeError = null;
    try {
      sheet.mergeCells('A2:C2');
    } catch (e) {
      mergeError = messageOf(e);
    }
    return {dupError, mergeError, rowCount: sheet.rowCount, row1, row2, merges: [...sheet.merges]};
  },

  // Insert a row then style a cell of it — for asserting the inserted cells stay mutable (no frozen,
  // "object is not extensible" style object) regardless of the requested style-inheritance mode. The
  // rewrite's copy-on-write style model makes every cell mutable by construction, so the mode is
  // immaterial; it is accepted and ignored.
  insertRowThenStyle(_styleMode = 'i') {
    const sheet = new Workbook().addWorksheet('S');
    sheet.getCell('A1').value = 'header';
    sheet.getCell('A1').font = {bold: true};
    sheet.getCell('A2').value = 'data';
    let error = null;
    let numFmt = null;
    try {
      sheet.insertRow(2, ['inserted']);
      const cell = sheet.getCell('A2');
      cell.numFmt = '$#,##0.00;[Red]-$#,##0.00';
      cell.font = {...cell.font, bold: true};
      numFmt = cell.numFmt;
    } catch (e) {
      error = messageOf(e);
    }
    return {error, numFmt};
  },

  // Attach a cell note and an outline level, insert a row above them, and round-trip through the real
  // writer/reader → { dataShifted, noteFollowsRow, outlineFollowsRow }. Both the note and the outline
  // level must follow their logical row through the insert and survive serialization.
  rowInsertPreservesNoteAndOutline() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'r1';
    ws.getCell('A2').value = 'r2';
    ws.getCell('A2').note = 'mynote';
    ws.getRow(2).outlineLevel = 1;
    ws.insertRow(1, ['new']); // r1 -> row 2, r2 (noted, outlined) -> row 3
    const s = readXlsx(writeXlsx(wb)).getWorksheet('S')!;
    return {
      dataShifted: s.getCell('A2').value === 'r1' && s.getCell('A3').value === 'r2',
      noteFollowsRow: !!s.getCell('A3').note,
      outlineFollowsRow: s.getRow(3).outlineLevel === 1,
    };
  },

  spliceShiftsRefs() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    // Table occupies A3:B5 (header + 2 rows); image anchored from row 5 (0-based).
    s.addTable({name: 'T', ref: 'A3', columns: [{name: 'H1'}, {name: 'H2'}], rowCount: 2});
    const id = wb.addImage({buffer: ONE_PX_PNG, extension: 'png'});
    s.addImage(id, {tl: {col: 0, row: 5}, br: {col: 2, row: 8}});
    s.spliceRows(1, 0, ['inserted']); // insert a row at the top → table and image shift down 1

    const parts = partMapOf(writeXlsx(wb));
    const tableXml = parts['xl/tables/table1.xml'] || '';
    const tableRef = (tableXml.match(/<table\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null;
    const drawingXml = parts['xl/drawings/drawing1.xml'] || '';
    const imageFromRow = (drawingXml.match(/<xdr:from>[\s\S]*?<xdr:row>(\d+)<\/xdr:row>/) || [])[1];

    // Duplicate table column names authored separately — construction disambiguates them into a
    // unique set rather than emitting a corrupt table (the same repair the reader applies on load).
    const w2 = new Workbook();
    const dupTable = w2
      .addWorksheet('S')
      .addTable({name: 'T2', ref: 'A1', columns: [{name: 'Dup'}, {name: 'Dup'}], rowCount: 1});
    writeXlsx(w2);
    const dupColumnNames = dupTable.columns.map((c) => c.name);
    const dupColumnNamesUnique =
      new Set(dupColumnNames.map((n) => n.toLowerCase())).size === dupColumnNames.length;

    return {
      tableRef,
      imageFromRow: imageFromRow != null ? Number(imageFromRow) : null,
      dupColumnNames,
      dupColumnNamesUnique,
    };
  },

  // Define four adjacent columns with identical width and outline level, write, and report whether
  // the write and reload succeed and how many <col> spans the part carries. Equivalent adjacent
  // columns must coalesce into fewer <col> spans than columns, without the collapse pass throwing.
  equivalentColumnCollapseReport() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    for (let i = 1; i <= 4; i++) {
      const col = s.getColumn(i);
      col.width = 12;
      col.outlineLevel = 1;
    }
    let writeOk = true;
    let writeError = null;
    let buffer = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = messageOf(e);
    }
    if (!writeOk) return {writeOk, writeError, reloadOk: false, colSpanCount: null};

    const parts = partMapOf(buffer!);
    const sheetPart = Object.keys(parts).find((n) => /xl\/worksheets\/sheet\d+\.xml$/.test(n));
    const colsBlock = (parts[sheetPart!]!.match(/<cols>[\s\S]*?<\/cols>/) || [])[0] ?? '';
    const colSpanCount = (colsBlock.match(/<col\b/g) || []).length;

    let reloadOk = true;
    try {
      readXlsx(buffer!);
    } catch {
      reloadOk = false;
    }
    return {writeOk, writeError, reloadOk, colSpanCount};
  },
};
