// Worksheet tables: their columns, styles, display names, and what editing one does to the
// cells underneath it.

import fs from 'node:fs';
import path from 'node:path';
import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {FIXTURES_ROOT, readFixture, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {buildFrom} from './spec-model.ts';

export const tables = {
  // Find a table by name across a loaded fixture's sheets and report its column names and data-row
  // count. The reader reconstructs the table from its part, deriving the data-row count from the
  // stored range (height minus the header and totals rows), so a loaded table exposes its rows.
  readFixtureTable(rel: Untyped, tableName: Untyped) {
    const wb = readFixture(rel);
    for (const s of wb.worksheets) {
      const table = s.tables.find((t) => t.name === tableName);
      if (table)
        return {
          found: true,
          columns: table.columns.map((c) => c.name),
          rowCount: table.options.rowCount,
        };
    }
    return {found: false, columns: null, rowCount: null};
  },

  // Load a fixture and report a named table's column count and names — used to prove a table with a
  // calculated column (a <calculatedColumnFormula> child the reader ignores) does not truncate the
  // column list or crash the read.
  loadFixtureTableColumns(rel: Untyped, tableName: Untyped) {
    try {
      const wb = readFixture(rel);
      for (const s of wb.worksheets) {
        const table = s.tables.find((t) => t.name === tableName);
        if (table) {
          return {
            loaded: true,
            error: null,
            columnCount: table.columns.length,
            columnNames: table.columns.map((c) => c.name),
          };
        }
      }
      return {loaded: true, error: null, columnCount: 0, columnNames: []};
    } catch (e) {
      return {
        loaded: false,
        error: String((e as Untyped)?.message || e),
        columnCount: null,
        columnNames: null,
      };
    }
  },

  // Build a table-bearing spec, report the full ref written into each table part, then read the
  // package back and re-write it, reporting the ref and well-formedness of each re-emitted part — so
  // a degenerate (empty-body or single-row) table is proven to survive a load→save round-trip.
  roundtripSpecTableFacts(spec: Untyped) {
    const tableFacts = (parts: Untyped) =>
      Object.keys(parts)
        .filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n))
        .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]))
        .map((n) => {
          const xml = parts[n];
          return {
            ref: (xml.match(/<table\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null,
            wellFormed: !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml),
          };
        });
    const write = tableFacts(partMapOf(writeXlsx(buildFrom(spec))));
    let loadOk = true;
    let loadError = null;
    let roundtrip: Untyped[] = [];
    try {
      const reloaded = readXlsx(writeXlsx(buildFrom(spec)));
      roundtrip = tableFacts(partMapOf(writeXlsx(reloaded)));
    } catch (e) {
      loadOk = false;
      loadError = String((e as Untyped)?.message || e);
    }
    return {write, roundtrip, loadOk, loadError};
  },

  // Author a five-column table, round-trip it, and report the loaded column count and names — the
  // reader must expose every column in order, not truncate to a fixed cap.
  wideTableColumnReadReport() {
    const wb = new Workbook();
    wb.addWorksheet('S').addTable({
      name: 'Wide',
      ref: 'A1',
      columns: [{name: 'C1'}, {name: 'C2'}, {name: 'C3'}, {name: 'C4'}, {name: 'C5'}],
      rowCount: 2,
    });
    const table = readXlsx(writeXlsx(wb)).getWorksheet('S')!.tables[0]!;
    return {colCount: table.columns.length, colNames: table.columns.map((c) => c.name)};
  },

  // Add a table and a list validation to each of five sheets, then report that the package writes
  // with unique table part ids, reloads with every table present, and keeps the first sheet's
  // validation — the loop over many sheets must not collide table ids or strip validations.
  multiSheetTableReport() {
    const wb = new Workbook();
    for (let i = 1; i <= 5; i++) {
      const s = wb.addWorksheet(`Sheet${i}`);
      s.addTable({name: `Tbl${i}`, ref: 'A1', columns: [{name: 'Col'}], rowCount: 1});
      s.addDataValidation('C1', {type: 'list', allowBlank: true, formulae: ['"a,b,c"']});
    }
    let writeOk = true;
    let writeError = null;
    let buffer = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = String((e as Untyped)?.message || e);
    }
    if (!writeOk)
      return {
        writeOk,
        writeError,
        reloadOk: false,
        tableCount: null,
        idsUnique: false,
        firstSheetDvSurvives: false,
      };

    const parts = partMapOf(buffer!);
    const ids = Object.keys(parts)
      .filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n))
      .map((n) => (parts[n]!.match(/<table\b[^>]*\bid="([^"]*)"/) || [])[1]);
    const idsUnique =
      ids.length > 0 && ids.every((x) => x != null) && new Set(ids).size === ids.length;

    let reloadOk = true;
    let tableCount = 0;
    let firstSheetDvSurvives = false;
    try {
      const back = readXlsx(buffer!);
      tableCount = back.worksheets.reduce((n, s) => n + s.tables.length, 0);
      firstSheetDvSurvives = !!back.getWorksheet('Sheet1')?.dataValidationAt('C1');
    } catch {
      reloadOk = false;
    }
    return {writeOk, writeError, reloadOk, tableCount, idsUnique, firstSheetDvSurvives};
  },

  // Write three tables — one with a real built-in style name, one with no style, one with the
  // sentinel theme "None" — and report the tableStyleInfo name attribute (or null when absent) each
  // emits, plus whether the "None" table kept its showRowStripes flag. Theme "None" must mean an
  // unstyled table (no name attribute), not a bogus name="None" referencing a non-existent style.
  tableStyleThemeReport() {
    const styleInfoOf = (style: Untyped) => {
      const wb = new Workbook();
      wb.addWorksheet('S').addTable({
        name: 'T',
        ref: 'A1',
        columns: [{name: 'A'}],
        rowCount: 1,
        style,
      });
      let ok = true;
      let tag = null;
      try {
        const parts = partMapOf(writeXlsx(wb));
        const part = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
        tag = (parts[part!]!.match(/<tableStyleInfo[^>]*\/?>/) || [])[0] ?? null;
      } catch (e) {
        ok = false;
        tag = String((e as Untyped)?.message || e);
      }
      const name = tag && ok ? ((tag.match(/\bname="([^"]*)"/) || [])[1] ?? null) : null;
      const hasStripes = !!(tag && ok && /\bshowRowStripes="1"/.test(tag));
      return {ok, name, hasStripes, tag};
    };
    return {
      real: styleInfoOf({name: 'TableStyleMedium2'}),
      // An explicit style object with no name is OOXML's "unstyled" — distinct from omitting the
      // style entirely, which a freshly-authored table fills with Excel's default (TableStyleMedium2).
      nullTheme: styleInfoOf({}),
      none: styleInfoOf({name: 'None', showRowStripes: true}),
    };
  },

  // Author a table from a header-name list (which may contain collisions), write it, and report the
  // column names emitted into the table part plus whether they are unique. OOXML requires unique
  // tableColumn names; colliding inputs must be disambiguated deterministically, not written verbatim
  // into a corrupt file → { ok, writtenNames, uniqueNames }.
  tableDuplicateColumnNamesReport(headers: Untyped) {
    const wb = new Workbook();
    let ok = true;
    let writtenNames = null;
    try {
      wb.addWorksheet('S').addTable({
        name: 'T',
        ref: 'A1',
        columns: headers.map((name: Untyped) => ({name})),
        rowCount: 1,
      });
      const parts = partMapOf(writeXlsx(wb));
      const part = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
      writtenNames = [...parts[part!]!.matchAll(/<tableColumn\b[^>]*\bname="([^"]*)"/g)].map(
        (m) => m[1],
      );
    } catch (e) {
      ok = false;
      writtenNames = String((e as Untyped)?.message || e);
    }
    const uniqueNames =
      Array.isArray(writtenNames) &&
      new Set(writtenNames.map((n) => n!.toLowerCase())).size === writtenNames.length;
    return {ok, writtenNames, uniqueNames};
  },

  // Author a two-column table whose first column carries a numFmt style, append two data rows, then
  // round-trip and report the numFmt read back on each column's body cells → { writeOk, reloadOk,
  // writeError, styledBody, unstyledBody }. The per-column style must bake into the styled column's
  // body cells and leave the unstyled column untouched.
  tableColumnStyleReport(numFmt: Untyped) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    const table = s.addTable({
      name: 'T',
      ref: 'A1',
      columns: [{name: 'Amount', style: {numFmt}}, {name: 'Label'}],
      rowCount: 0,
    });
    s.getCell('A1').value = 'Amount';
    s.getCell('B1').value = 'Label';
    table.addRow([1234.5, 'x']);
    table.addRow([6789, 'y']);

    let writeOk = true;
    let writeError = null;
    let buffer = null;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = String((e as Untyped)?.message || e);
    }
    if (!writeOk)
      return {writeOk, writeError, reloadOk: false, styledBody: null, unstyledBody: null};

    let reloadOk = true;
    let styledBody = null;
    let unstyledBody = null;
    try {
      const back = readXlsx(buffer!).getWorksheet('S')!;
      styledBody = [back.getCell('A2').numFmt ?? null, back.getCell('A3').numFmt ?? null];
      unstyledBody = [back.getCell('B2').numFmt ?? null, back.getCell('B3').numFmt ?? null];
    } catch (e) {
      reloadOk = false;
      writeError = String((e as Untyped)?.message || e);
    }
    return {writeOk, writeError, reloadOk, styledBody, unstyledBody};
  },

  // Build a table-bearing spec, round-trip it through a write→read, fetch the named table on the
  // reloaded model, append the requested rows, then re-write and re-read to report the final row
  // count. A table read from a file must expose its data rows and accept appends exactly like a
  // freshly-created one → { hasTable, loadedRowCount, addError, committed, finalRowCount }.
  roundtripTableAppend(spec: Untyped, {tableName, appendRows}: Untyped) {
    const reloaded = readXlsx(writeXlsx(buildFrom(spec)));
    let table: Untyped = null;
    for (const s of reloaded.worksheets) {
      const found = s.getTable(tableName);
      if (found) {
        table = found;
        break;
      }
    }
    const hasTable = table !== null;
    if (!hasTable) {
      return {
        hasTable,
        loadedRowCount: null,
        addError: null,
        committed: false,
        finalRowCount: null,
      };
    }
    const loadedRowCount = table.rowCount;

    let addError = null;
    for (const row of appendRows) {
      try {
        table.addRow(row);
      } catch (e) {
        addError = String((e as Untyped)?.message || e);
        break;
      }
    }

    let committed = false;
    let finalRowCount = null;
    if (addError === null) {
      try {
        const out = writeXlsx(reloaded);
        committed = true;
        const back = readXlsx(out);
        for (const s of back.worksheets) {
          const found = s.getTable(tableName);
          if (found) {
            finalRowCount = found.rowCount;
            break;
          }
        }
      } catch (e) {
        addError = String((e as Untyped)?.message || e);
      }
    }
    return {hasTable, loadedRowCount, addError, committed, finalRowCount};
  },

  // Author a table over A1:B3, populate its cells, load the package, edit a body cell (B2 → 999),
  // and re-write — reporting that both writes and the reload succeed, the table part and its unique
  // worksheet relationship survive, and the edited value reads back. Editing a cell inside a table's
  // range must not truncate or corrupt the table part or its rels.
  tableCellEditRoundtrip() {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.addTable({name: 'T', ref: 'A1', columns: [{name: 'H1'}, {name: 'H2'}], rowCount: 2});
    s.getCell('A1').value = 'H1';
    s.getCell('B1').value = 'H2';
    s.getCell('A2').value = 'a';
    s.getCell('B2').value = 1;
    s.getCell('A3').value = 'b';
    s.getCell('B3').value = 2;

    let writeOk = true;
    let writeError = null;
    let firstBuffer = null;
    try {
      firstBuffer = writeXlsx(wb);
    } catch (e) {
      writeOk = false;
      writeError = String((e as Untyped)?.message || e);
    }
    if (!writeOk) {
      return {
        writeOk,
        writeError,
        reloadOk: false,
        hasTablePart: false,
        tablePresent: false,
        editedValue: null,
        relUnique: false,
      };
    }

    let reloadOk = true;
    let hasTablePart = false;
    let tablePresent = false;
    let editedValue = null;
    let relUnique = false;
    try {
      const reloaded = readXlsx(firstBuffer!);
      const sheet = reloaded.getWorksheet('S')!;
      sheet.getCell('B2').value = 999;
      const out = writeXlsx(reloaded);
      const parts = partMapOf(out);
      const tablePart = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
      hasTablePart = tablePart !== undefined;
      const relPart = Object.keys(parts).find((n) =>
        /xl\/worksheets\/_rels\/sheet\d+\.xml\.rels$/.test(n),
      );
      const relIds = relPart ? [...parts[relPart]!.matchAll(/Id="([^"]*)"/g)].map((m) => m[1]) : [];
      relUnique = relIds.length > 0 && new Set(relIds).size === relIds.length;
      const back = readXlsx(out);
      const backSheet = back.getWorksheet('S')!;
      tablePresent = backSheet.tables.some((t) => t.name === 'T');
      editedValue = backSheet.getCell('B2').value;
    } catch (e) {
      reloadOk = false;
      writeError = String((e as Untyped)?.message || e);
    }
    return {writeOk, writeError, reloadOk, hasTablePart, tablePresent, editedValue, relUnique};
  },

  // Write a table whose first column name embeds CR/LF line breaks, then report the first
  // <tableColumn> tag and whether it carries a raw (unescaped) CR or LF. A raw control char in an
  // attribute value is not preserved by XML normalisation (a CR reparses as a space) and makes the
  // package suspect, so the name must be emitted XML-escaped (&#13;&#10;), not raw.
  tableColumnNameControlChars() {
    const wb = new Workbook();
    wb.addWorksheet('S').addTable({
      name: 'T',
      ref: 'A1',
      columns: [{name: 'Test\r\nmultiple\r\nlines'}, {name: 'Plain'}],
      rowCount: 1,
    });
    let writeOk = true;
    let writeError = null;
    let firstColumnTag = null;
    let rawControlChars = null;
    try {
      const parts = partMapOf(writeXlsx(wb));
      const part = Object.keys(parts).find((n) => /^xl\/tables\/table\d+\.xml$/.test(n));
      firstColumnTag = (parts[part!]!.match(/<tableColumn\b[^>]*\/?>/) || [])[0] ?? null;
      rawControlChars = firstColumnTag === null ? null : /[\r\n]/.test(firstColumnTag);
    } catch (e) {
      writeOk = false;
      writeError = String((e as Untyped)?.message || e);
    }
    return {writeOk, writeError, firstColumnTag, rawControlChars};
  },

  // Read every table part of a fixture, then load→save it and read the re-emitted parts, reporting
  // each table's autoFilter / header-row / totals-row / column-count facts before and after. A no-op
  // round-trip of a table that has no autoFilter must not inject one, flip the header row off, or turn
  // totalsRowShown on; a table that does have one must keep its ref and column count.
  roundtripFixtureTableXml(rel: Untyped) {
    const facts = (xml: Untyped) => ({
      hasAutoFilter: /<(?:\w+:)?autoFilter\b/.test(xml),
      autoFilterRef: (xml.match(/<(?:\w+:)?autoFilter\b[^>]*\bref="([^"]*)"/) || [])[1] ?? null,
      headerRowCount: (xml.match(/\bheaderRowCount="([^"]*)"/) || [])[1] ?? null,
      totalsRowShown: (xml.match(/\btotalsRowShown="([^"]*)"/) || [])[1] ?? null,
      columnCount: (xml.match(/<tableColumns\b[^>]*\bcount="([^"]*)"/) || [])[1] ?? null,
    });
    const tablePartsInOrder = (parts: Untyped) =>
      Object.keys(parts)
        .filter((n) => /^xl\/tables\/table\d+\.xml$/.test(n))
        .sort((a, b) => Number(a.match(/\d+/)![0]) - Number(b.match(/\d+/)![0]))
        .map((n) => parts[n]);
    const buffer = fs.readFileSync(path.join(FIXTURES_ROOT, rel));
    const source = tablePartsInOrder(partMapOf(buffer));
    const rewritten = tablePartsInOrder(partMapOf(writeXlsx(readXlsx(buffer))));
    return {
      tables: source.map((xml, i) => ({
        name: (xml.match(/<table\b[^>]*\bname="([^"]*)"/) || [])[1] ?? null,
        source: facts(xml),
        rewritten: rewritten[i] ? facts(rewritten[i]) : null,
      })),
    };
  },

  // Author a table whose display name differs from its internal name, then report the displayName
  // written into the table part and the internal/display names read back from the reloaded model —
  // a serializer that mis-keys the property drops the display name to the internal default.
  tableDisplayNameReport(display: Untyped) {
    const wb = new Workbook();
    wb.addWorksheet('S').addTable({
      name: 'MyTable',
      displayName: display,
      ref: 'A1',
      columns: [{name: 'C'}],
      rowCount: 1,
    });
    const buffer = writeXlsx(wb);
    const part = partMapOf(buffer)['xl/tables/table1.xml'] || '';
    const writtenDisplayName = (part.match(/\bdisplayName="([^"]*)"/) || [])[1] ?? null;
    const table = readXlsx(buffer).getWorksheet('S')!.tables[0];
    return {
      writtenDisplayName,
      reloadedDisplayName: table ? table.displayName : null,
      reloadedName: table ? table.name : null,
    };
  },
};
