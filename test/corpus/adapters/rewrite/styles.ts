// Fonts, fills, borders, alignment, number formats and the style-deduplication that decides
// which of them survive a round-trip.

import type {CorpusApi} from '../../case.ts';
import {partMapOf} from './package-facts.ts';
import {fixtureBytes, readFixture, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {applyStyle, buildFrom} from './spec-model.ts';
import {reloadPatched} from './xml-probes.ts';

export const styles = {
  // Assign a valid format-code string (alongside font/alignment/protection) to one cell and a
  // structured OBJECT (a parsed `{id, formatCode}` shape) to another, write, and report the string's
  // survival plus whether the styles part was corrupted → { controlNumFmtReload, stylesHasObjectObject }.
  // The object must never serialize as `formatCode="[object Object]"`.
  numFmtObjectCorruptionReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const control = sheet.getCell('A1');
    control.value = 45000;
    control.numFmt = 'yyyy-mmm-dd';
    control.font = {bold: true};
    control.alignment = {horizontal: 'center'};
    control.protection = {locked: false};
    const bad = sheet.getCell('A2');
    bad.value = 42;
    // A caller wrongly assigns the structured numFmt object Excel parses a cell's format into.
    bad.numFmt = {id: 164, formatCode: '0.00'} as CorpusApi;
    const stylesXml = partMapOf(writeXlsx(wb))['xl/styles.xml'] || '';
    const back = readXlsx(writeXlsx(wb)).getWorksheet('S')!;
    return {
      controlNumFmtReload: back.getCell('A1').numFmt ?? null,
      stylesHasObjectObject: stylesXml.includes('[object Object]'),
    };
  },

  // Mark a cell whose content looks like a formula with the quote-prefix flag, write, and report
  // whether its cell-format record carries quotePrefix and whether the flag survives a round-trip →
  // { writtenQuotePrefix, reloaded }. The flag forces formula-like content to be stored as literal text.
  quotePrefixReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const cell = sheet.getCell('A1');
    cell.value = '=1+1';
    cell.quotePrefix = true;
    const buffer = writeXlsx(wb);
    const writtenQuotePrefix = /<xf\b[^>]*quotePrefix="1"/.test(
      partMapOf(buffer)['xl/styles.xml'] || '',
    );
    const reloaded = readXlsx(buffer).getWorksheet('S')!.getCell('A1').quotePrefix === true;
    return {writtenQuotePrefix, reloaded};
  },

  // Load a fixture whose A1 fill lives only in a named cell style (cellXfs xfId → cellStyleXfs), and
  // report: the source's cellStyleXfs count, A1's resolved fill, and — after a load→save round-trip —
  // the re-emitted cellStyleXfs count and whether A1's cellXfs entry still links via xfId → so a case
  // asserts the named-style layer is honoured on read and preserved (with its link) on write.
  namedStyleFillReport(rel: CorpusApi) {
    const countCellStyleXfs = (xml: CorpusApi) =>
      Number((xml.match(/<cellStyleXfs count="(\d+)"/) || [])[1] || 0);
    const srcCellStyleXfsCount = countCellStyleXfs(
      partMapOf(fixtureBytes(rel))['xl/styles.xml'] || '',
    );

    const workbook = readXlsx(fixtureBytes(rel));
    const readFill = workbook.worksheets[0]?.getCell('A1').fill ?? null;

    const out = writeXlsx(workbook);
    const outParts = partMapOf(out);
    const outStyles = outParts['xl/styles.xml'] || '';
    const roundtripCellStyleXfsCount = countCellStyleXfs(outStyles);
    // Resolve A1's cellXfs entry via its `s` index and check it retains a non-zero xfId link.
    const sIndex = Number(
      (outParts['xl/worksheets/sheet1.xml']?.match(/<c r="A1"[^>]*\bs="(\d+)"/) || [])[1] || 0,
    );
    const cellXfsBlock =
      (outStyles.match(/<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/) || [])[1] || '';
    const cellXfEntries = cellXfsBlock.match(/<xf\b[^>]*(?:\/>|>[\s\S]*?<\/xf>)/g) || [];
    const a1Xf = cellXfEntries[sIndex] || '';
    const xfIdMatch = a1Xf.match(/\bxfId="(\d+)"/);
    const roundtripCellHasXfIdLink = xfIdMatch !== null && Number(xfIdMatch[1]) > 0;

    return {srcCellStyleXfsCount, readFill, roundtripCellStyleXfsCount, roundtripCellHasXfIdLink};
  },

  // Write a value under a date number format and report the sheet XML's health → { ok, hasNaN,
  // hasInvalidDate, cellXml }. A string, a null (empty cell), or an Invalid Date under a date format
  // must never leak a bare "NaN" or "Invalid Date" into the cell value.
  dateNumFmtValueReport(kind: CorpusApi) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const cell = sheet.getCell('A1');
    if (kind === 'string') cell.value = 'not a date';
    else if (kind === 'null') cell.value = null;
    else if (kind === 'invalidDate') cell.value = new Date(NaN);
    cell.numFmt = 'yyyy-mm-dd';
    let ok = true;
    let sheetXml = '';
    try {
      sheetXml = partMapOf(writeXlsx(workbook))['xl/worksheets/sheet1.xml'] || '';
    } catch {
      ok = false;
    }
    const cellXml = (sheetXml.match(/<c r="A1"[\s\S]*?(?:\/>|<\/c>)/) || [])[0] ?? '';
    return {
      ok,
      hasNaN: /NaN/.test(sheetXml),
      hasInvalidDate: /Invalid Date/.test(sheetXml),
      cellXml,
    };
  },

  // Apply a font to each named cell, then read each requested cell's font back → { <ref>:
  // font|null }. Each cell owns its own font, so a font set on one cell is observable there
  // and nowhere else — the isolation this reports. In-memory, matching the contract: the
  // <fonts>-table write/read path is exercised by the io/xlsx unit tests.
  probeCellFonts({apply = [], read = []}: CorpusApi) {
    const sheet = new Workbook().addWorksheet('sheet');
    for (const {cell, font} of apply) sheet.getCell(cell).font = font;
    const fonts: Record<string, CorpusApi> = {};
    for (const address of read) fonts[address] = sheet.getCell(address).font ?? null;
    return JSON.parse(JSON.stringify(fonts));
  },

  // Write the spec and report the shared style table's size plus the index each requested cell
  // resolved to → { cellXfCount, indices: { <ref>: index|null } }. styles.xml is a SHARED table
  // referenced by index, so identically-styled cells must collapse to one <cellXfs> entry (one
  // shared index) — dedup neither inflating to one entry per cell nor over-collapsing distinct
  // styles. A cell left at the default style carries no `s` and reports null.
  styleDedupReport(spec: CorpusApi, cells = []) {
    const parts = partMapOf(writeXlsx(buildFrom(spec)));
    const styles = parts['xl/styles.xml'] || '';
    const xfBlock = (styles.match(/<cellXfs\b[\s\S]*?<\/cellXfs>/) || [''])[0];
    const cellXfCount = (xfBlock!.match(/<xf\b/g) || []).length;
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const indices: Record<string, CorpusApi> = {};
    for (const ref of cells) {
      const m = sheetXml.match(new RegExp(`<c\\b[^>]*\\br="${ref}"[^>]*\\bs="(\\d+)"`));
      indices[ref] = m ? Number(m[1]) : null;
    }
    return {cellXfCount, indices};
  },

  // Write one plain-valued, unformatted cell, round-trip it, and report the font it reads back →
  // { hasFont, fontName, fontSize }. An unstyled cell renders in the workbook default font, so it
  // must resolve a concrete name/size rather than reporting no font at all.
  unstyledCellFontReport() {
    const wb = new Workbook();
    wb.addWorksheet('S').getCell('A1').value = 'hello';
    const cell = readXlsx(writeXlsx(wb)).getWorksheet('S')!.getCell('A1');
    const font = cell.font || null;
    return {
      hasFont: !!font,
      fontName: font ? (font.name ?? null) : null,
      fontSize: font ? (font.size ?? null) : null,
    };
  },

  // Colour a sheet's tab with an 8-digit ARGB (alpha first) alongside an uncoloured sheet, then
  // report the round-trip → { tabColorArgbWritten, reReadArgb, uncoloredHasTab }. A set tab colour
  // must survive verbatim while a sheet with none must not gain a spurious one.
  tabColorRoundtrip() {
    const wb = new Workbook();
    const colored = wb.addWorksheet('Colored');
    colored.tabColor = {argb: 'FFFF0000'};
    colored.getCell('A1').value = 'x';
    const plain = wb.addWorksheet('Plain');
    plain.getCell('A1').value = 'y';
    const buffer = writeXlsx(wb);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const written = (sheetXml.match(/<tabColor\b[^>]*rgb="([^"]*)"/) || [null, null])[1];
    const reload = readXlsx(buffer);
    const coloredTab = reload.getWorksheet('Colored')!.tabColor;
    return {
      tabColorArgbWritten: written,
      reReadArgb: coloredTab?.argb || null,
      uncoloredHasTab: !!reload.getWorksheet('Plain')!.tabColor,
    };
  },

  // Write a solid fill twice — once with a clean bare ARGB, once with a CSS-habit '#'-prefixed one —
  // and report the emitted <fgColor rgb="..."> for each → { validRgb, hashRgb }. Both must serialize
  // as valid 8-hex-digit values; a '#'-prefixed input must be normalized, never passed through as a
  // malformed 9-character colour.
  fillArgbHashPrefixReport() {
    const emittedFgColor = (argb: CorpusApi) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb}};
      const stylesXml = partMapOf(writeXlsx(wb))['xl/styles.xml'] || '';
      return (stylesXml.match(/<fgColor rgb="([^"]*)"/) || [null, null])[1];
    };
    return {validRgb: emittedFgColor('FFBFBFBF'), hashRgb: emittedFgColor('#FFBFBFBF')};
  },

  // Author a solid fill with a 6-hex RGB (no alpha) and with a malformed value, and report how the
  // writer treats each → { sixHexRgb, rejectsMalformed }. A 6-hex RGB is the common "colour without
  // its alpha channel" case: it must be promoted to a valid opaque 8-hex ARGB, not emitted as a
  // 6-char rgb that Excel renders black. A value that is neither 6 nor 8 hex digits is a programming
  // error and must be rejected, never written as a colour Excel silently renders black.
  argbNormalizationReport() {
    const emittedFgColor = (argb: CorpusApi) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb}};
      const stylesXml = partMapOf(writeXlsx(wb))['xl/styles.xml'] || '';
      return (stylesXml.match(/<fgColor rgb="([^"]*)"/) || [null, null])[1];
    };
    let rejectsMalformed = false;
    try {
      emittedFgColor('12345');
    } catch {
      rejectsMalformed = true;
    }
    return {sixHexRgb: emittedFgColor('00FF00'), rejectsMalformed};
  },

  outlinePropertiesRoundtrip() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.outline.summaryBelow = false;
    ws.outline.summaryRight = false;
    ws.getCell('A1').value = 'x';
    const buffer = writeXlsx(wb);
    const sheetXml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const outlinePr = (sheetXml.match(/<outlinePr\b[^>]*\/?>/) || [''])[0];
    const reload = readXlsx(buffer);
    const outline = reload.getWorksheet('S')!.outline;
    return {
      outlinePrEmitted: /summaryBelow="0"/.test(outlinePr) && /summaryRight="0"/.test(outlinePr),
      reReadSummaryBelow: outline.summaryBelow ?? null,
      reReadSummaryRight: outline.summaryRight ?? null,
    };
  },

  // Author a bold cell, then rewrite the emitted <b/> flag to each explicit form and report how
  // the reader reads bold back → { bareTag, valOne, valZero }. A boolean font flag's `val` governs:
  // a bare tag or val="1" is on, val="0" is off — presence alone must not force true.
  fontExplicitFalseBoldReport() {
    const readBoldWith = (tag: CorpusApi) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').font = {bold: true};
      const reloaded = reloadPatched(writeXlsx(wb), {
        'xl/styles.xml': (xml) => xml.replace(/<b ?\/>/, tag),
      });
      const font = reloaded.getWorksheet('S')!.getCell('A1').font;
      return !!font?.bold;
    };
    return {
      bareTag: readBoldWith('<b/>'),
      valOne: readBoldWith('<b val="1"/>'),
      valZero: readBoldWith('<b val="0"/>'),
    };
  },

  // Author cells with italic/strike/underline on, rewrite each flag to its explicit-off form, and
  // report what the reader reads back → { italic, strike, underline }. val="0" turns a boolean flag
  // off; <u val="none"/> is the ABSENCE of an underline, so it must read back falsy — never the
  // truthy string "none".
  fontExplicitOffFlagsReport() {
    const readWith = (baseFont: CorpusApi, tagRe: CorpusApi, tag: CorpusApi, field: CorpusApi) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.getCell('A1').font = baseFont;
      const reloaded = reloadPatched(writeXlsx(wb), {
        'xl/styles.xml': (xml) => xml.replace(tagRe, tag),
      });
      const font = reloaded.getWorksheet('S')!.getCell('A1').font || {};
      return (font as Record<string, CorpusApi>)[field] ?? null;
    };
    return {
      italic: readWith({italic: true}, /<i ?\/>/, '<i val="0"/>', 'italic'),
      strike: readWith({strike: true}, /<strike ?\/>/, '<strike val="0"/>', 'strike'),
      underline: readWith({underline: true}, /<u ?\/>/, '<u val="none"/>', 'underline'),
    };
  },

  // Inject an xf whose alignment element carries only an explicit-false boolean (wrapText="0" /
  // shrinkToFit="0"), point A1 at it, and report the alignment the reader surfaces → { wrapTextZero,
  // shrinkZero }. An all-false alignment carries no information and must read back as no alignment
  // at all — the raw "0" is a truthy JS string, so a reader guarding on the raw value rather than the
  // parsed boolean would wrongly report a present alignment.
  alignmentFalseBooleanReport() {
    const readWithAlignment = (alignAttr: CorpusApi) => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      let injectedIndex = -1;
      const reloaded = reloadPatched(writeXlsx(wb), {
        'xl/styles.xml': (xml) => {
          const count = Number(xml.match(/<cellXfs count="(\d+)">/)![1]);
          injectedIndex = count;
          return xml
            .replace(/<cellXfs count="\d+">/, `<cellXfs count="${count + 1}">`)
            .replace(
              /<\/cellXfs>/,
              `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1">` +
                `<alignment ${alignAttr}/></xf></cellXfs>`,
            );
        },
        'xl/worksheets/sheet1.xml': (xml) =>
          xml.replace(/<c r="A1"([^>]*)>/, `<c r="A1" s="${injectedIndex}"$1>`),
      });
      const alignment = reloaded.getWorksheet('S')!.getCell('A1').alignment;
      return alignment ? JSON.parse(JSON.stringify(alignment)) : null;
    };
    return {
      wrapTextZero: readWithAlignment('wrapText="0"'),
      shrinkZero: readWithAlignment('shrinkToFit="0"'),
    };
  },

  // Assign a vertical-alignment enum token to a cell — routed through the untyped adapter surface so a
  // value TypeScript would reject at compile time (e.g. the ExcelJS-era "middle") reaches the writer as
  // an untyped caller would smuggle it — then attempt a full write→read cycle → { writeThrew,
  // writeError, readBackVertical }. Captures the writer-boundary contract: a value the writer accepts
  // must survive read-back, so an out-of-contract token must be rejected at write or preserved, never
  // silently serialized into a schema-invalid file the library's own reader then discards.
  alignmentVerticalEnumReport(vertical: CorpusApi) {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    const cell = sheet.getCell('A1');
    cell.value = 'x';
    cell.alignment = {vertical};
    let writeThrew = false;
    let writeError: string | null = null;
    let readBackVertical: string | null = null;
    try {
      const back = readXlsx(writeXlsx(wb)).getWorksheet('S')!.getCell('A1').alignment;
      readBackVertical = back?.vertical ?? null;
    } catch (e) {
      writeThrew = true;
      writeError = String((e as CorpusApi)?.message || e);
    }
    return {writeThrew, writeError, readBackVertical};
  },

  // Give one cell a fill and another a border but NO value, leave a third cell entirely untouched,
  // round-trip, and report each → { filledArgb, filledValue, borderedStyle, borderedValue, untouched }.
  // A formatted-but-empty cell is a real cell Excel keeps: its style must survive the write and the
  // cell must stay value-less, while a cell with neither value nor style must not be fabricated.
  styledEmptyCellReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getCell('A1').value = 'anchor';
    sheet.getCell('B2').fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: 'FF00FF00'}};
    sheet.getCell('C3').border = {top: {style: 'thin', color: {argb: 'FF000000'}}};
    sheet.getCell('D4'); // materialised but never given a value or style
    const back = readXlsx(writeXlsx(wb)).getWorksheet('S')!;
    const filled = back.getCell('B2');
    const bordered = back.getCell('C3');
    // fgColor lives on the pattern-fill variant, past the general Fill union surface.
    const filledFill = filled.fill as CorpusApi;
    return {
      filledArgb: filledFill?.fgColor ? filledFill.fgColor.argb : null,
      filledValue: filled.value,
      borderedStyle: bordered.border?.top ? bordered.border.top.style : null,
      borderedValue: bordered.value,
      untouched: !!(back.getCell('D4').fill || back.getCell('D4').value),
    };
  },

  readFixtureCellStyles(rel: CorpusApi, cells: CorpusApi = []) {
    const wb = readFixture(rel);
    const out: Record<string, CorpusApi> = {};
    for (const key of cells) {
      const [sheetName, addr] = key.split('!');
      const sheet = wb.getWorksheet(sheetName);
      const cell = sheet ? sheet.getCell(addr) : null;
      out[key] = cell
        ? {
            fill: cell.fill?.type ? JSON.parse(JSON.stringify(cell.fill)) : null,
            fontColor: cell.font?.color ? JSON.parse(JSON.stringify(cell.font.color)) : null,
          }
        : null;
    }
    return out;
  },

  // Read a real fixture, write it straight back out, read the result, and report whether any
  // cell's visible fill colour or border-edge colours changed across that no-op round-trip →
  // { checked, fillMismatches, borderMismatches, fillSample, borderSample }. Theme+tint and
  // indexed-palette references must survive verbatim so the sheet renders identically after a
  // pure open-then-save. Colour comparison is key-order-insensitive.
  roundtripFixtureColorFidelity(rel: CorpusApi) {
    const before = readFixture(rel);
    const after = readXlsx(writeXlsx(before));

    const realFill = (cell: CorpusApi) =>
      cell.fill && cell.fill.type === 'pattern' && cell.fill.pattern !== 'none' ? cell.fill : null;
    const borderColors = (cell: CorpusApi) => {
      if (!cell.border) return null;
      const edges: Record<string, CorpusApi> = {};
      for (const edge of ['top', 'left', 'right', 'bottom']) {
        if (cell.border[edge]?.color) edges[edge] = cell.border[edge].color;
      }
      return Object.keys(edges).length ? edges : null;
    };
    const stableSort = (v: CorpusApi): CorpusApi => {
      if (Array.isArray(v)) return v.map(stableSort);
      if (v && typeof v === 'object') {
        const sorted: Record<string, CorpusApi> = {};
        for (const k of Object.keys(v).sort()) sorted[k] = stableSort(v[k]);
        return sorted;
      }
      return v;
    };
    const norm = (v: CorpusApi) => JSON.stringify(stableSort(v ?? null));

    let checked = 0;
    let fillMismatches = 0;
    let borderMismatches = 0;
    let fillSample = null;
    let borderSample = null;
    for (const sheet of before.worksheets) {
      const other = after.getWorksheet(sheet.name);
      for (const {cells} of sheet.rows()) {
        for (const cell of cells) {
          if (!realFill(cell) && !borderColors(cell)) continue;
          checked += 1;
          const oc = other ? other.getCell(cell.address) : null;
          const bf = norm(realFill(cell));
          const af = oc ? norm(realFill(oc)) : '(missing)';
          if (bf !== af) {
            fillMismatches += 1;
            if (!fillSample)
              fillSample = {cell: `${sheet.name}!${cell.address}`, before: bf, after: af};
          }
          const bb = norm(borderColors(cell));
          const ab = oc ? norm(borderColors(oc)) : '(missing)';
          if (bb !== ab) {
            borderMismatches += 1;
            if (!borderSample)
              borderSample = {cell: `${sheet.name}!${cell.address}`, before: bb, after: ab};
          }
        }
      }
    }
    return {checked, fillMismatches, borderMismatches, fillSample, borderSample};
  },

  // Give one column a right border and later columns only a width, then round-trip and report
  // each cell's right border → { a1, b1, c1 }. A column's border is a default for its own cells,
  // so the declaring column's cell carries it while columns without a style of their own get
  // nothing — column styles are independent, not bled into subsequent columns.
  columnBorderScopedReport() {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('S');
    sheet.getColumn(1).border = {right: {style: 'thin', color: {argb: 'FF000000'}}};
    sheet.getColumn(2).width = 10;
    sheet.getCell('A1').value = 'a';
    sheet.getCell('B1').value = 'b';
    sheet.getCell('C1').value = 'c';
    const s = readXlsx(writeXlsx(wb)).getWorksheet('S')!;
    const rightBorder = (ref: CorpusApi) => {
      const b = s.getCell(ref).border;
      return !!b?.right?.style;
    };
    return {a1: rightBorder('A1'), b1: rightBorder('B1'), c1: rightBorder('C1')};
  },

  // Read a fixture, extract its column widths and pageSetup, write it straight back, re-read, and
  // report the same facts → { source, rewritten }. A faithful no-op round-trip must reproduce every
  // fractional column width and the print-scaling attributes the real file carries.
  roundtripFixtureStyleFacts(rel: CorpusApi) {
    // Model-level facts (column widths, pageSetup, dxfs) come from the parsed workbook; the custom
    // indexed-color palette is a raw styles.xml fact, so it is read straight from the part bytes on
    // each side — matching the legacy oracle, which extracts the same block from the zip.
    const facts = (workbook: CorpusApi, stylesXml: CorpusApi) => {
      const sheet = workbook.worksheets[0];
      const ps = sheet ? sheet.pageSetup : {};
      // Differential styles are preserved as verbatim `<dxf>` fragments; a rule's number format is
      // whatever formatCode the fragment carries, so a coerced "[object Object]" can never appear.
      const dxfs = workbook.differentialStyles;
      const dxfFormatCodes = dxfs.flatMap((f: CorpusApi) =>
        [...f.matchAll(/formatCode="([^"]*)"/g)].map((m) => m[1]),
      );
      const palette = stylesXml.match(/<indexedColors>([\s\S]*?)<\/indexedColors>/);
      return {
        columnWidths: sheet
          ? [...sheet.columns()].map((c) => c.properties.width).filter((w) => w !== undefined)
          : [],
        pageSetup: {
          scale: ps.scale ?? null,
          fitToWidth: ps.fitToWidth ?? null,
          fitToHeight: ps.fitToHeight ?? null,
          pageOrder: ps.pageOrder ?? null,
          orientation: ps.orientation ?? null,
          paperSize: ps.paperSize ?? null,
        },
        dxfCount: dxfs.length,
        dxfFormatCodes,
        hasIndexedColors: !!palette,
        indexedColorSample: palette
          ? [...palette[1].matchAll(/rgb="([0-9a-fA-F]+)"/g)].slice(0, 6).map((m) => m[1])
          : [],
      };
    };
    const before = readFixture(rel);
    const source = facts(before, partMapOf(fixtureBytes(rel))['xl/styles.xml'] || '');
    const buffer = writeXlsx(before);
    const after = readXlsx(buffer);
    return {source, rewritten: facts(after, partMapOf(buffer)['xl/styles.xml'] || '')};
  },

  // Assign one base style object to two cells, then spread-reassign one cell's font color →
  // { a1Color, a2Color, bled }. The sibling given the same base must keep its original font.
  sharedBaseStyleFontMutation() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    const base = {font: {name: 'Arial', size: 11}};
    sheet.getCell('A1').value = 'YES';
    sheet.getCell('A2').value = 'NO';
    applyStyle(sheet.getCell('A1'), base);
    applyStyle(sheet.getCell('A2'), base);
    sheet.getCell('A1').font = {...sheet.getCell('A1').font, color: {argb: 'FF00FF00'}};
    const s = readXlsx(writeXlsx(workbook)).getWorksheet('S')!;
    const colorOf = (ref: CorpusApi) => {
      const f = s.getCell(ref).font;
      return f?.color ? (f.color.argb ?? null) : null;
    };
    const a1Color = colorOf('A1');
    const a2Color = colorOf('A2');
    return {a1Color, a2Color, bled: a2Color === 'FF00FF00'};
  },

  // Author three cells sharing one style record, load, border ONE, round-trip →
  // { a1, a2, a3, bled }. Only the targeted cell may gain a border.
  loadMutateCellBorder() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const r of [1, 2, 3]) {
      const c = sheet.getCell(`A${r}`);
      c.value = 'x';
      c.font = {bold: true};
    }
    const loaded = readXlsx(writeXlsx(workbook));
    loaded.getWorksheet('S')!.getCell('A1').border = {
      top: {style: 'thin'},
      left: {style: 'thin'},
      bottom: {style: 'thin'},
      right: {style: 'thin'},
    };
    const s = readXlsx(writeXlsx(loaded)).getWorksheet('S')!;
    const hasBorder = (ref: CorpusApi) => {
      const b = s.getCell(ref).border;
      return !!b?.top?.style;
    };
    return {
      a1: hasBorder('A1'),
      a2: hasBorder('A2'),
      a3: hasBorder('A3'),
      bled: hasBorder('A2') || hasBorder('A3'),
    };
  },

  // Author two cells with one shared fill, load, replace ONE cell's fill, read the sibling in
  // memory and after write-back → { sibling, mutatedTo, original, bled, diskSibling, diskBled }.
  loadMutateCellStyle({sharedFill = 'FFFF0000', mutateTo = 'FF00FF00'}: CorpusApi = {}) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getCell('A1').value = 'a';
    s.getCell('B1').value = 'b';
    const fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: sharedFill}} as CorpusApi;
    s.getCell('A1').fill = fill;
    s.getCell('B1').fill = fill; // identical formatting → one shared style index on disk
    // fgColor lives on the pattern-fill variant, past the general Fill union surface.
    const fgOf = (cell: CorpusApi) =>
      (cell.fill as CorpusApi)?.fgColor ? ((cell.fill as CorpusApi).fgColor.argb ?? null) : null;

    const wb2 = readXlsx(writeXlsx(wb));
    const s2 = wb2.getWorksheet('S')!;
    s2.getCell('A1').fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: mutateTo},
    } as CorpusApi;
    const sibling = fgOf(s2.getCell('B1'));

    const diskSibling = fgOf(readXlsx(writeXlsx(wb2)).getWorksheet('S')!.getCell('B1'));
    return {
      sibling,
      mutatedTo: mutateTo,
      original: sharedFill,
      bled: sibling === mutateTo,
      diskSibling,
      diskBled: diskSibling === mutateTo,
    };
  },

  // Author two cells with one shared font, load, spread-reassign ONE cell's font color, read the
  // sibling → { edited, sibling, original, mutatedTo, bled }.
  loadMutateCellFont({original = 'FF000000', mutateTo = 'FFFF0000'}: CorpusApi = {}) {
    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    const font = {name: 'Arial', size: 12, color: {argb: original}};
    s.getCell('A1').value = 'a';
    s.getCell('A1').font = font;
    s.getCell('B1').value = 'b';
    s.getCell('B1').font = font; // identical formatting → one shared style index on disk

    const s2 = readXlsx(writeXlsx(wb)).getWorksheet('S')!;
    s2.getCell('A1').font = {...s2.getCell('A1').font, color: {argb: mutateTo}};
    const colorOf = (cell: CorpusApi) => (cell.font?.color ? (cell.font.color.argb ?? null) : null);
    const sibling = colorOf(s2.getCell('B1'));
    return {
      edited: colorOf(s2.getCell('A1')),
      sibling,
      original,
      mutatedTo: mutateTo,
      bled: sibling === mutateTo,
    };
  },

  // Load two cells sharing one style record, set ONE style facet (alignment | numFmt | protection)
  // on one via its setter, and report whether it bled into the sibling in memory and on disk →
  // { facet, target, sibling, original, bled, diskSibling, diskBled }. The remaining facets of the
  // copy-on-write family, alongside fill/font/border above.
  loadMutateCellFacet(facet: CorpusApi = 'alignment') {
    const readFacet = (
      {
        alignment: (c: CorpusApi) => c.alignment?.horizontal || null,
        numFmt: (c: CorpusApi) => c.numFmt || null,
        protection: (c: CorpusApi) =>
          c.protection && typeof c.protection.locked === 'boolean' ? c.protection.locked : null,
      } as Record<string, CorpusApi>
    )[facet];
    const apply = (
      {
        alignment: (c: CorpusApi) => {
          c.alignment = {horizontal: 'center'};
        },
        numFmt: (c: CorpusApi) => {
          c.numFmt = '#,##0';
        },
        protection: (c: CorpusApi) => {
          c.protection = {locked: false};
        },
      } as Record<string, CorpusApi>
    )[facet];
    if (!readFacet) throw new Error(`unknown style facet: ${facet}`);

    const wb = new Workbook();
    const s = wb.addWorksheet('S');
    s.getCell('A1').value = 'a';
    s.getCell('B1').value = 'b';
    const base = {numFmt: '0.00'}; // one identical non-default style → both cells dedup to one xf index
    applyStyle(s.getCell('A1'), base);
    applyStyle(s.getCell('B1'), base);

    const wb2 = readXlsx(writeXlsx(wb));
    const s2 = wb2.getWorksheet('S')!;
    const original = readFacet(s2.getCell('B1'));
    apply(s2.getCell('A1'));
    const target = readFacet(s2.getCell('A1'));
    const sibling = readFacet(s2.getCell('B1'));

    const diskSibling = readFacet(readXlsx(writeXlsx(wb2)).getWorksheet('S')!.getCell('B1'));
    return {
      facet,
      target,
      sibling,
      original,
      bled: sibling !== original,
      diskSibling,
      diskBled: diskSibling !== original,
    };
  },

  // --- CSV (src/io/csv) -------------------------------------------------------------------------
  // The contract mirrors the oracle's ExcelJS-shaped options; here they translate onto the rewrite's
  // cleaner CsvReadOptions/CsvWriteOptions. A read yields a JSON-serializable 2-D array of typed
  // values (Date → { date: iso }, error → { error }, else the scalar or null), matching the oracle so
  // the same cases assert unchanged.
};
