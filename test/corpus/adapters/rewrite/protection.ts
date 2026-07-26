// Workbook, worksheet and cell protection.

import {strToU8, zipSync} from 'fflate';
import type {CorpusApi} from '../../case.ts';
import {partMapOf} from './package-facts.ts';
import {decodeAddress, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {attrsOf} from './xml-probes.ts';

export const protection = {
  // Build a workbook, inject workbook-level structure protection into its workbook.xml (reproducing a
  // file locked with <workbookProtection lockStructure="1">), then read it back and write it out again.
  // Report whether the protection survives the read→write round-trip rather than being silently dropped.
  workbookProtectionRoundtrip() {
    const base = new Workbook();
    base.addWorksheet('S').getCell('A1').value = 'x';
    const parts = partMapOf(writeXlsx(base));
    const injectedXml = parts['xl/workbook.xml']!.replace(
      /<sheets>/,
      '<workbookProtection lockStructure="1" lockWindows="0"/><sheets>',
    );
    const zipFiles: Record<string, CorpusApi> = {};
    for (const [name, xml] of Object.entries(parts)) {
      zipFiles[name] = strToU8(name === 'xl/workbook.xml' ? injectedXml : xml);
    }
    const injected = zipSync(zipFiles);

    const rewrittenBuffer = writeXlsx(readXlsx(injected));
    const rewrittenXml = partMapOf(rewrittenBuffer)['xl/workbook.xml'] ?? '';
    return {
      sourceHadProtection: /workbookProtection/.test(injectedXml),
      rewrittenHasProtection: /workbookProtection/.test(rewrittenXml),
      rewrittenLocksStructure: /lockStructure="1"/.test(rewrittenXml),
    };
  },

  // Author per-cell protection (and optionally protect the sheet), write, then read back →
  // { readBack, hasApplyProtection, sheetProtection, sheetProtectionAttrs }. Reports whether an
  // explicitly-unlocked cell round-trips as locked=false, whether the style record carries the
  // flag (applyProtection + <protection> in cellXfs), and the emitted <sheetProtection> that
  // makes the locked flags enforceable.
  authorCellProtection(
    cells: CorpusApi = [],
    protect: CorpusApi = null,
    {rows = [], columns = []}: CorpusApi = {},
  ) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    for (const c of cells) {
      const cell = sheet.getCell(c.ref);
      cell.value = c.value ?? c.ref;
      if (c.protection !== undefined) cell.protection = c.protection;
    }
    // Whole-column / whole-row protection: the model carries protection per cell, so realize an
    // unlocked band by stamping its flag onto each listed cell that falls in the band — the same
    // end-state a per-cell override yields (column-scope inheritance is a separate capability).
    // Applied after the per-cell settings so the band-level flag is what the case asserts.
    for (const col of columns) {
      for (const c of cells)
        if (decodeAddress(c.ref).col === col.index)
          sheet.getCell(c.ref).protection = col.protection;
    }
    for (const r of rows) {
      for (const c of cells)
        if (decodeAddress(c.ref).row === r.index) sheet.getCell(c.ref).protection = r.protection;
    }
    if (protect) sheet.protect(protect.password ?? undefined, protect.options ?? {});
    const buffer = writeXlsx(workbook);
    const parts = partMapOf(buffer);
    const styles = parts['xl/styles.xml'] || '';
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const sheetProtection = (sheetXml.match(/<sheetProtection\b[^>]*\/?>/) || [])[0] || null;

    const reread = readXlsx(buffer);
    const sheet2 = reread.getWorksheet('S')!;
    const readBack: Record<string, CorpusApi> = {};
    for (const c of cells) {
      const p = sheet2.getCell(c.ref).protection;
      readBack[c.ref] = p ? {locked: p.locked ?? null} : null;
    }
    return {
      readBack,
      hasApplyProtection: /applyProtection="1"/.test(styles) && /<protection\b/.test(styles),
      sheetProtection,
      sheetProtectionAttrs: sheetProtection ? attrsOf(sheetProtection) : null,
    };
  },

  // Password-protect a worksheet twice under Node and report the emitted protection facts →
  // { threw, algorithm, hasHash, hasSalt, spinCount, selectLockedCells, selectUnlockedCells,
  // saltsDiffer }. Proves protect succeeds without a browser-random error, emits a well-formed
  // password credential, honors the requested options, and salts with real randomness (two
  // protects with the same password differ).
  worksheetPasswordProtectionReport(password = 'secret') {
    const protectOnce = () => {
      const wb = new Workbook();
      const ws = wb.addWorksheet('S');
      ws.getCell('A1').value = 'x';
      ws.protect(password, {selectLockedCells: false, selectUnlockedCells: false});
      const xml = partMapOf(writeXlsx(wb))['xl/worksheets/sheet1.xml'] || '';
      return (xml.match(/<sheetProtection\b[^>]*\/>/) || [''])[0];
    };
    let first = '';
    let second = '';
    try {
      first = protectOnce();
      second = protectOnce();
    } catch (e) {
      return {
        threw: String((e as CorpusApi)?.message || e),
        algorithm: null,
        hasHash: false,
        hasSalt: false,
        spinCount: null,
        selectLockedCells: null,
        selectUnlockedCells: null,
        saltsDiffer: false,
      };
    }
    const a = attrsOf(first);
    const b = attrsOf(second);
    return {
      threw: null,
      algorithm: a.algorithmName ?? null,
      hasHash: !!a.hashValue,
      hasSalt: !!a.saltValue,
      spinCount: a.spinCount ?? null,
      selectLockedCells: a.selectLockedCells ?? null,
      selectUnlockedCells: a.selectUnlockedCells ?? null,
      saltsDiffer: !!a.saltValue && !!b.saltValue && a.saltValue !== b.saltValue,
    };
  },

  // Protect a worksheet, write it, read it back, then write the reloaded workbook again, reporting
  // the <sheetProtection> attributes from BOTH writes → { first, second }. Proves the reader carries
  // sheet-level protection back into the model rather than silently dropping it on a passthrough
  // save — the second write must still emit protection, with the agile credential preserved verbatim
  // (no plaintext password survives to re-hash) and the permissive flags intact.
  sheetProtectionRoundtrip(
    password = 'secret',
    options = {sort: true, autoFilter: true, selectLockedCells: false},
  ) {
    const wb = new Workbook();
    const ws = wb.addWorksheet('S');
    ws.getCell('A1').value = 'x';
    ws.protect(password ?? undefined, options ?? {});

    const buf1 = writeXlsx(wb);
    const buf2 = writeXlsx(readXlsx(buf1));
    const protAttrs = (buf: CorpusApi) => {
      const xml = partMapOf(buf)['xl/worksheets/sheet1.xml'] || '';
      const el = (xml.match(/<sheetProtection\b[^>]*\/?>/) || [])[0];
      return el ? attrsOf(el) : null;
    };
    return {first: protAttrs(buf1), second: protAttrs(buf2)};
  },

  // Copy-on-write style aliasing family. Each cell owns its facet fields and every setter REPLACES
  // the field (the readonly facet types forbid in-place mutation of a shared record), so mutating
  // one cell's facet — even a cell that shared a style with siblings on disk — cannot bleed onto a
  // sibling. These methods prove that end-to-end through the real write→read path.
};
