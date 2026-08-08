import {messageOf} from '../../thrown.ts';
// Hyperlinks, including the internal (same-workbook) form and how it serializes.

import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {readFixture, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {attrsOf} from './xml-probes.ts';

export const hyperlinks = {
  // Read a fixture and report every hyperlink cell as { <addr>: { hyperlink, text } }, with a rich
  // display label flattened to its concatenated text — for asserting a foreign file's links (and the
  // rejoining of an external URL's fragment carried in the location attribute) are read faithfully.
  async readFixtureHyperlinks(rel: Untyped) {
    const flatten = (t: Untyped) =>
      t == null
        ? null
        : typeof t === 'string'
          ? t
          : Array.isArray(t.richText)
            ? t.richText.map((r: Untyped) => r.text).join('')
            : t;
    const sheet = readFixture(rel).worksheets[0];
    const out: Record<string, Untyped> = {};
    if (sheet) {
      for (const {cells} of sheet.rows()) {
        for (const cell of cells) {
          const v = cell.value as Untyped;
          if (v && typeof v === 'object' && 'hyperlink' in v) {
            out[cell.address] = {hyperlink: v.hyperlink ?? null, text: flatten(v.text)};
          }
        }
      }
    }
    return out;
  },

  // Build a workbook with one internal ('#'-prefixed) hyperlink, write it, and report how the link
  // serialized → { writeOk, hasLocation, location, hasExternalRel, hasRid, reloadOk }. An internal
  // target must ride in a `location` attribute with no external-mode relationship, and the package
  // must reload.
  async internalHyperlinkReport(target = "#'Target'!A1") {
    const wb = new Workbook();
    const sheet = wb.addWorksheet('Main');
    wb.addWorksheet('Target');
    sheet.getCell('A1').value = {text: 'go', hyperlink: target};
    let buffer: Untyped;
    try {
      buffer = writeXlsx(wb);
    } catch (e) {
      return {writeOk: false, writeError: messageOf(e)};
    }
    const parts = partMapOf(buffer);
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const relsXml = parts['xl/worksheets/_rels/sheet1.xml.rels'] || '';
    const a = attrsOf((sheetXml.match(/<hyperlink\b[^>]*\/?>/) || [''])[0]);
    let reloadOk = true;
    try {
      readXlsx(buffer);
    } catch {
      reloadOk = false;
    }
    return {
      writeOk: true,
      hasLocation: a.location != null,
      location: a.location ?? null,
      hasExternalRel: /TargetMode="External"/.test(relsXml),
      hasRid: a['r:id'] != null,
      reloadOk,
    };
  },

  // Build a workbook with an internal '#Sheet2!A1' hyperlink (plus a tooltip), write it, and report
  // the serialized distinctions → { hasWorksheetRels, hyperlinkHasRid, hyperlinkLocation,
  // relTargetMode, reReadHyperlink }. The internal form must carry a location and NO external
  // relationship, and the target must survive a reload.
  async internalHyperlinkSerializationReport() {
    const wb = new Workbook();
    const ws = wb.addWorksheet('Sheet1');
    wb.addWorksheet('Sheet2');
    ws.getCell('A1').value = {text: 'go', hyperlink: '#Sheet2!A1', tooltip: 'tt'};
    const buffer = writeXlsx(wb);
    const parts = partMapOf(buffer);
    const sheetXml = parts['xl/worksheets/sheet1.xml'] || '';
    const relsXml = parts['xl/worksheets/_rels/sheet1.xml.rels'] || '';
    const hyperlinkEl = (sheetXml.match(/<hyperlink\b[^>]*\/?>/) || [''])[0];
    const relEl = (relsXml.match(/<Relationship\b[^>]*hyperlink[^>]*\/?>/) || [''])[0];
    const reReadHyperlink =
      (readXlsx(buffer).getWorksheet('Sheet1')!.getCell('A1').value as Untyped).hyperlink ?? null;
    return {
      hasWorksheetRels: /Type="[^"]*\/hyperlink"/.test(relsXml),
      hyperlinkHasRid: /r:id="/.test(hyperlinkEl),
      hyperlinkLocation: (hyperlinkEl.match(/location="([^"]*)"/) || [null, null])[1],
      relTargetMode: (relEl.match(/TargetMode="([^"]*)"/) || [null, null])[1],
      reReadHyperlink,
    };
  },
};
