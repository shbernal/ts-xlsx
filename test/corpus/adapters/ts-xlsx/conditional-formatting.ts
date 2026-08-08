// Conditional formatting rules and their evaluation order.

import fs from 'node:fs';
import path from 'node:path';
import {messageOf} from '../../thrown.ts';
import type {Untyped} from '../../untyped.ts';
import {partMapOf} from './package-facts.ts';
import {FIXTURES_ROOT, readFixture, readXlsx, Workbook, writeXlsx} from './runtime.ts';
import {attrsOf} from './xml-probes.ts';

export const conditionalFormatting = {
  // Author a conditional-formatting rule, write it, and report the emitted CF XML facts plus what the
  // reader surfaces on reload → { writeOk, writeError, xml:{blockCount, sqrefs, ruleCount, hasDataBar,
  // cfvoCount, hasColor, wellFormed}, reload:{type, color, gradient, cfvo} }.
  authorConditionalFormatting(cf: Untyped) {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    // Populate the ref's first column so the rule binds to real cells.
    const rows = Number((cf.ref.match(/(\d+)\s*$/) || [])[1] || 3);
    for (let r = 1; r <= rows; r++) sheet.getCell(`A${r}`).value = r / rows;
    let buffer: Uint8Array;
    try {
      sheet.addConditionalFormatting(cf);
      buffer = writeXlsx(workbook);
    } catch (e) {
      return {
        writeOk: false,
        writeError: messageOf(e),
        xml: null,
        reload: null,
      };
    }
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const cfBlock = (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/) || [
      '',
    ])[0];
    const dataBar = (cfBlock.match(/<dataBar\b[\s\S]*?<\/dataBar>|<dataBar\b[^>]*\/>/) || [''])[0];
    const rule =
      readXlsx(buffer).getWorksheet('S')?.conditionalFormattings?.[0]?.rules?.[0] ?? null;
    return {
      writeOk: true,
      writeError: null,
      xml: {
        blockCount: [...xml.matchAll(/<conditionalFormatting\b/g)].length,
        sqrefs: [...xml.matchAll(/<conditionalFormatting\b[^>]*sqref="([^"]*)"/g)].map((m) => m[1]),
        ruleCount: [...cfBlock.matchAll(/<cfRule\b/g)].length,
        hasDataBar: /<dataBar\b/.test(cfBlock),
        cfvoCount: [...dataBar.matchAll(/<cfvo\b/g)].length,
        hasColor: /<color\b/.test(dataBar),
        wellFormed: cfBlock
          ? !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(cfBlock)
          : false,
      },
      reload: rule
        ? {
            type: rule.type ?? null,
            color: rule.color ? (rule.color.argb ?? null) : null,
            gradient: rule.gradient ?? null,
            cfvo: (rule.cfvo || []).map((v) => ({type: v.type ?? null, value: v.value ?? null})),
          }
        : null,
    };
  },

  // Apply a stopIfTrue rule, write, reload → { xmlHasStopIfTrue, reloadStopIfTrue }.
  conditionalFormattingStopIfTrue() {
    const workbook = new Workbook();
    const sheet = workbook.addWorksheet('S');
    sheet.getCell('A1').value = 5;
    sheet.addConditionalFormatting({
      ref: 'A1:A10',
      rules: [
        {
          type: 'cellIs',
          operator: 'greaterThan',
          formulae: [3],
          stopIfTrue: true,
          style: {fill: {type: 'pattern', pattern: 'solid', bgColor: {argb: 'FFFF0000'}}},
        },
      ],
    });
    const buffer = writeXlsx(workbook);
    const xml = partMapOf(buffer)['xl/worksheets/sheet1.xml'] || '';
    const rule = readXlsx(buffer).getWorksheet('S')?.conditionalFormattings?.[0]?.rules?.[0];
    return {
      xmlHasStopIfTrue: /stopIfTrue="1"/.test(xml),
      reloadStopIfTrue: rule ? (rule.stopIfTrue ?? false) : null,
    };
  },

  // Read a fixture's first-sheet conditional-formatting facts, write it back, and report the same
  // before/after → { source, rewritten } each { blockCount, rules:[{type, dxfId, priority}] }.
  roundtripFixtureConditionalFormatting(rel: string) {
    const cfFacts = (xml: string) => ({
      blockCount: [...xml.matchAll(/<conditionalFormatting\b/g)].length,
      rules: [...xml.matchAll(/<cfRule\b([^>]*?)\/?>/g)].map((m) => {
        const a = attrsOf(`<x ${m[1]}>`);
        return {type: a.type ?? null, dxfId: a.dxfId ?? null, priority: a.priority ?? null};
      }),
    });
    const srcParts = partMapOf(fs.readFileSync(path.join(FIXTURES_ROOT, rel)));
    const srcName = Object.keys(srcParts).find((n) => /sheet1\.xml$/.test(n));
    const source = cfFacts(srcName === undefined ? '' : (srcParts[srcName] ?? ''));
    const outXml = partMapOf(writeXlsx(readFixture(rel)))['xl/worksheets/sheet1.xml'] || '';
    return {source, rewritten: cfFacts(outXml)};
  },
};
