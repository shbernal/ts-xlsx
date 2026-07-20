// Pure OOXML package-fact extraction, shared by every adapter.
//
// Given a workbook spec and the written package's parts as a plain { path: xmlString }
// map, it derives the same JSON-serializable facts a case asserts on (content-type
// declarations, sheet entries, relationships, per-worksheet cell/formula/margin facts,
// style/theme presence, comment VML, cross-part consistency). Keeping this
// implementation-blind — it knows only OOXML, never how any library is shaped — lets the
// `current` (legacy) and `rewrite` adapters unzip their own way yet return byte-identical
// facts, so a case compares like with like across implementations.

import type {CorpusApi} from '../case.ts';

/** The parts of a written package, keyed by their zip path. */
type PartMap = Record<string, string>;

/** Attribute name → value for one XML tag; absent names read back as `undefined`. */
type Attrs = Record<string, string | undefined>;

const attrs = (tag: string | null | undefined): Attrs => {
  const out: Attrs = {};
  for (const m of String(tag || '').matchAll(/([\w:]+)="([^"]*)"/g)) out[m[1]!] = m[2]!;
  return out;
};

// Cheap structural well-formedness check: a raw & that isn't an entity means a strict
// consumer would choke. A real parser is the reader's concern; here we only need to
// catch an unescaped special leaking into serialized XML.
const xmlWellFormed = (xml: string): boolean =>
  !/&(?!(amp|lt|gt|quot|apos|#\d+|#x[0-9a-fA-F]+);)/.test(xml);

/**
 * @param spec  the workbook spec that produced the package (drives per-sheet lookup)
 * @param partMap  { [zipPath]: xmlString } for every non-directory package part
 */
export function packageFacts(spec: CorpusApi, partMap: PartMap) {
  const parts = Object.keys(partMap).sort();
  const read = (f: string): string | null => (f in partMap ? partMap[f]! : null);

  const contentTypes = read('[Content_Types].xml') || '';
  const workbookXml = read('xl/workbook.xml') || '';
  const relsXml = read('xl/_rels/workbook.xml.rels') || '';

  const worksheetParts = parts.filter((p) => /^xl\/worksheets\/sheet\d+\.xml$/.test(p));
  const wsRelsXml = read('xl/worksheets/_rels/sheet1.xml.rels') || '';
  const worksheetRels = [...wsRelsXml.matchAll(/<Relationship\b[^>]*?\/?>/g)].map((t) => {
    const a = attrs(t[0]);
    return {id: a.Id, target: a.Target, type: (a.Type || '').split('/').pop()};
  });
  const wsRelIds = worksheetRels.map((r) => r.id);
  const overrides = [...contentTypes.matchAll(/<Override[^>]*PartName="([^"]*)"[^>]*\/>/g)].map(
    (m) => m[1]!,
  );
  const contentTypeDefaults = [...contentTypes.matchAll(/<Default\b[^>]*\/>/g)].map((t) => {
    const a = attrs(t[0]);
    return {extension: a.Extension ?? null, contentType: a.ContentType ?? null};
  });
  const sheetEntries = [...workbookXml.matchAll(/<sheet\b[^>]*?\/?>/g)].map((t) => {
    const a = attrs(t[0]);
    return {name: a.name, rid: a['r:id'], state: a.state ?? null};
  });
  // Workbook-level defined names: the element's attributes plus its refersTo text content, sorted by
  // name so a case compares a stable set regardless of emission order.
  const definedNames = [...workbookXml.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)]
    .map((m) => {
      const a = attrs(`<x ${m[1]}>`);
      return {
        name: a.name ?? null,
        localSheetId: a.localSheetId !== undefined ? Number(a.localSheetId) : null,
        hidden: a.hidden === '1' || a.hidden === 'true',
        refersTo: m[2]!,
      };
    })
    .sort((x, y) => String(x.name).localeCompare(String(y.name)));
  const rels = [...relsXml.matchAll(/<Relationship\b[^>]*?\/?>/g)].map((t) => {
    const a = attrs(t[0]);
    return {id: a.Id, target: a.Target, type: (a.Type || '').split('/').pop()};
  });

  const sheets: Record<string, unknown> = {};
  const sheetIndex: Record<string, string> = {};
  (spec.sheets || []).forEach((s: CorpusApi, i: number) => {
    sheetIndex[s.name] = `xl/worksheets/sheet${i + 1}.xml`;
  });
  for (const s of spec.sheets || []) {
    const xml = read(sheetIndex[s.name]!) || '';
    const marginTag = (xml.match(/<pageMargins\b[^>]*\/>/) || [''])[0]!;
    const marginAttrs = attrs(marginTag);
    const sheetViewTags = [...xml.matchAll(/<sheetView\b[^>]*(?:\/>|>)/g)];
    const formulas: Record<string, string> = {};
    for (const m of xml.matchAll(/<c\b[^>]*r="([^"]*)"[^>]*>[\s\S]*?<f\b[^>]*>([\s\S]*?)<\/f>/g)) {
      formulas[m[1]!] = m[2]!;
    }
    const columnGroups = [...xml.matchAll(/<col\b[^>]*\/>/g)].map((t) => {
      const a = attrs(t[0]);
      return {
        min: a.min ? Number(a.min) : null,
        max: a.max ? Number(a.max) : null,
        width: a.width ?? null,
      };
    });
    const posOf = (tag: string): number => xml.indexOf(tag);
    const posDrawing = posOf('<drawing ');
    const posLegacy = posOf('<legacyDrawing ');
    const posTable = posOf('<tableParts');
    const ordered = (a: number, b: number): boolean | null => (a >= 0 && b >= 0 ? a < b : null);
    const hfBlock = (xml.match(
      /<headerFooter\b[\s\S]*?<\/headerFooter>|<headerFooter\b[^>]*\/>/,
    ) || [''])[0]!;
    const hfChild = (tag: string): string | null => {
      const m = hfBlock.match(new RegExp(`<${tag}\\b[^>]*>([\\s\\S]*?)</${tag}>`));
      return m ? m[1]! : null;
    };
    const hfFlag = (name: string): boolean => new RegExp(`\\b${name}="(1|true)"`).test(hfBlock);
    const rowAttrs: Record<string, {outlineLevel: number; hidden: boolean; collapsed: boolean}> =
      {};
    for (const t of xml.matchAll(/<row\b[^>]*>/g)) {
      const a = attrs(t[0]);
      if (a.r === undefined) continue;
      rowAttrs[a.r] = {
        outlineLevel: a.outlineLevel !== undefined ? Number(a.outlineLevel) : 0,
        hidden: a.hidden === '1' || a.hidden === 'true',
        collapsed: a.collapsed === '1' || a.collapsed === 'true',
      };
    }
    sheets[s.name] = {
      pageMargins: {present: Object.keys(marginAttrs), values: marginAttrs},
      hasSheetViews: /<sheetViews>/.test(xml),
      sheetViewCount: sheetViewTags.length,
      hasDimension: /<dimension\b/.test(xml),
      dimensionRef: (xml.match(/<dimension\b[^>]*ref="([^"]*)"/) || [])[1] ?? null,
      autoFilterRef: (xml.match(/<autoFilter\b[^>]*ref="([^"]*)"/) || [])[1] ?? null,
      formulas,
      columnGroups,
      maxColumnIndex: columnGroups.reduce((m, g) => Math.max(m, g.max ?? 0), 0),
      elementOrder: {
        drawing: posDrawing,
        legacyDrawing: posLegacy,
        tableParts: posTable,
        drawingBeforeLegacy: ordered(posDrawing, posLegacy),
        legacyBeforeTableParts: ordered(posLegacy, posTable),
        drawingBeforeTableParts: ordered(posDrawing, posTable),
      },
      headerFooter: {
        present: hfBlock !== '',
        oddHeader: hfChild('oddHeader'),
        oddFooter: hfChild('oddFooter'),
        evenHeader: hfChild('evenHeader'),
        evenFooter: hfChild('evenFooter'),
        firstHeader: hfChild('firstHeader'),
        firstFooter: hfChild('firstFooter'),
        differentOddEven: hfFlag('differentOddEven'),
        differentFirst: hfFlag('differentFirst'),
      },
      rows: rowAttrs,
      hasBackgroundPicture: /<picture\b[^>]*r:id=/.test(xml),
      sheetFormat: (() => {
        const a = attrs((xml.match(/<sheetFormatPr\b[^>]*\/?>/) || [''])[0]!);
        return {
          defaultRowHeight: a.defaultRowHeight != null ? Number(a.defaultRowHeight) : null,
          defaultColWidth: a.defaultColWidth != null ? Number(a.defaultColWidth) : null,
          customHeight: a.customHeight === '1' || a.customHeight === 'true',
        };
      })(),
      xmlWellFormed: xmlWellFormed(xml),
    };
  }

  const tables = [];
  for (const p of parts.filter((f) => /^xl\/tables\/table\d+\.xml$/.test(f))) {
    const xml = read(p) || '';
    const a = attrs((xml.match(/<table\b[^>]*>/) || [''])[0]!);
    const af = xml.match(/<autoFilter\b[^>]*ref="([^"]*)"/);
    tables.push({
      ref: a.ref ?? null,
      name: a.name ?? null,
      autoFilterRef: af ? af[1]! : null,
      columnCount: [...xml.matchAll(/<tableColumn\b/g)].length,
      headerRowCount: a.headerRowCount ?? '1',
      xmlWellFormed: xmlWellFormed(xml),
    });
  }

  const stylesXml = read('xl/styles.xml') || '';
  const defaultFontBlock = (stylesXml.match(/<font>[\s\S]*?<\/font>/) || [''])[0]!;
  const defaultFontColor = attrs((defaultFontBlock.match(/<color\b[^>]*\/?>/) || [''])[0]!);
  const hasThemePart = parts.some((p) => /^xl\/theme\/theme\d+\.xml$/.test(p));
  const styles = {
    hasThemePart,
    defaultFontColor,
    defaultFontUsesTheme: 'theme' in defaultFontColor,
    themeColorResolvable: !('theme' in defaultFontColor) || hasThemePart,
  };

  const vmlTextboxStyles: string[] = [];
  for (const p of parts.filter((f) => /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(f))) {
    const vml = read(p) || '';
    for (const t of vml.matchAll(/<(?:v:)?textbox\b[^>]*\bstyle="([^"]*)"/g))
      vmlTextboxStyles.push(t[1]!);
  }
  const vml = {
    textboxStyles: vmlTextboxStyles,
    allTextboxesFitToText:
      vmlTextboxStyles.length > 0 &&
      vmlTextboxStyles.every((s) => /mso-fit-shape-to-text\s*:\s*t/i.test(s)),
  };

  const declaredConsistent = worksheetParts.every((part) => {
    const over = overrides.includes(`/${part}`);
    const rid = rels.find(
      (r) =>
        `xl/${r.target}`.replace('xl/xl/', 'xl/') === part || r.target === part.replace('xl/', ''),
    );
    return over && !!rid;
  });

  return {
    parts,
    worksheetParts,
    overrides,
    contentTypeDefaults,
    sheetEntries,
    definedNames,
    rels,
    worksheetRels,
    sheets,
    tables,
    styles,
    vml,
    packageParts: {
      hasCommentsPart: parts.some((p) => /^xl\/comments\d+\.xml$/.test(p)),
      hasVmlDrawingPart: parts.some((p) => /^xl\/drawings\/vmlDrawing\d+\.vml$/.test(p)),
      hasTablePart: parts.some((p) => /^xl\/tables\/table\d+\.xml$/.test(p)),
    },
    consistency: {
      worksheetPartCount: worksheetParts.length,
      sheetEntryCount: sheetEntries.length,
      declaredConsistent,
      worksheetRelIdsUnique: new Set(wsRelIds).size === wsRelIds.length,
    },
  };
}
