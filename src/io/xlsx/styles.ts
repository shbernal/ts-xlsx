// The write-side style table: the interned `<numFmts>`/`<fills>`/`<cellXfs>` backing styles.xml.
//
// OOXML styles are a *shared* table referenced by index: a cell (or a formatted row/column)
// names a `<cellXfs>` entry via its `s` attribute, and that entry names a fill by id and a
// number format by id. Identical styles must collapse to one entry — both to produce
// well-formed OOXML and to keep write cost bounded on large, lightly-formatted sheets (the
// historical performance cliff came from re-serialising a distinct style per cell). The
// registry interns each distinct fill, number format, and xf, handing back a stable index.
//
// Fills, number formats, fonts, borders, alignment, and protection are modelled today.
// Fills/fonts/borders are shared sub-tables the xf names by id, whereas alignment and protection
// are child elements *of* the xf — so each is interned into the xf signature directly rather than
// into its own id table, and an aligned/protected xf carries them as body children in that order.
// An unstyled cell/row/column resolves to xf 0.

import type {DifferentialStyle} from '../../core/conditional-formatting.ts';
import type {
  Alignment,
  Border,
  BorderEdge,
  Color,
  Fill,
  Font,
  GradientFill,
  NamedCellStyle,
  Protection,
  UnderlineStyle,
} from '../../core/style.ts';
import {escapeAttr, XML_DECLARATION} from './xml.ts';

// Excel reserves fill ids 0 and 1 for the "none" and "gray125" patterns it always emits;
// custom fills are numbered from 2 so a foreign reader's built-in assumptions still hold.
const RESERVED_FILL_COUNT = 2;

// Font id 0 is the always-present default font (Calibri 11); custom fonts are numbered from 1.
const RESERVED_FONT_COUNT = 1;

// numFmt ids below 164 are reserved by ECMA-376 for the built-in formats every consumer
// knows implicitly; custom format codes are numbered from 164 up. Id 0 is General (no code).
const CUSTOM_NUMFMT_BASE = 164;

// Border id 0 is the always-present empty border (every edge absent); custom borders from 1.
const RESERVED_BORDER_COUNT = 1;

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
// The default font's inner fragment, in the exact child order `fontXml` emits. The reader
// surfaces this face on every otherwise-unstyled cell (font id 0 is a real font, not an
// absence), so a cell carrying exactly the default must intern back to id 0 rather than a
// redundant custom entry — keeping a read→write round-trip byte-stable.
const DEFAULT_FONT_BODY =
  '<sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>';
const DEFAULT_FONT = `<font>${DEFAULT_FONT_BODY}</font>`;
// The empty border: all five edges present but styleless. A border that overrides no edge
// serialises to exactly this, so it interns to the default border id 0 rather than a new one.
const DEFAULT_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';

/** A cell's style facets as the writer composes them: cell overrides atop row/column defaults. */
export interface CellStyle {
  readonly fill?: Fill | undefined;
  readonly numFmt?: string | undefined;
  readonly font?: Partial<Font> | undefined;
  readonly border?: Border | undefined;
  readonly alignment?: Alignment | undefined;
  readonly protection?: Protection | undefined;
  /** The quote-prefix flag — an attribute on the xf, not a shared sub-table entry. */
  readonly quotePrefix?: boolean | undefined;
  /** The `xfId` link into `cellStyleXfs` — the named cell style this format inherits from (0 = Normal). */
  readonly xfId?: number | undefined;
}

// One interned cell format. `fillId` 0 is no fill; `numFmtId` 0 is the General format;
// `fontId` 0 is the default font; `borderId` 0 is the empty border. `alignment` and `protection`
// hold the serialised `<alignment>`/`<protection>` attribute strings (empty when the cell has no
// explicit facet), carried inline because both are children of the xf rather than shared sub-tables.
interface CellFormat {
  readonly fillId: number;
  readonly numFmtId: number;
  readonly fontId: number;
  readonly borderId: number;
  readonly alignment: string;
  readonly protection: string;
  readonly quotePrefix: boolean;
  // The named-style link (`xfId`) a cellXfs entry carries; 0 = Normal. A cellStyleXfs entry does not
  // itself nest, so it is serialised with this omitted.
  readonly xfId: number;
}

// The default xf: no facet, General format, linked to the Normal named style (xfId 0). Shared as the
// first entry of both the cell-format and named-style tables; never mutated (formats only append).
const DEFAULT_FORMAT: CellFormat = {
  fillId: 0,
  numFmtId: 0,
  fontId: 0,
  borderId: 0,
  alignment: '',
  protection: '',
  quotePrefix: false,
  xfId: 0,
};

export class StyleRegistry {
  // Custom fill xml fragments, in id order; the emitted id is RESERVED_FILL_COUNT + index.
  readonly #fillXml: string[] = [];
  readonly #fillIdBySignature = new Map<string, number>();

  // Custom number-format codes, in id order; the emitted id is CUSTOM_NUMFMT_BASE + index.
  readonly #numFmtCodes: string[] = [];
  readonly #numFmtIdByCode = new Map<string, number>();

  // Custom font xml fragments, in id order; the emitted id is RESERVED_FONT_COUNT + index.
  readonly #fontXml: string[] = [];
  readonly #fontIdBySignature = new Map<string, number>();

  // Custom border xml fragments, in id order; the emitted id is RESERVED_BORDER_COUNT + index.
  readonly #borderXml: string[] = [];
  readonly #borderIdBySignature = new Map<string, number>();

  // xf 0 is the default (no fill/font/border/alignment/protection, General format); further entries append as styles appear.
  readonly #formats: CellFormat[] = [DEFAULT_FORMAT];
  readonly #xfIndexBySignature = new Map<string, number>();

  // The named-style layer (`<cellStyleXfs>` / `<cellStyles>`): the base formats a cell's `xfId` links
  // into, and the names that label them. Index 0 is always Normal. A file with named styles seeds this
  // in place of the default via {@link seedNamedStyles}; otherwise the default alone is emitted.
  readonly #cellStyleXfs: CellFormat[] = [DEFAULT_FORMAT];
  readonly #cellStyleNames: {name: string; builtinId?: number; xfId: number}[] = [
    {name: 'Normal', builtinId: 0, xfId: 0},
  ];

  // Differential styles (`<dxfs>`) that conditional formatting references by index. Fragments read
  // from a file are seeded first and kept verbatim so a foreign rule's dxfId stays valid; a style
  // authored on a rule is serialised and appended after them, dedup'd by its fragment.
  readonly #dxfXml: string[] = [];
  readonly #dxfIndexByFragment = new Map<string, number>();

  // A custom indexed-color palette (`<colors><indexedColors>`) read from a file, each entry a verbatim
  // `<rgbColor rgb="…"/>`. Preserved and re-emitted unchanged so cells/fonts/borders that reference a
  // colour by `indexed="…"` keep their intended RGB; dropping it would silently resolve every indexed
  // colour to a different default-palette entry. Empty for a workbook that never overrode the palette.
  readonly #indexedColors: string[] = [];

  /**
   * The `<cellXfs>` index for a composed cell/row/column style. A style with no facet needs
   * no entry and resolves to the default xf 0, so its owner emits no `s` attribute at all.
   */
  styleId(style: CellStyle): number {
    const format = this.#composeFormat(style, style.xfId ?? 0);
    // An all-default format that links to no named style needs no entry and resolves to xf 0, so its
    // owner emits no `s` attribute. A non-zero xfId is itself information — the cell inherits a named
    // style — so it forces a real entry even when the direct facets are empty.
    if (isDefaultFormat(format)) return 0;

    const signature = formatSignature(format);
    let index = this.#xfIndexBySignature.get(signature);
    if (index === undefined) {
      index = this.#formats.length;
      this.#formats.push(format);
      this.#xfIndexBySignature.set(signature, index);
    }
    return index;
  }

  // Compose a style's facets into an interned {@link CellFormat}, interning each fill/font/border/
  // number-format into its shared sub-table. Shared by the cell-format path ({@link styleId}) and the
  // named-style path ({@link seedNamedStyles}), which differ only in which table the result lands in.
  #composeFormat(style: CellStyle, xfId: number): CellFormat {
    // A `none` pattern is the reserved fill 0; a gradient is always a real, interned fill.
    const paints =
      style.fill !== undefined && (style.fill.type === 'gradient' || style.fill.pattern !== 'none');
    const fillId = paints ? this.#internFill(style.fill as Fill) : 0;
    // A number format is a format-code *string*; a caller that assigns a structured object (e.g. a
    // parsed `{id, formatCode}` copied from another cell) must not have it stringified into the styles
    // part as `formatCode="[object Object]"`, which Excel reports as a corrupt package. A non-string
    // format is dropped to the General format rather than corrupting the file.
    const numFmtId =
      typeof style.numFmt === 'string' && style.numFmt !== ''
        ? this.#internNumFmt(style.numFmt)
        : 0;
    const fontId = style.font ? this.#internFont(style.font) : 0;
    const borderId = style.border ? this.#internBorder(style.border) : 0;
    const alignment = style.alignment ? alignmentAttrs(style.alignment) : '';
    const protection = style.protection ? protectionAttrs(style.protection) : '';
    const quotePrefix = style.quotePrefix === true;
    return {fillId, numFmtId, fontId, borderId, alignment, protection, quotePrefix, xfId};
  }

  /**
   * Seed the named cell styles (`<cellStyleXfs>`/`<cellStyles>`) read from a file, in place of the
   * lone default, interning each style's facets into the shared sub-tables so its `fillId`/`fontId`/…
   * references stay valid against the rebuilt tables. Index 0 stays Normal. A cell's `xfId` indexes
   * this table, so it must be seeded before any {@link styleId} that carries an `xfId`.
   */
  seedNamedStyles(styles: readonly NamedCellStyle[]): void {
    if (styles.length === 0) return;
    this.#cellStyleXfs.length = 0;
    this.#cellStyleNames.length = 0;
    styles.forEach((style, index) => {
      this.#cellStyleXfs.push(this.#composeFormat(style, 0));
      const entry: {name: string; builtinId?: number; xfId: number} = {
        name: style.name ?? `Style ${index}`,
        xfId: index,
      };
      if (style.builtinId !== undefined) entry.builtinId = style.builtinId;
      this.#cellStyleNames.push(entry);
    });
  }

  /**
   * Seed the differential-style table with fragments read from a file, keeping each `<dxf>…</dxf>`
   * verbatim and at its original index so a conditional-formatting rule's `dxfId` still resolves. Call
   * once before any {@link differentialStyleId}; authored styles append after these.
   */
  seedDifferentialStyles(fragments: readonly string[]): void {
    for (const fragment of fragments) {
      const index = this.#dxfXml.length;
      this.#dxfXml.push(fragment);
      // A seeded fragment can still be reused by an authored style identical to it, so index it too.
      if (!this.#dxfIndexByFragment.has(fragment)) this.#dxfIndexByFragment.set(fragment, index);
    }
  }

  /**
   * Seed the custom indexed-color palette (`<colors><indexedColors>`) read from a file, each entry a
   * verbatim `<rgbColor rgb="…"/>` fragment. Re-emitting it unchanged is what keeps an `indexed="…"`
   * colour reference resolving to the RGB the source intended. An empty list leaves the workbook on
   * the default palette and emits no `<colors>` element.
   */
  seedIndexedColors(fragments: readonly string[]): void {
    this.#indexedColors.length = 0;
    this.#indexedColors.push(...fragments);
  }

  /** Intern a differential style authored on a rule, returning its `<dxfs>` index for the cfRule's
   * `dxfId`. Identical styles collapse to one entry. */
  differentialStyleId(style: DifferentialStyle): number {
    const fragment = dxfXml(style);
    let index = this.#dxfIndexByFragment.get(fragment);
    if (index === undefined) {
      index = this.#dxfXml.length;
      this.#dxfXml.push(fragment);
      this.#dxfIndexByFragment.set(fragment, index);
    }
    return index;
  }

  #internFill(fill: Fill): number {
    const signature = fillSignature(fill);
    let id = this.#fillIdBySignature.get(signature);
    if (id === undefined) {
      id = RESERVED_FILL_COUNT + this.#fillXml.length;
      this.#fillXml.push(fillXml(fill));
      this.#fillIdBySignature.set(signature, id);
    }
    return id;
  }

  #internNumFmt(code: string): number {
    let id = this.#numFmtIdByCode.get(code);
    if (id === undefined) {
      id = CUSTOM_NUMFMT_BASE + this.#numFmtCodes.length;
      this.#numFmtCodes.push(code);
      this.#numFmtIdByCode.set(code, id);
    }
    return id;
  }

  // A font whose partial carries no facet that differs from the default contributes nothing
  // and maps to font id 0; otherwise its serialised form is interned and dedup'd like a fill.
  #internFont(font: Partial<Font>): number {
    const xml = fontXml(font);
    if (xml === '' || xml === DEFAULT_FONT_BODY) return 0;
    let id = this.#fontIdBySignature.get(xml);
    if (id === undefined) {
      id = RESERVED_FONT_COUNT + this.#fontXml.length;
      this.#fontXml.push(`<font>${xml}</font>`);
      this.#fontIdBySignature.set(xml, id);
    }
    return id;
  }

  // A border that overrides no edge serialises to the empty default border and maps to id 0;
  // otherwise its serialised form is interned and dedup'd like a fill or font.
  #internBorder(border: Border): number {
    const xml = borderXml(border);
    if (xml === DEFAULT_BORDER) return 0;
    let id = this.#borderIdBySignature.get(xml);
    if (id === undefined) {
      id = RESERVED_BORDER_COUNT + this.#borderXml.length;
      this.#borderXml.push(xml);
      this.#borderIdBySignature.set(xml, id);
    }
    return id;
  }

  /** Serialise the accumulated table into a complete, valid styles.xml part. */
  toXml(): string {
    const fillCount = RESERVED_FILL_COUNT + this.#fillXml.length;
    const fills =
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      this.#fillXml.join('');
    const cellXfs = this.#formats.map((format) => xfXml(format, format.xfId)).join('');
    const cellStyleXfs = this.#cellStyleXfs.map((format) => xfXml(format, null)).join('');
    const cellStyles = this.#cellStyleNames.map(cellStyleTag).join('');
    const fontCount = RESERVED_FONT_COUNT + this.#fontXml.length;
    const fonts = DEFAULT_FONT + this.#fontXml.join('');
    const borderCount = RESERVED_BORDER_COUNT + this.#borderXml.length;
    const borders = DEFAULT_BORDER + this.#borderXml.join('');
    return (
      XML_DECLARATION +
      `<styleSheet xmlns="${NS_MAIN}">` +
      this.#numFmtsXml() +
      `<fonts count="${fontCount}">${fonts}</fonts>` +
      `<fills count="${fillCount}">${fills}</fills>` +
      `<borders count="${borderCount}">${borders}</borders>` +
      `<cellStyleXfs count="${this.#cellStyleXfs.length}">${cellStyleXfs}</cellStyleXfs>` +
      `<cellXfs count="${this.#formats.length}">${cellXfs}</cellXfs>` +
      `<cellStyles count="${this.#cellStyleNames.length}">${cellStyles}</cellStyles>` +
      this.#dxfsXml() +
      this.#colorsXml() +
      '</styleSheet>'
    );
  }

  // <colors> (holding the custom <indexedColors> palette) is a late child of <styleSheet>, after
  // <dxfs> and <tableStyles>. It is emitted only when a file overrode the default palette, so an
  // ordinary workbook stays on the built-in indexed colours and writes no <colors> element.
  #colorsXml(): string {
    if (this.#indexedColors.length === 0) return '';
    return `<colors><indexedColors>${this.#indexedColors.join('')}</indexedColors></colors>`;
  }

  // <dxfs> holds the differential styles conditional formatting references by index. An empty table
  // is still emitted as a self-closing count="0" element, the shape Excel writes; a populated one
  // lists the seeded (foreign) fragments first, then any authored styles, preserving every index.
  #dxfsXml(): string {
    if (this.#dxfXml.length === 0) return '<dxfs count="0"/>';
    return `<dxfs count="${this.#dxfXml.length}">${this.#dxfXml.join('')}</dxfs>`;
  }

  // <numFmts> is the first child of <styleSheet> and is omitted entirely when no custom
  // format was used, matching how Excel writes an all-built-in workbook.
  #numFmtsXml(): string {
    if (this.#numFmtCodes.length === 0) return '';
    const entries = this.#numFmtCodes
      .map(
        (code, i) =>
          `<numFmt numFmtId="${CUSTOM_NUMFMT_BASE + i}" formatCode="${escapeFormatCode(code)}"/>`,
      )
      .join('');
    return `<numFmts count="${this.#numFmtCodes.length}">${entries}</numFmts>`;
  }
}

// Whether a format is the do-nothing default: no facet, General number format, no quote prefix, and
// linked to the Normal named style. Such a cellXfs entry adds nothing, so its owner needs no `s`.
function isDefaultFormat(format: CellFormat): boolean {
  return (
    format.fillId === 0 &&
    format.numFmtId === 0 &&
    format.fontId === 0 &&
    format.borderId === 0 &&
    format.alignment === '' &&
    format.protection === '' &&
    !format.quotePrefix &&
    format.xfId === 0
  );
}

// A stable, collision-free key for a composed format so identical formats intern to one cellXfs entry.
function formatSignature(format: CellFormat): string {
  return (
    `fill:${format.fillId}|numFmt:${format.numFmtId}|font:${format.fontId}|border:${format.borderId}|` +
    `align:${format.alignment}|protect:${format.protection}|quote:${format.quotePrefix}|xfId:${format.xfId}`
  );
}

// Serialise one `<xf>`. A cellXfs entry passes its named-style link as `xfId`; a cellStyleXfs entry
// (the base a cell links *to*) passes `null` so the attribute is omitted, since it nests no further.
function xfXml(format: CellFormat, xfId: number | null): string {
  const applyNumberFormat = format.numFmtId !== 0 ? ' applyNumberFormat="1"' : '';
  const applyFont = format.fontId !== 0 ? ' applyFont="1"' : '';
  const applyFill = format.fillId !== 0 ? ' applyFill="1"' : '';
  const applyBorder = format.borderId !== 0 ? ' applyBorder="1"' : '';
  const applyAlignment = format.alignment !== '' ? ' applyAlignment="1"' : '';
  const applyProtection = format.protection !== '' ? ' applyProtection="1"' : '';
  // `quotePrefix` is a CT_Xf attribute (after xfId, before the apply flags in schema order); it is
  // its own switch — there is no `applyQuotePrefix` flag — so it is emitted only when set.
  const quotePrefix = format.quotePrefix ? ' quotePrefix="1"' : '';
  const xfIdAttr = xfId === null ? '' : ` xfId="${xfId}"`;
  const open =
    `<xf numFmtId="${format.numFmtId}" fontId="${format.fontId}" fillId="${format.fillId}" ` +
    `borderId="${format.borderId}"${xfIdAttr}${quotePrefix}` +
    `${applyNumberFormat}${applyFont}${applyFill}${applyBorder}${applyAlignment}${applyProtection}`;
  // Alignment and protection are child elements of the xf, in that schema order; an xf carrying
  // either (or both) is not self-closing, while a plain one stays self-closing as before.
  const body =
    (format.alignment === '' ? '' : `<alignment ${format.alignment}/>`) +
    (format.protection === '' ? '' : `<protection ${format.protection}/>`);
  return body === '' ? `${open}/>` : `${open}>${body}</xf>`;
}

// One `<cellStyle>` entry mapping a name (and, for a built-in, its gallery id) to a cellStyleXfs index.
function cellStyleTag(entry: {name: string; builtinId?: number; xfId: number}): string {
  const builtin = entry.builtinId === undefined ? '' : ` builtinId="${entry.builtinId}"`;
  return `<cellStyle name="${escapeAttr(entry.name)}" xfId="${entry.xfId}"${builtin}/>`;
}

// Serialise a cell's alignment as `<alignment>` attributes in ECMA-376 CT_CellAlignment order.
// A facet at its default contributes nothing; an all-default alignment yields the empty string,
// so it forces neither an <alignment> child nor a distinct xf.
function alignmentAttrs(alignment: Alignment): string {
  const parts: string[] = [];
  // `general` is the type-dependent default and is expressed by omitting the attribute.
  if (alignment.horizontal !== undefined && alignment.horizontal !== 'general') {
    parts.push(`horizontal="${alignment.horizontal}"`);
  }
  if (alignment.vertical !== undefined) parts.push(`vertical="${alignment.vertical}"`);
  if (alignment.textRotation !== undefined && alignment.textRotation !== 0) {
    parts.push(`textRotation="${numberAttr(alignment.textRotation)}"`);
  }
  if (alignment.wrapText) parts.push('wrapText="1"');
  if (alignment.indent !== undefined && alignment.indent !== 0) {
    parts.push(`indent="${numberAttr(alignment.indent)}"`);
  }
  if (alignment.shrinkToFit) parts.push('shrinkToFit="1"');
  if (alignment.readingOrder !== undefined && alignment.readingOrder !== 0) {
    parts.push(`readingOrder="${numberAttr(alignment.readingOrder)}"`);
  }
  return parts.join(' ');
}

// Serialise a cell's protection as `<protection>` attributes. `locked` defaults to true in OOXML,
// so only an explicitly unlocked cell writes `locked="0"`; `hidden` defaults to false, so only a
// hidden cell writes `hidden="1"`. An all-default protection yields the empty string, forcing
// neither a <protection> child nor a distinct xf.
function protectionAttrs(protection: Protection): string {
  const parts: string[] = [];
  if (protection.locked === false) parts.push('locked="0"');
  if (protection.hidden === true) parts.push('hidden="1"');
  return parts.join(' ');
}

/**
 * Extract the custom indexed-color palette (`<colors><indexedColors>`) from styles.xml as verbatim
 * `<rgbColor rgb="…"/>` fragments, or an empty list when the file rides the default palette. Kept raw
 * — rather than parsed into RGB and re-serialised — so the exact entries (count, order, casing) a
 * source file declared survive a round-trip and every `indexed="…"` reference keeps its RGB.
 */
export function parseIndexedColors(stylesXml: string): string[] {
  const block = /<indexedColors\b[^>]*>([\s\S]*?)<\/indexedColors>/.exec(stylesXml);
  if (block === null) return [];
  const inner = block[1] ?? '';
  return [...inner.matchAll(/<rgbColor\b[^>]*\/>|<rgbColor\b[^>]*>[\s\S]*?<\/rgbColor>/g)].map(
    (m) => m[0] ?? '',
  );
}

// Serialise the facets a font overrides, in ECMA-376 child order. A boolean flag is emitted only
// when true (its absence is the default false); an empty result means the font differs from the
// default in nothing and needs no entry at all. The face element differs by context: a styles
// `<font>` names it `<name>` (CT_Font), a rich-text run's `<rPr>` names it `<rFont>` (CT_RPrElt) —
// otherwise the two share every child, so `nameTag` selects the face element and the rest is common.
export function fontXml(font: Partial<Font>, nameTag: 'name' | 'rFont' = 'name'): string {
  const parts: string[] = [];
  if (font.bold) parts.push('<b/>');
  if (font.italic) parts.push('<i/>');
  if (font.strike) parts.push('<strike/>');
  if (font.outline) parts.push('<outline/>');
  const underline = underlineXml(font.underline);
  if (underline !== '') parts.push(underline);
  if (font.vertAlign !== undefined) parts.push(`<vertAlign val="${font.vertAlign}"/>`);
  if (font.size !== undefined) parts.push(`<sz val="${numberAttr(font.size)}"/>`);
  if (font.color !== undefined) parts.push(`<color ${colorAttrs(font.color)}/>`);
  if (font.name !== undefined) parts.push(`<${nameTag} val="${escapeAttr(font.name)}"/>`);
  if (font.family !== undefined) parts.push(`<family val="${numberAttr(font.family)}"/>`);
  if (font.charset !== undefined) parts.push(`<charset val="${numberAttr(font.charset)}"/>`);
  if (font.scheme !== undefined && font.scheme !== 'none')
    parts.push(`<scheme val="${font.scheme}"/>`);
  return parts.join('');
}

// Serialise a differential style (CT_Dxf) in schema child order: font, numFmt, fill, border. Only the
// facets present are emitted — a dxf overrides exactly what it names and lets the cell's own style show
// through the rest. A dxf's pattern fill states the highlight through `bgColor`, matching how Excel
// writes a "fill with colour" conditional format.
export function dxfXml(style: DifferentialStyle): string {
  const parts: string[] = [];
  if (style.font !== undefined) {
    const font = fontXml(style.font);
    if (font !== '') parts.push(`<font>${font}</font>`);
  }
  // A dxf numFmt still needs an id; the code is what matters (dxf formats are not shared by id like
  // cell formats), so a fixed custom id carries it without a <numFmts> entry.
  if (typeof style.numFmt === 'string' && style.numFmt !== '') {
    parts.push(
      `<numFmt numFmtId="${CUSTOM_NUMFMT_BASE}" formatCode="${escapeFormatCode(style.numFmt)}"/>`,
    );
  }
  if (style.fill !== undefined) parts.push(dxfFillXml(style.fill));
  if (style.border !== undefined) parts.push(borderXml(style.border));
  return `<dxf>${parts.join('')}</dxf>`;
}

// A differential fill: the pattern with whatever colours it names. Unlike a cell fill, a dxf does not
// force a placeholder background on a solid pattern — a dxf states only the overrides it carries.
function dxfFillXml(fill: Fill): string {
  if (fill.type === 'gradient') return `<fill>${gradientFillXml(fill)}</fill>`;
  const fg = fill.fgColor ? `<fgColor ${colorAttrs(fill.fgColor)}/>` : '';
  const bg = fill.bgColor ? `<bgColor ${colorAttrs(fill.bgColor)}/>` : '';
  return `<fill><patternFill patternType="${fill.pattern}">${fg}${bg}</patternFill></fill>`;
}

// The gradient element shared by cell fills and dxf fills. Linear gradients carry a `degree`; path
// gradients carry inner-rectangle insets. A zero-valued attribute is its OOXML default and is omitted.
function gradientFillXml(fill: GradientFill): string {
  const attrs =
    (fill.gradient === 'path' ? ' type="path"' : '') +
    (fill.degree ? ` degree="${numberAttr(fill.degree)}"` : '') +
    insetAttr('left', fill.left) +
    insetAttr('right', fill.right) +
    insetAttr('top', fill.top) +
    insetAttr('bottom', fill.bottom);
  const stops = fill.stops
    .map(
      (stop) =>
        `<stop position="${numberAttr(stop.position)}"><color ${colorAttrs(stop.color)}/></stop>`,
    )
    .join('');
  return `<gradientFill${attrs}>${stops}</gradientFill>`;
}

function insetAttr(name: string, value: number | undefined): string {
  return value ? ` ${name}="${numberAttr(value)}"` : '';
}

// `<u/>` is single underline (the same as an explicit "single"); the named variants carry a
// val; false and "none" are the default no-underline and emit nothing.
function underlineXml(underline: UnderlineStyle | undefined): string {
  if (underline === undefined || underline === false || underline === 'none') return '';
  if (underline === true || underline === 'single') return '<u/>';
  return `<u val="${underline}"/>`;
}

function numberAttr(value: number): string {
  if (!Number.isFinite(value)) {
    throw new Error(`cannot serialise a non-finite font metric (${value})`);
  }
  return String(value);
}

// Serialise a border in ECMA-376 CT_Border child order (left, right, top, bottom, diagonal).
// Every edge element is always present — a styleless `<left/>` is how OOXML says "no left
// border" — so an all-absent border round-trips to the empty default rather than a new id.
function borderXml(border: Border): string {
  const attrs =
    (border.diagonalUp ? ' diagonalUp="1"' : '') + (border.diagonalDown ? ' diagonalDown="1"' : '');
  return (
    `<border${attrs}>` +
    edgeXml('left', border.left) +
    edgeXml('right', border.right) +
    edgeXml('top', border.top) +
    edgeXml('bottom', border.bottom) +
    edgeXml('diagonal', border.diagonal) +
    '</border>'
  );
}

// One border edge: a styleless self-closing tag when absent, else the style attribute plus an
// optional colour child.
function edgeXml(tag: string, edge: BorderEdge | undefined): string {
  if (edge === undefined) return `<${tag}/>`;
  if (edge.color === undefined) return `<${tag} style="${edge.style}"/>`;
  return `<${tag} style="${edge.style}"><color ${colorAttrs(edge.color)}/></${tag}>`;
}

// A format code sits in the `formatCode` attribute; only the markup-significant characters
// need escaping. A code can legitimately contain `"` (quoted literals like `"$"`), `<`, `&`.
function escapeFormatCode(code: string): string {
  return code
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// A stable, collision-free key for a fill: identical fills share it, distinct ones don't.
function fillSignature(fill: Fill): string {
  if (fill.type === 'gradient') {
    const stops = fill.stops.map((s) => `${s.position}:${colorSignature(s.color)}`).join(',');
    return `grad|${fill.gradient}|${fill.degree ?? ''}|${fill.left ?? ''}/${fill.right ?? ''}/${fill.top ?? ''}/${fill.bottom ?? ''}|${stops}`;
  }
  return `${fill.pattern}|${colorSignature(fill.fgColor)}|${colorSignature(fill.bgColor)}`;
}

function colorSignature(color: Color | undefined): string {
  if (color === undefined) return '';
  return `${color.argb ?? ''}/${color.theme ?? ''}/${color.tint ?? ''}/${color.indexed ?? ''}`;
}

function fillXml(fill: Fill): string {
  if (fill.type === 'gradient') return `<fill>${gradientFillXml(fill)}</fill>`;
  const fg = fill.fgColor ? `<fgColor ${colorAttrs(fill.fgColor)}/>` : '';
  // A solid fill's background is the automatic indexed placeholder unless one is stated;
  // omitting it makes Excel render the fill as flat black, so it is always emitted.
  const bg = fill.bgColor
    ? `<bgColor ${colorAttrs(fill.bgColor)}/>`
    : fill.pattern === 'solid'
      ? '<bgColor indexed="64"/>'
      : '';
  return `<fill><patternFill patternType="${fill.pattern}">${fg}${bg}</patternFill></fill>`;
}

// OOXML wants a bare 8-hex ARGB (alpha + RGB). This single choke point — through which every
// fill/font/border/tab colour flows — accepts two developer conveniences and rejects the rest loudly,
// because a malformed rgb value does not error in Excel: it silently renders as flat black.
//   - A leading '#' is a CSS habit and is stripped ('#FFBFBFBF' → 'FFBFBFBF').
//   - A 6-hex RGB is promoted to ARGB with a fully-opaque alpha ('00FF00' → 'FF00FF00'), the common
//     case of a colour written without its alpha channel.
// Anything not then exactly 8 hex digits is a programming error at the API surface, so it throws with
// the offending value rather than writing corrupt XML. Casing is preserved so foreign files round-trip.
function normalizeArgb(argb: string): string {
  const hex = argb.startsWith('#') ? argb.slice(1) : argb;
  const rgb = hex.length === 6 ? `FF${hex}` : hex;
  if (!/^[0-9a-fA-F]{8}$/.test(rgb)) {
    throw new Error(
      `Invalid ARGB colour ${JSON.stringify(argb)}: expected 6 or 8 hexadecimal digits`,
    );
  }
  return rgb;
}

export function colorAttrs(color: Color): string {
  const parts: string[] = [];
  if (color.argb !== undefined) parts.push(`rgb="${normalizeArgb(color.argb)}"`);
  if (color.theme !== undefined) parts.push(`theme="${color.theme}"`);
  if (color.tint !== undefined) parts.push(`tint="${color.tint}"`);
  if (color.indexed !== undefined) parts.push(`indexed="${color.indexed}"`);
  return parts.join(' ');
}
