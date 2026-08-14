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

import {
  type Alignment,
  type Border,
  type BorderEdge,
  type Color,
  type DifferentialStyle,
  type Fill,
  type Font,
  type GradientFill,
  isBorderStyle,
  isFillPatternType,
  isFontScheme,
  isFontVerticalAlignment,
  isHorizontalAlignment,
  isNamedUnderlineStyle,
  isVerticalAlignment,
  type NamedCellStyle,
  type Protection,
  type TableStyleTable,
  type UnderlineStyle,
} from '../../core/style.ts';
import {TABLE_STYLE_ELEMENT_TYPES, type TableStyle} from '../../core/table-style.ts';
import {AuthoringError} from '../../errors.ts';
import {escapeAttr, XML_DECLARATION} from '../../xml/xml.ts';
import {decodeEntities} from '../../xml/xml-read.ts';
import {colorAttrs} from './color-xml.ts';
import {MARKUP_COMPATIBILITY_NS, SPREADSHEETML_NS} from './namespaces.ts';

// Excel reserves fill ids 0 and 1 for the "none" and "gray125" patterns it always emits;
// custom fills are numbered from 2 so a foreign reader's built-in assumptions still hold.
const RESERVED_FILL_COUNT = 2;

// Font id 0 is the always-present default font — the workbook's, not an assumed one; custom fonts
// are numbered from 1.
const RESERVED_FONT_COUNT = 1;

// numFmt ids below 164 are reserved by ECMA-376 for the built-in formats every consumer
// knows implicitly; custom format codes are numbered from 164 up. Id 0 is General (no code).
const CUSTOM_NUMFMT_BASE = 164;

// Border id 0 is the always-present empty border (every edge absent); custom borders from 1.
const RESERVED_BORDER_COUNT = 1;

// The Office default font's inner fragment, in the exact child order `fontXml` emits — the font 0 a
// registry built without a workbook falls back to. A registry built *with* one derives font 0 from
// its `defaultFont` instead, and for a plain `new Workbook()` that derivation lands on exactly this
// string; `styles.test.ts` guards the two against drifting.
const DEFAULT_FONT_BODY =
  '<sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/>';
// The empty border: all five edges present but styleless. A border that overrides no edge
// serialises to exactly this, so it interns to the default border id 0 rather than a new one.
const DEFAULT_BORDER = '<border><left/><right/><top/><bottom/><diagonal/></border>';

/** A cell's style facets as the writer composes them: cell overrides atop row/column defaults. */
export interface CellStyle {
  readonly fill?: Fill | undefined;
  readonly numFmt?: string | undefined;
  readonly font?: Font | undefined;
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

/** What a {@link StyleRegistry} needs from its workbook before any style is interned. */
export interface StyleRegistryOptions {
  /**
   * The font every cell naming none of its own renders in — emitted as id 0. Take it from
   * {@link Workbook.defaultFont}, which resolves the authored/declared/theme chain and guarantees a
   * complete entry. Omitted, the registry falls back to the Office default.
   */
  readonly defaultFont?: Font;

  /**
   * The default font the *source package* declared, when a workbook was read
   * ({@link Workbook.declaredDefaultFont}). It also interns to id 0, which is what lets an authored
   * default font reach cells the reader had already resolved.
   *
   * The reader flattens font 0 onto every xf that names it, so a cell that merely inherited the
   * file's default arrives carrying it as a concrete face. That face is an artefact of reading, not
   * an authored intent — in the source file the cell said nothing about its font — so it must
   * collapse back to id 0 and follow the new default rather than pin itself to the old one through a
   * custom entry.
   */
  readonly declaredDefaultFont?: Font;
}

export class StyleRegistry {
  // The `<font>` body emitted as id 0.
  readonly #defaultFontBody: string;

  // Every serialised font body that means "id 0" — the emitted default, plus the one a source file
  // declared. See {@link StyleRegistryOptions.declaredDefaultFont} for why the second belongs here.
  readonly #font0Bodies: ReadonlySet<string>;

  constructor(options: StyleRegistryOptions = {}) {
    this.#defaultFontBody =
      options.defaultFont === undefined ? DEFAULT_FONT_BODY : fontXml(options.defaultFont);
    this.#font0Bodies = new Set(
      options.declaredDefaultFont === undefined
        ? [this.#defaultFontBody]
        : [this.#defaultFontBody, fontXml(options.declaredDefaultFont)],
    );
  }

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

  // The most-recently-used colour swatches (`<colors><mruColors>`) read from a file, each entry a
  // verbatim `<color rgb="…"/>` fragment. Purely a UI convenience — the palette Excel offers under
  // "Recent Colors" — but it is the author's own working set, so dropping it on a re-write quietly
  // resets a habit. Empty for a workbook that never picked a custom colour.
  readonly #mruColors: string[] = [];

  // The custom table/pivot style definitions (`<tableStyles>`) read from a file, kept verbatim, plus
  // the gallery names it nominates as defaults. A table's `tableStyleInfo/@name` can point at one of
  // these definitions, so dropping the block leaves that reference dangling and the table renders
  // unstyled. Excel writes the container (with both default attributes) into essentially every file
  // even when it declares no custom style at all.
  #tableStyles: TableStyleTable = {styles: []};

  // Table styles authored on the workbook, serialised on registration and keyed by name so a second
  // definition of the same name replaces the first — as does one that overrides a preserved
  // definition, since two `<tableStyle>` elements sharing a name leave a table's reference ambiguous.
  readonly #authoredTableStyles = new Map<string, string>();

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
   * verbatim and at its original index. Call once before any {@link differentialStyleId}; authored
   * styles append after these.
   *
   * **Index stability is a contract, not an implementation detail.** A `dxfId` is an index into this
   * one table, and more than one preserved construct resolves through it: a conditional-formatting
   * rule's `dxfId`, and every `<tableStyleElement dxfId="…">` inside a preserved `<tableStyle>` (see
   * {@link seedTableStyles}). Those constructs are carried as opaque XML precisely *because* the
   * indices they name do not move. Renumbering, reordering, or de-duplicating the seeded entries
   * would silently re-point every one of them at a different format — a change no schema check and no
   * round-trip of our own can catch, because the file stays perfectly valid and merely renders wrong.
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

  /**
   * Seed the most-recently-used colour swatches (`<colors><mruColors>`) read from a file, each entry a
   * verbatim `<color rgb="…"/>` fragment. An empty list emits no `<mruColors>` element.
   */
  seedMruColors(fragments: readonly string[]): void {
    this.#mruColors.length = 0;
    this.#mruColors.push(...fragments);
  }

  /**
   * Seed the custom table-style definitions (`<tableStyles>`) read from a file, each `<tableStyle>`
   * kept verbatim so a table's `tableStyleInfo/@name` still resolves and each element's `dxfId` still
   * indexes the differential-style table {@link seedDifferentialStyles} preserves at its original
   * indices. Replaces any block already held.
   */
  seedTableStyles(table: TableStyleTable): void {
    this.#tableStyles = table;
  }

  /**
   * Serialise a table style authored on the workbook, interning each element's formatting into the
   * differential-style table and emitting the `dxfId` that reaches it. Call after
   * {@link seedTableStyles}, whose preserved definitions these append after.
   *
   * A definition here **replaces** a preserved one of the same name. Two `<tableStyle>` elements
   * sharing a name is ambiguous — a table's `tableStyleInfo/@name` would reach whichever a consumer
   * happened to index first — so authoring a name the source already used is read as overriding it,
   * which is what asking for it means.
   */
  addTableStyle(style: TableStyle): void {
    const elements = TABLE_STYLE_ELEMENT_TYPES.flatMap((type) => {
      const element = style.elements[type];
      if (element === undefined) return [];
      // `size` defaults to 1, so it is written only when a band is genuinely wider than one row.
      const size =
        element.size !== undefined && element.size !== 1 ? ` size="${element.size}"` : '';
      return [
        `<tableStyleElement type="${type}"${size} dxfId="${this.differentialStyleId(element)}"/>`,
      ];
    });
    // `pivot`/`table` default to true, so each is written only when the caller opts a style out.
    const flags =
      (style.pivot === false ? ' pivot="0"' : '') + (style.table === false ? ' table="0"' : '');
    this.#authoredTableStyles.set(
      style.name,
      `<tableStyle name="${escapeAttr(style.name)}"${flags} count="${elements.length}">` +
        `${elements.join('')}</tableStyle>`,
    );
  }

  /**
   * Intern an authored differential style, returning the `<dxfs>` index that references it — a
   * conditional-formatting rule's `dxfId`, or a table style element's. Identical styles collapse to
   * one entry, whichever feature asked for them, so a highlight rule and a table style's header row
   * painted the same way share a single `<dxf>`.
   *
   * Authored entries append after the seeded ones ({@link seedDifferentialStyles}), which is what
   * keeps every preserved reference pointing where it did.
   */
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
      this.#fillXml.push(patternFillXml(fill, {solidBgFallback: true}));
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
  #internFont(font: Font): number {
    const xml = fontXml(font);
    if (xml === '' || this.#font0Bodies.has(xml)) return 0;
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
    const fonts = `<font>${this.#defaultFontBody}</font>${this.#fontXml.join('')}`;
    const borderCount = RESERVED_BORDER_COUNT + this.#borderXml.length;
    const borders = DEFAULT_BORDER + this.#borderXml.join('');
    return (
      XML_DECLARATION +
      `<styleSheet xmlns="${SPREADSHEETML_NS}"${this.#foreignNamespaceAttrs()}>` +
      this.#numFmtsXml() +
      `<fonts count="${fontCount}">${fonts}</fonts>` +
      `<fills count="${fillCount}">${fills}</fills>` +
      `<borders count="${borderCount}">${borders}</borders>` +
      `<cellStyleXfs count="${this.#cellStyleXfs.length}">${cellStyleXfs}</cellStyleXfs>` +
      `<cellXfs count="${this.#formats.length}">${cellXfs}</cellXfs>` +
      `<cellStyles count="${this.#cellStyleNames.length}">${cellStyles}</cellStyles>` +
      this.#dxfsXml() +
      this.#tableStylesXml() +
      this.#colorsXml() +
      '</styleSheet>'
    );
  }

  // The `xmlns:…` declarations a preserved `<tableStyle>` fragment depends on, plus the
  // markup-compatibility attributes that tell a consumer to ignore what it does not understand. A
  // workbook carrying no such fragment emits nothing, so the ordinary stylesheet root is unchanged.
  //
  // This is the cost of verbatim preservation: a fragment carries its prefixes with it, and a prefix
  // no ancestor declares makes the whole part unparseable — a much louder failure than the dropped
  // table style the preservation exists to prevent. See {@link TableStyleTable.namespaces}.
  #foreignNamespaceAttrs(): string {
    const namespaces = this.#tableStyles.namespaces ?? [];
    if (namespaces.length === 0) return '';
    const declarations = namespaces.map((ns) => ` xmlns:${ns.prefix}="${escapeAttr(ns.uri)}"`);
    const ignorable = namespaces.filter((ns) => ns.ignorable).map((ns) => ns.prefix);
    const compatibility =
      ignorable.length === 0
        ? ''
        : ` xmlns:mc="${MARKUP_COMPATIBILITY_NS}" mc:Ignorable="${escapeAttr(ignorable.join(' '))}"`;
    return declarations.join('') + compatibility;
  }

  // <tableStyles> sits between <dxfs> and <colors> in CT_Stylesheet's child sequence. It is emitted
  // only when the workbook has something to say there — a preserved or authored style definition, or
  // a nominated default — so a workbook that authors none leaves every table on the built-in gallery
  // and writes nothing. `count` counts the definitions, not the attributes, so a container that only
  // nominates defaults (the shape Excel writes into nearly every file) is self-closing with count="0".
  //
  // Preserved definitions come first and authored ones after, except that an authored style replaces
  // the preserved definition it shares a name with — see {@link addTableStyle}.
  #tableStylesXml(): string {
    const {defaultTableStyle, defaultPivotStyle} = this.#tableStyles;
    const authored = this.#authoredTableStyles;
    const preserved = this.#tableStyles.styles.filter(
      (fragment) => !authored.has(tableStyleName(fragment)),
    );
    const styles = [...preserved, ...authored.values()];
    if (styles.length === 0 && defaultTableStyle === undefined && defaultPivotStyle === undefined) {
      return '';
    }
    const attrs =
      ` count="${styles.length}"` +
      (defaultTableStyle === undefined
        ? ''
        : ` defaultTableStyle="${escapeAttr(defaultTableStyle)}"`) +
      (defaultPivotStyle === undefined
        ? ''
        : ` defaultPivotStyle="${escapeAttr(defaultPivotStyle)}"`);
    if (styles.length === 0) return `<tableStyles${attrs}/>`;
    return `<tableStyles${attrs}>${styles.join('')}</tableStyles>`;
  }

  // <colors> is the last modelled child of <styleSheet>, after <dxfs> and <tableStyles>. It holds the
  // custom <indexedColors> palette and then the <mruColors> swatch list, in that CT_Colors order. It
  // is emitted only when a file carried one of them, so an ordinary workbook stays on the built-in
  // indexed colours and writes no <colors> element.
  #colorsXml(): string {
    const indexed =
      this.#indexedColors.length === 0
        ? ''
        : `<indexedColors>${this.#indexedColors.join('')}</indexedColors>`;
    const mru =
      this.#mruColors.length === 0 ? '' : `<mruColors>${this.#mruColors.join('')}</mruColors>`;
    if (indexed === '' && mru === '') return '';
    return `<colors>${indexed}${mru}</colors>`;
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

// Reject an enum-typed style token the writer would otherwise emit verbatim. The public types already
// forbid an out-of-contract value (VerticalAlignment, BorderStyle, FillPatternType, …), so this fires
// only for a value smuggled past the types by an untyped caller — but the writer must never serialise
// it: it would be schema-invalid OOXML that Excel silently tolerates yet the library's own reader
// (which narrows every such token through the same guard) discards on read-back. Rejecting at the write
// boundary keeps the writer symmetric with the reader — garbage out refused exactly as garbage in — so
// a value the writer accepts is always one that round-trips.
function checkedToken(
  value: string,
  isValid: (candidate: string) => boolean,
  kind: string,
): string {
  if (!isValid(value)) {
    throw new AuthoringError(
      `Invalid ${kind} ${JSON.stringify(value)}: not a value the OOXML enumeration allows`,
    );
  }
  return value;
}

// Serialise a cell's alignment as `<alignment>` attributes in ECMA-376 CT_CellAlignment order.
// A facet at its default contributes nothing; an all-default alignment yields the empty string,
// so it forces neither an <alignment> child nor a distinct xf.
function alignmentAttrs(alignment: Alignment): string {
  const parts: string[] = [];
  // `general` is the type-dependent default and is expressed by omitting the attribute.
  if (alignment.horizontal !== undefined && alignment.horizontal !== 'general') {
    parts.push(
      `horizontal="${checkedToken(alignment.horizontal, isHorizontalAlignment, 'horizontal alignment')}"`,
    );
  }
  if (alignment.vertical !== undefined) {
    parts.push(
      `vertical="${checkedToken(alignment.vertical, isVerticalAlignment, 'vertical alignment')}"`,
    );
  }
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

// The `name` a `<tableStyle>` fragment declares — the key a table's `tableStyleInfo/@name` matches
// and, here, the key an authored definition overrides a preserved one by. Read out of the fragment
// rather than stored beside it, so the two cannot drift; `name` is required by CT_TableStyle, and a
// fragment without one is unreachable anyway and so can never collide.
function tableStyleName(fragment: string): string {
  return decodeEntities(/<tableStyle\b[^>]*\bname="([^"]*)"/.exec(fragment)?.[1] ?? '');
}

// Serialise the facets a font overrides, in ECMA-376 child order. A boolean flag is emitted only
// when true (its absence is the default false); an empty result means the font differs from the
// default in nothing and needs no entry at all. The face element differs by context: a styles
// `<font>` names it `<name>` (CT_Font), a rich-text run's `<rPr>` names it `<rFont>` (CT_RPrElt) —
// otherwise the two share every child, so `nameTag` selects the face element and the rest is common.
export function fontXml(font: Font, nameTag: 'name' | 'rFont' = 'name'): string {
  const parts: string[] = [];
  if (font.bold) parts.push('<b/>');
  if (font.italic) parts.push('<i/>');
  if (font.strike) parts.push('<strike/>');
  if (font.outline) parts.push('<outline/>');
  const underline = underlineXml(font.underline);
  if (underline !== '') parts.push(underline);
  if (font.vertAlign !== undefined) {
    parts.push(
      `<vertAlign val="${checkedToken(font.vertAlign, isFontVerticalAlignment, 'font vertical alignment')}"/>`,
    );
  }
  if (font.size !== undefined) parts.push(`<sz val="${numberAttr(font.size)}"/>`);
  if (font.color !== undefined) parts.push(`<color ${colorAttrs(font.color)}/>`);
  if (font.name !== undefined) parts.push(`<${nameTag} val="${escapeAttr(font.name)}"/>`);
  if (font.family !== undefined) parts.push(`<family val="${numberAttr(font.family)}"/>`);
  if (font.charset !== undefined) parts.push(`<charset val="${numberAttr(font.charset)}"/>`);
  if (font.scheme !== undefined && font.scheme !== 'none')
    parts.push(`<scheme val="${checkedToken(font.scheme, isFontScheme, 'font scheme')}"/>`);
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
  if (style.fill !== undefined) parts.push(patternFillXml(style.fill, {solidBgFallback: false}));
  if (style.border !== undefined) parts.push(borderXml(style.border));
  return `<dxf>${parts.join('')}</dxf>`;
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
  return `<u val="${checkedToken(underline, isNamedUnderlineStyle, 'underline style')}"/>`;
}

function numberAttr(value: number): string {
  if (!Number.isFinite(value)) {
    throw new AuthoringError(`cannot serialise a non-finite font metric (${value})`);
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
  const style = checkedToken(edge.style, isBorderStyle, 'border style');
  if (edge.color === undefined) return `<${tag} style="${style}"/>`;
  return `<${tag} style="${style}"><color ${colorAttrs(edge.color)}/></${tag}>`;
}

// A format code sits in the `formatCode` attribute; only the markup-significant characters
// need escaping. A code can legitimately contain `"` (quoted literals like `"$"`), `<`, `&`.
// Unlike `escapeAttr`, a lone `'` is left untouched: it is not markup-significant inside a
// double-quoted attribute, and Excel writes format codes with bare apostrophes, so leaving it
// keeps the round-tripped code byte-identical to the source.
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

// The `<fill>` element for a pattern or gradient fill. The two callers differ only in the solid-fill
// background fallback: a cell fill forces the automatic indexed placeholder onto a solid pattern that
// names no background — omitting it makes Excel render the fill as flat black — whereas a dxf states
// only the overrides it carries, so `solidBgFallback` gates that placeholder.
function patternFillXml(fill: Fill, {solidBgFallback}: {solidBgFallback: boolean}): string {
  if (fill.type === 'gradient') return `<fill>${gradientFillXml(fill)}</fill>`;
  const fg = fill.fgColor ? `<fgColor ${colorAttrs(fill.fgColor)}/>` : '';
  const bg = fill.bgColor
    ? `<bgColor ${colorAttrs(fill.bgColor)}/>`
    : solidBgFallback && fill.pattern === 'solid'
      ? '<bgColor indexed="64"/>'
      : '';
  const pattern = checkedToken(fill.pattern, isFillPatternType, 'fill pattern');
  return `<fill><patternFill patternType="${pattern}">${fg}${bg}</patternFill></fill>`;
}
