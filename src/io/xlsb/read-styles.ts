// The BIFF12 style-sheet reader: `xl/styles.bin` in, the same {@link StyleTable} the XML reader
// produces out. Every collection (`fmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`,
// `styles`) is a Begin/End record pair around its entries, so the pass tracks which collection it is
// inside — `BrtXF` appears in two of them and is meaningless without that context.
//
// The resolution *rules* are deliberately not restated here: number-format ids go through the XML
// reader's `numFmtCodeFor`, and each facet is emitted with the same "only when it differs from the
// default" discipline the XML path uses. That is what makes the binary and XML readings of one
// workbook the same model rather than two similar ones — a bottom vertical alignment, a locked cell,
// or a General number format is written explicitly in BIFF12 and omitted in XML, so the binary side
// has to drop exactly what the XML side never had.

import type {
  Alignment,
  Border,
  BorderEdge,
  BorderStyle,
  Color,
  Fill,
  FillPatternType,
  Font,
  HorizontalAlignment,
  NamedCellStyle,
  Protection,
  UnderlineStyle,
  VerticalAlignment,
} from '../../core/style.ts';
import {assignStyleFacets} from '../../core/style.ts';
import {numFmtCodeFor, type StyleTable, type XfStyle} from '../style/xf-style.ts';
import {RecordReader} from './primitives.ts';
import {readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';

// A `<cellStyle>`'s BIFF12 counterpart: the name/builtinId labelling one cellStyleXfs entry.
interface StyleLabel {
  readonly xfId: number;
  readonly name?: string;
  readonly builtinId?: number;
}

// Which Begin/End-delimited collection the pass is currently inside. `undefined` outside all of them,
// which is also what an unrecognised nested block collapses to — so a record we do not model can
// never be mistaken for an entry of the collection that happened to precede it.
type Collection =
  | 'fmts'
  | 'fonts'
  | 'fills'
  | 'borders'
  | 'cellStyleXfs'
  | 'cellXfs'
  | 'styles'
  | undefined;

const COLLECTION_STARTS: ReadonlyMap<number, Collection> = new Map<number, Collection>([
  [BRT.BeginFmts, 'fmts'],
  [BRT.BeginFonts, 'fonts'],
  [BRT.BeginFills, 'fills'],
  [BRT.BeginBorders, 'borders'],
  [BRT.BeginCellStyleXFs, 'cellStyleXfs'],
  [BRT.BeginCellXFs, 'cellXfs'],
  [BRT.BeginStyles, 'styles'],
]);

const COLLECTION_ENDS: ReadonlySet<number> = new Set([
  BRT.EndFmts,
  BRT.EndFonts,
  BRT.EndFills,
  BRT.EndBorders,
  BRT.EndCellStyleXFs,
  BRT.EndCellXFs,
  BRT.EndStyles,
]);

/** Parse `xl/styles.bin` into the flat cell-format table a worksheet's style indices resolve against. */
export function parseStyleTable(part: Uint8Array | undefined): StyleTable {
  if (part === undefined) return {cellXfs: [], namedStyles: []};

  const numFmtCodes = new Map<number, string>();
  const fonts: Array<Font | undefined> = [];
  const fills: Array<Fill | undefined> = [];
  const borders: Array<Border | undefined> = [];
  const namedXfs: XfStyle[] = [];
  const directXfs: XfStyle[] = [];
  const labels: StyleLabel[] = [];
  let collection: Collection;

  for (const record of readRecords(part)) {
    if (COLLECTION_ENDS.has(record.type)) {
      collection = undefined;
      continue;
    }
    const started = COLLECTION_STARTS.get(record.type);
    if (started !== undefined) {
      collection = started;
      continue;
    }
    const reader = new RecordReader(record.data);
    switch (record.type) {
      case BRT.Fmt:
        if (collection === 'fmts') numFmtCodes.set(reader.u16(), reader.wideString());
        break;
      case BRT.Font:
        if (collection === 'fonts') fonts.push(readFont(reader));
        break;
      case BRT.Fill:
        if (collection === 'fills') fills.push(readFill(reader));
        break;
      case BRT.Border:
        if (collection === 'borders') borders.push(readBorder(reader));
        break;
      case BRT.XF:
        if (collection === 'cellXfs' || collection === 'cellStyleXfs') {
          const deps = {fonts, fills, borders, numFmtCodes};
          (collection === 'cellXfs' ? directXfs : namedXfs).push(
            readXf(reader, deps, collection === 'cellXfs'),
          );
        }
        break;
      case BRT.Style:
        if (collection === 'styles') labels.push(readStyleLabel(reader));
        break;
      default:
        break;
    }
  }

  // Layer each direct format over the named style its xfId links to, exactly as the XML reader does:
  // a facet the cell's own xf sets wins, one it leaves unset falls through to the named base.
  const cellXfs = directXfs.map((xf) => {
    if (xf.xfId === undefined) return xf;
    const named = namedXfs[xf.xfId];
    return named === undefined ? xf : {...named, ...xf};
  });

  const namedStyles: NamedCellStyle[] = namedXfs.map((xf, index) => {
    const label = labels.find((entry) => entry.xfId === index);
    const style: {-readonly [K in keyof NamedCellStyle]?: NamedCellStyle[K]} = {};
    assignStyleFacets(style, xf);
    if (label?.name !== undefined) style.name = label.name;
    if (label?.builtinId !== undefined) style.builtinId = label.builtinId;
    return style;
  });

  return {cellXfs, namedStyles};
}

// The shared sub-tables an XF resolves its facet indices against.
interface XfDeps {
  readonly fonts: ReadonlyArray<Font | undefined>;
  readonly fills: ReadonlyArray<Fill | undefined>;
  readonly borders: ReadonlyArray<Border | undefined>;
  readonly numFmtCodes: ReadonlyMap<number, string>;
}

// `BrtXF` ([MS-XLSB] 2.4.876): five facet indices, the two inline alignment scalars, then two flag
// words. `isDirect` distinguishes a cellXfs entry (whose ixfeParent links to a named style) from a
// cellStyleXfs entry (whose ixfeParent is the 0xFFFF "I am the named style" marker).
function readXf(reader: RecordReader, deps: XfDeps, isDirect: boolean): XfStyle {
  const parent = reader.u16();
  const numFmt = numFmtCodeFor(reader.u16(), deps.numFmtCodes);
  const font = deps.fonts[reader.u16()];
  const fill = deps.fills[reader.u16()];
  const borderIndex = reader.u16();
  // Border 0 is the genuinely empty default (font 0, by contrast, is the workbook's real default
  // face), so only a custom index names an actual border — the same asymmetry the XML reader keeps.
  const border = borderIndex > 0 ? deps.borders[borderIndex] : undefined;
  const rotation = reader.u8();
  const indent = reader.u8();
  const flags = reader.u16();

  const draft: {-readonly [K in keyof XfStyle]?: XfStyle[K]} = {};
  if (numFmt !== undefined) draft.numFmt = numFmt;
  if (font) draft.font = font;
  if (fill) draft.fill = fill;
  if (border) draft.border = border;
  const alignment = readAlignment(flags, rotation, indent);
  if (alignment !== undefined) draft.alignment = alignment;
  const protection = readProtection(flags);
  if (protection !== undefined) draft.protection = protection;
  // f123Prefix is the binary spelling of the `quotePrefix` attribute: the cell's text is displayed
  // with a leading apostrophe stripped.
  if ((flags & 0x8000) !== 0) draft.quotePrefix = true;
  // Only a link beyond the Normal default (0) carries information, and only a direct format has one.
  if (isDirect && parent !== NOT_A_CELL_XF && parent > 0) draft.xfId = parent;
  return draft;
}

// The `ixfeParent` value a cell *style* XF carries in place of a link, since it is itself the base.
const NOT_A_CELL_XF = 0xffff;

// BIFF12 states every alignment field on every xf, where XML omits the ones at their default. Keep
// only what the XML reader would have seen, so the two readings of one workbook agree: `general`
// horizontal, `bottom` vertical, and zero rotation/indent/reading-order are absences, not values.
function readAlignment(flags: number, rotation: number, indent: number): Alignment | undefined {
  const out: {-readonly [K in keyof Alignment]?: Alignment[K]} = {};
  const horizontal = HORIZONTAL_ALIGNMENTS[flags & 0b111];
  if (horizontal !== undefined) out.horizontal = horizontal;
  const vertical = VERTICAL_ALIGNMENTS[(flags >> 3) & 0b111];
  if (vertical !== undefined) out.vertical = vertical;
  if (rotation !== 0) out.textRotation = rotation;
  if ((flags & 0x0040) !== 0) out.wrapText = true;
  if (indent !== 0) out.indent = indent;
  if ((flags & 0x0100) !== 0) out.shrinkToFit = true;
  const readingOrder = (flags >> 10) & 0b11;
  if (readingOrder !== 0) out.readingOrder = readingOrder;
  return Object.keys(out).length > 0 ? out : undefined;
}

// `locked` defaults to TRUE in OOXML, so an *unlocked* cell is the state that carries information;
// `hidden` defaults to false, so only a set flag does. A default xf yields no protection at all.
function readProtection(flags: number): Protection | undefined {
  const out: {-readonly [K in keyof Protection]?: Protection[K]} = {};
  if ((flags & 0x1000) === 0) out.locked = false;
  if ((flags & 0x2000) !== 0) out.hidden = true;
  return Object.keys(out).length > 0 ? out : undefined;
}

// `alc` ([MS-XLSB] 2.4.876), indexed by its stored value. `general` is index 0 and is left out
// deliberately: it is the type-dependent default, which the model spells as no horizontal alignment.
const HORIZONTAL_ALIGNMENTS: ReadonlyArray<HorizontalAlignment | undefined> = [
  undefined,
  'left',
  'center',
  'right',
  'fill',
  'justify',
  'centerContinuous',
  'distributed',
];

// `alcv`, likewise — with `bottom` (index 2) left out as the default.
const VERTICAL_ALIGNMENTS: ReadonlyArray<VerticalAlignment | undefined> = [
  'top',
  'center',
  undefined,
  'justify',
  'distributed',
];

// `BrtFont` ([MS-XLSB] 2.4.690). Weight is a numeric scale (400 normal, 700 bold) rather than a flag,
// and the italic/strike/outline bits live in a separate word from it.
function readFont(reader: RecordReader): Font {
  const height = reader.u16();
  const flags = reader.u16();
  const weight = reader.u16();
  const script = reader.u16();
  const underline = reader.u8();
  const family = reader.u8();
  const charset = reader.u8();
  reader.skip(1); // unused
  const color = reader.color();
  const scheme = reader.u8();
  const name = reader.wideString();

  const font: {-readonly [K in keyof Font]?: Font[K]} = {};
  if (name !== '') font.name = name;
  // Stored in twips; the model (like the XML) carries points.
  if (height > 0) font.size = height / 20;
  if (family !== 0) font.family = family;
  if (charset !== 0) font.charset = charset;
  if (scheme === 1) font.scheme = 'major';
  else if (scheme === 2) font.scheme = 'minor';
  if (color !== undefined) font.color = color;
  // Each boolean facet is recorded only when on, mirroring XML's present-or-absent `<b/>`/`<i/>`:
  // a non-bold font must not read back as `bold: false`.
  if (weight >= BOLD_WEIGHT) font.bold = true;
  if ((flags & 0b0000_0010) !== 0) font.italic = true;
  if ((flags & 0b0000_1000) !== 0) font.strike = true;
  if ((flags & 0b0001_0000) !== 0) font.outline = true;
  const underlineStyle = UNDERLINE_STYLES.get(underline);
  if (underlineStyle !== undefined) font.underline = underlineStyle;
  if (script === 1) font.vertAlign = 'superscript';
  else if (script === 2) font.vertAlign = 'subscript';
  return font;
}

const BOLD_WEIGHT = 700;

// `uls` ([MS-XLSB] 2.4.690). A single underline is `true`, not `'single'`, because that is what XML's
// bare `<u/>` — which is what Excel writes for it — reads back as.
const UNDERLINE_STYLES: ReadonlyMap<number, UnderlineStyle> = new Map<number, UnderlineStyle>([
  [0x01, true],
  [0x02, 'double'],
  [0x21, 'singleAccounting'],
  [0x22, 'doubleAccounting'],
]);

// `BrtFill` ([MS-XLSB] 2.4.681). The pattern code and OOXML's `ST_PatternType` enumerate the same
// patterns in the same order, so the code indexes the name list directly.
function readFill(reader: RecordReader): Fill | undefined {
  const pattern = FILL_PATTERNS[reader.u32()];
  // `none` is the absence of a fill, and an unmodelled pattern (a gradient — see below) is dropped
  // rather than guessed, so an unfilled cell reads back unfilled either way.
  if (pattern === undefined) return undefined;
  // BIFF12 always states both colours; XML states only the ones the fill actually has, using the two
  // legacy-palette sentinels for the rest — 64 is "automatic foreground", 65 "automatic background".
  // Dropping each in its own slot reproduces exactly what the XML reader sees: an untouched hatch
  // pattern carries no colours at all, while a solid fill keeps the explicit `bgColor indexed="64"`
  // Excel writes beside its foreground.
  const fgColor = notSentinel(reader.color(), AUTOMATIC_FOREGROUND);
  const bgColor = notSentinel(reader.color(), AUTOMATIC_BACKGROUND);
  return {
    type: 'pattern',
    pattern,
    ...(fgColor ? {fgColor} : {}),
    ...(bgColor ? {bgColor} : {}),
  };
}

const AUTOMATIC_FOREGROUND = 64;
const AUTOMATIC_BACKGROUND = 65;

function notSentinel(color: Color | undefined, sentinel: number): Color | undefined {
  return color?.indexed === sentinel ? undefined : color;
}

// Indexed by the stored `fls` value. Index 0 (`none`) is deliberately absent: an unfilled cell
// carries no fill. Gradient fills (`fls` 0x28) are not decoded in this cut — the stop array's layout
// is the one piece of BrtFill this reader has no Excel-authored sample to check against, and a
// silently wrong gradient is worse than none.
const FILL_PATTERNS: ReadonlyArray<FillPatternType | undefined> = [
  undefined,
  'solid',
  'mediumGray',
  'darkGray',
  'lightGray',
  'darkHorizontal',
  'darkVertical',
  'darkDown',
  'darkUp',
  'darkGrid',
  'darkTrellis',
  'lightHorizontal',
  'lightVertical',
  'lightDown',
  'lightUp',
  'lightGrid',
  'lightTrellis',
  'gray125',
  'gray0625',
];

// `BrtBorder` ([MS-XLSB] 2.4.314): the two diagonal-direction bits, then five `Blxf` edges in the
// order top, bottom, left, right, diagonal — which is *not* the model's or the schema's order, so
// the edges are read positionally and named here.
function readBorder(reader: RecordReader): Border | undefined {
  const flags = reader.u8();
  const top = readEdge(reader);
  const bottom = readEdge(reader);
  const left = readEdge(reader);
  const right = readEdge(reader);
  const diagonal = readEdge(reader);

  const border: {-readonly [K in keyof Border]?: Border[K]} = {};
  if (left) border.left = left;
  if (right) border.right = right;
  if (top) border.top = top;
  if (bottom) border.bottom = bottom;
  if (diagonal) border.diagonal = diagonal;
  if ((flags & 0b01) !== 0) border.diagonalDown = true;
  if ((flags & 0b10) !== 0) border.diagonalUp = true;
  // An all-default border is the empty one every unbordered cell shares; it carries nothing.
  return Object.keys(border).length > 0 ? border : undefined;
}

// A `Blxf` ([MS-XLSB] 2.5.5): a line style, a reserved byte, and a colour. Style 0 is "no edge",
// which the model spells by omitting the edge rather than by a `none` value.
function readEdge(reader: RecordReader): BorderEdge | undefined {
  const style = BORDER_STYLES[reader.u8()];
  reader.skip(1); // reserved
  const color = reader.color();
  if (style === undefined) return undefined;
  return color === undefined ? {style} : {style, color};
}

// Indexed by the stored `dg` value; index 0 (`none`) is absent, as above. The order is the binary
// format's own and differs from `ST_BorderStyle`'s declaration order, so it cannot be shared.
const BORDER_STYLES: ReadonlyArray<BorderStyle | undefined> = [
  undefined,
  'thin',
  'medium',
  'dashed',
  'dotted',
  'thick',
  'double',
  'hair',
  'mediumDashed',
  'dashDot',
  'mediumDashDot',
  'dashDotDot',
  'mediumDashDotDot',
  'slantDashDot',
];

// `BrtStyle` ([MS-XLSB] 2.4.809): which cellStyleXfs entry this names, and how it is labelled. The
// gallery index is only meaningful for a built-in style, which the flag word declares.
function readStyleLabel(reader: RecordReader): StyleLabel {
  const xfId = reader.u32();
  const flags = reader.u16();
  const builtinId = reader.u8();
  reader.skip(1); // iLevel: the outline depth of a built-in RowLevel/ColLevel style.
  const name = reader.wideString();
  return {
    xfId,
    ...(name !== '' ? {name} : {}),
    ...((flags & 0b1) !== 0 ? {builtinId} : {}),
  };
}

// Re-exported so a caller reading an `.xlsb` never needs to reach into the XML cluster for the type
// its style table is expressed in.
export type {StyleTable, XfStyle};
