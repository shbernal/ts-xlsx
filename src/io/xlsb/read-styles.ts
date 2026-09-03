// The BIFF12 style-sheet reader: `xl/styles.bin` in, the same {@link StyleTable} the XML reader
// produces out. Every collection (`fmts`, `fonts`, `fills`, `borders`, `cellStyleXfs`, `cellXfs`,
// `styles`) is a Begin/End record pair around its entries, so the pass tracks which collection it is
// inside: `BrtXF` appears in two of them and is meaningless without that context.
//
// The resolution *rules* are deliberately not restated here: number-format ids go through
// `numFmtCodeFor` in `../style/xf-style.ts`, the shared layer both codecs sit above, and each facet is
// emitted with the same "only when it differs from the default" discipline the XML path uses. That is what makes the binary and XML readings of one
// workbook the same model rather than two similar ones: a bottom vertical alignment, a locked cell,
// or a General number format is written explicitly in BIFF12 and omitted in XML, so the binary side
// has to drop exactly what the XML side never had.

import {
  type Alignment,
  ALIGNMENT_FACETS,
  type Border,
  type BorderEdge,
  type BorderStyle,
  type Color,
  type Fill,
  FILL_PATTERNS_IN_SCHEMA_ORDER,
  type Font,
  type HorizontalAlignment,
  type Protection,
  type UnderlineStyle,
  type VerticalAlignment,
} from '../../core/style.ts';
import {
  numFmtCodeFor,
  NO_PRESERVED_STYLE_TABLES,
  protectionFrom,
  resolveStyleTable,
  type StyleLabel,
  type StyleTable,
  type XfDeps,
  type XfStyle,
} from '../style/xf-style.ts';
import {RecordReader} from './primitives.ts';
import {blockTracker, readRecords} from './record-stream.ts';
import {BRT} from './record-types.ts';

// The Begin/End-delimited collections this pass reads. A record outside all of them belongs to none,
// so one we do not model can never be mistaken for an entry of the collection that happened to
// precede it.
type Collection = 'fmts' | 'fonts' | 'fills' | 'borders' | 'cellStyleXfs' | 'cellXfs' | 'styles';

// Each end names the collection it closes on the same line as its start, so a stray `EndFonts` cannot
// silently close `<fills>` and neither can a mis-typed table entry.
const COLLECTIONS: readonly (readonly [number, number, Collection])[] = [
  [BRT.BeginFmts, BRT.EndFmts, 'fmts'],
  [BRT.BeginFonts, BRT.EndFonts, 'fonts'],
  [BRT.BeginFills, BRT.EndFills, 'fills'],
  [BRT.BeginBorders, BRT.EndBorders, 'borders'],
  [BRT.BeginCellStyleXFs, BRT.EndCellStyleXFs, 'cellStyleXfs'],
  [BRT.BeginCellXFs, BRT.EndCellXFs, 'cellXfs'],
  [BRT.BeginStyles, BRT.EndStyles, 'styles'],
];

/** Parse `xl/styles.bin` into the flat cell-format table a worksheet's style indices resolve against. */
export function parseStyleTable(part: Uint8Array | undefined): StyleTable {
  if (part === undefined)
    return {cellXfs: [], namedStyles: [], preserved: NO_PRESERVED_STYLE_TABLES};

  const numFmtCodes = new Map<number, string>();
  const fonts: Array<Font | undefined> = [];
  const fills: Array<Fill | undefined> = [];
  const borders: Array<Border | undefined> = [];
  const namedXfs: XfStyle[] = [];
  const directXfs: XfStyle[] = [];
  const labels: StyleLabel[] = [];
  const blocks = blockTracker(COLLECTIONS);

  for (const record of readRecords(part)) {
    if (blocks.boundary(record.type)) continue;
    const reader = new RecordReader(record.data);
    switch (record.type) {
      case BRT.Fmt:
        if (blocks.isOpen('fmts')) numFmtCodes.set(reader.u16(), reader.wideString());
        break;
      case BRT.Font:
        if (blocks.isOpen('fonts')) fonts.push(readFont(reader));
        break;
      case BRT.Fill:
        if (blocks.isOpen('fills')) fills.push(readFill(reader));
        break;
      case BRT.Border:
        if (blocks.isOpen('borders')) borders.push(readBorder(reader));
        break;
      case BRT.XF:
        if (blocks.isOpen('cellXfs') || blocks.isOpen('cellStyleXfs')) {
          const deps = {fonts, fills, borders, numFmtCodes};
          (blocks.isOpen('cellXfs') ? directXfs : namedXfs).push(
            readXf(reader, deps, blocks.isOpen('cellXfs')),
          );
        }
        break;
      case BRT.Style:
        if (blocks.isOpen('styles')) labels.push(readStyleLabel(reader));
        break;
      default:
        break;
    }
  }

  // BIFF12 records carry no verbatim XML, so there is nothing of the four preserved sub-tables for
  // this reader to keep; the model is the whole of what a `.xlsb` stylesheet says.
  return {
    ...resolveStyleTable({directXfs, namedXfs, labels, fonts}),
    preserved: NO_PRESERVED_STYLE_TABLES,
  };
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
  // face), so only a custom index names an actual border: the same asymmetry the XML reader keeps.
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

/** The three fields of a BIFF12 xf record that carry alignment between them. */
interface AlignmentBits {
  /** `grbitAtr`-adjacent flag word: the two enumerations, the two booleans and the reading order. */
  readonly flags: number;
  /** `trot`, a byte of its own. */
  readonly rotation: number;
  /** `cIndent`, likewise. */
  readonly indent: number;
}

/**
 * Where each {@link Alignment} facet lives in a BIFF12 xf record, and what counts as its absence.
 *
 * A `Record` keyed by the facet name, so an eighth facet added to `Alignment` is a compile error
 * *here* as well as in `ALIGNMENT_FACETS`. That was the gap: `EveryAlignmentFacetIsDeclared` proves
 * the core table covers the type, and both XML directions walk that table, so a new facet lit up on
 * three of the four paths and was silently dropped by the binary reader -- which is precisely the
 * "facet added to the model, invisible in a file we read back" failure the table exists to prevent.
 *
 * The bit layout stays here rather than joining the core table, and that is the deliberate half of
 * the fix: `ALIGNMENT_FACETS` is format-blind on purpose (the layering gate forbids `core/` importing
 * a serialisation), so giving it masks and shifts would have made the model carry a BIFF12 fact to
 * spare this file a `Record`. The *set* of facets is decided once, in core; each format supplies its
 * own reading.
 *
 * BIFF12 states every field on every xf where XML omits the ones at their default, so each entry
 * answers `{}` for the value the XML reader would not have seen: `general` horizontal, `bottom`
 * vertical, and a zero rotation, indent or reading order are absences rather than values.
 */
const BIFF12_ALIGNMENT: Record<keyof Alignment, (bits: AlignmentBits) => Alignment> = {
  horizontal: ({flags}) => {
    const value = HORIZONTAL_ALIGNMENTS[flags & 0b111];
    return value === undefined ? {} : {horizontal: value};
  },
  vertical: ({flags}) => {
    const value = VERTICAL_ALIGNMENTS[(flags >> 3) & 0b111];
    return value === undefined ? {} : {vertical: value};
  },
  textRotation: ({rotation}) => (rotation === 0 ? {} : {textRotation: rotation}),
  wrapText: ({flags}) => ((flags & 0x0040) === 0 ? {} : {wrapText: true}),
  indent: ({indent}) => (indent === 0 ? {} : {indent}),
  shrinkToFit: ({flags}) => ((flags & 0x0100) === 0 ? {} : {shrinkToFit: true}),
  readingOrder: ({flags}) => {
    const value = (flags >> 10) & 0b11;
    return value === 0 ? {} : {readingOrder: value};
  },
};

function readAlignment(flags: number, rotation: number, indent: number): Alignment | undefined {
  const bits: AlignmentBits = {flags, rotation, indent};
  // Walking the core list rather than this file's own `Record` keys, so the *order* of facets is
  // decided in one place too and this reader cannot fall out of step with the writer on it.
  const out: Alignment = {};
  for (const facet of ALIGNMENT_FACETS) Object.assign(out, BIFF12_ALIGNMENT[facet.key](bits));
  return Object.keys(out).length > 0 ? out : undefined;
}

// The two protection bits, read through the rule both codecs share (`protectionFrom`): which of the
// two flags carries information is an OOXML default question, not a BIFF12 one.
function readProtection(flags: number): Protection | undefined {
  return protectionFrom({locked: (flags & 0x1000) !== 0, hidden: (flags & 0x2000) !== 0});
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

// `alcv`, likewise, with `bottom` (index 2) left out as the default.
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
// bare `<u/>`, which is what Excel writes for it, reads back as.
const UNDERLINE_STYLES: ReadonlyMap<number, UnderlineStyle> = new Map<number, UnderlineStyle>([
  [0x01, true],
  [0x02, 'double'],
  [0x21, 'singleAccounting'],
  [0x22, 'doubleAccounting'],
]);

// `BrtFill` ([MS-XLSB] 2.4.681). `fls` is an index into `ST_PatternType`, so the enumeration itself
// is the table: `FILL_PATTERNS_IN_SCHEMA_ORDER` is where it lives, beside the union it names, and
// this codec indexes it rather than carrying a third copy of the nineteen names. Index 0 is `none`,
// which the model spells as no fill at all, and an unmodelled pattern (a gradient, `fls` 0x28, whose
// stop-array layout this reader has no Excel-authored sample to check against) is dropped rather
// than guessed, so an unfilled cell reads back unfilled either way.
function readFill(reader: RecordReader): Fill | undefined {
  const raw = FILL_PATTERNS_IN_SCHEMA_ORDER[reader.u32()];
  const pattern = raw === 'none' ? undefined : raw;
  if (pattern === undefined) return undefined;
  // BIFF12 always states both colours; XML states only the ones the fill actually has, using the two
  // legacy-palette sentinels for the rest: 64 is "automatic foreground", 65 "automatic background".
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

// `BrtBorder` ([MS-XLSB] 2.4.314): the two diagonal-direction bits, then five `Blxf` edges in the
// order top, bottom, left, right, diagonal, which is *not* the model's or the schema's order, so
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
