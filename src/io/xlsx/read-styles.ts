// The style-table reader: `xl/styles.xml` in, a flat table of resolved cell formats out. It is a
// single streaming pass over the shared sub-tables (`<numFmts>`, `<fills>`, `<fonts>`, `<borders>`)
// and the two xf tables (`<cellXfs>`, `<cellStyleXfs>`), flattening the id-indirection so a cell's
// `s` index maps straight to its facets. A construct it does not recognise is skipped, never guessed.

import type {
  Alignment,
  Border,
  BorderStyle,
  Color,
  Fill,
  FillPatternType,
  Font,
  FontVerticalAlignment,
  GradientFill,
  GradientStop,
  HorizontalAlignment,
  NamedCellStyle,
  Protection,
  UnderlineStyle,
  VerticalAlignment,
} from '../../core/style.ts';
import {parseColor} from './styles.ts';
import {boolPresent, boolStrict, localName, parseXml} from './xml-read.ts';

// The style facets an xf resolves to. Absent facets stay undefined, matching the contract
// that an unset facet is simply not present on the reconstructed cell.
export interface XfStyle {
  readonly fill?: Fill;
  readonly numFmt?: string;
  readonly font?: Partial<Font>;
  readonly border?: Border;
  readonly alignment?: Alignment;
  readonly protection?: Protection;
  readonly quotePrefix?: boolean;
  /** The `xfId` link into the named-style layer (`cellStyleXfs`); absent for the Normal default (0). */
  readonly xfId?: number;
}

// A mutable xf accumulator while an <xf> element streams in: its facet ids resolve on open, but
// the <alignment>/<protection> children (when present) arrive before the element closes, so the
// xf is held here and pushed on close rather than on open.
type XfDraft = {-readonly [K in keyof XfStyle]?: XfStyle[K]};

// A mutable font accumulator while a <font> element's children stream in; frozen into a
// Partial<Font> on close.
export type FontDraft = {-readonly [K in keyof Font]?: Font[K]};

// A mutable border accumulator while a <border> element's edges stream in; frozen into a
// Border on close. The five edges match Border's; a bare styleless edge is simply never set.
type BorderDraft = {-readonly [K in keyof Border]?: Border[K]};

// A mutable gradient accumulator while a <gradientFill> streams in. `fill` builds up the frozen
// GradientFill (its stops appended as <stop>/<color> pairs close); `stopPosition`/`stopColor` hold the
// current <stop> until its close commits a {position, color} pair.
type GradientDraft = {
  fill: {-readonly [K in keyof GradientFill]: GradientFill[K]};
  stopPosition: number | null;
  stopColor: Color | undefined;
};

// The four sides plus the diagonal — the edge elements a <border> can hold, in the order the
// schema lists them. Membership drives edge parsing without a per-name branch.
type BorderEdgeName = 'left' | 'right' | 'top' | 'bottom' | 'diagonal';
const BORDER_EDGES = new Set<string>(['left', 'right', 'top', 'bottom', 'diagonal']);

// Style-table elements that commit on their close: a bare <font/>/<border/>/<patternFill/>/
// <gradientFill/>/<xf/> or a self-closing border edge is expanded to open+close so each commits
// exactly once in onClose, never in a duplicated (and easily-forgotten) self-closing branch.
const STYLE_EMPTY_CLOSES: ReadonlySet<string> = new Set([
  'font',
  'border',
  'patternFill',
  'gradientFill',
  'xf',
  ...BORDER_EDGES,
]);

// ECMA-376 reserves numFmt ids below 164 for formats every consumer knows implicitly, so a
// foreign file may name one with no <numFmt> entry. This maps the standard ids to their
// codes; id 0 (General) and any unknown id resolve to no format. The writer never emits
// these — it always defines a custom id — but reading them keeps foreign files faithful.
const BUILTIN_NUMFMTS: ReadonlyMap<number, string> = new Map([
  [1, '0'],
  [2, '0.00'],
  [3, '#,##0'],
  [4, '#,##0.00'],
  [9, '0%'],
  [10, '0.00%'],
  [11, '0.00E+00'],
  [12, '# ?/?'],
  [13, '# ??/??'],
  [14, 'mm-dd-yy'],
  [15, 'd-mmm-yy'],
  [16, 'd-mmm'],
  [17, 'mmm-yy'],
  [18, 'h:mm AM/PM'],
  [19, 'h:mm:ss AM/PM'],
  [20, 'h:mm'],
  [21, 'h:mm:ss'],
  [22, 'm/d/yy h:mm'],
  [37, '#,##0 ;(#,##0)'],
  [38, '#,##0 ;[Red](#,##0)'],
  [39, '#,##0.00;(#,##0.00)'],
  [40, '#,##0.00;[Red](#,##0.00)'],
  [45, 'mm:ss'],
  [46, '[h]:mm:ss'],
  [47, 'mmss.0'],
  [48, '##0.0E+0'],
  [49, '@'],
  // Ids 27..36 and 50..58 are reserved for locale-specific built-in East Asian date/time formats;
  // a file authored in a CJK locale styles date cells with them and, being built-ins, emits no
  // <numFmt>. The exact code is locale-defined — these are the representative Excel forms — but what
  // matters for reading is that each resolves to a non-empty date/time code so the serial reads as a
  // date rather than a bare number.
  [27, '[$-404]e/m/d'],
  [28, '[$-404]e"年"m"月"d"日"'],
  [29, '[$-404]e"年"m"月"d"日"'],
  [30, '[$-404]m/d/yy'],
  [31, '[$-404]yyyy"年"m"月"d"日"'],
  [32, '[$-404]h"時"mm"分"'],
  [33, '[$-404]h"時"mm"分"ss"秒"'],
  [34, '上午/下午h"時"mm"分"'],
  [35, '上午/下午h"時"mm"分"ss"秒"'],
  [36, '[$-404]e/m/d'],
  [50, '[$-404]e/m/d'],
  [51, '[$-404]e"年"m"月"d"日"'],
  [52, '[$-404]yyyy"年"m"月"'],
  [53, '[$-404]m"月"d"日"'],
  [54, '[$-404]e"年"m"月"d"日"'],
  [55, '上午/下午h"時"mm"分"'],
  [56, '上午/下午h"時"mm"分"ss"秒"'],
  [57, '[$-404]yyyy"年"m"月"'],
  [58, '[$-404]m"月"d"日"'],
]);

// styles.xml is a shared table: <numFmts> defines custom format codes by id, <fills> lists
// the fills, and <cellXfs> lists the cell formats, each naming a fill and a number format by
// id. We flatten that indirection into one array — cellXfs index → resolved {fill, numFmt} —
// so a cell/row/column style index maps straight to its facets. The schema orders <numFmts>
// and <fills> before <cellXfs>, so both lookups are complete before an xf references them.
/** The parsed style table: the cell formats a cell/row/column `s` indexes, plus the named cell-style
 * layer a cell's `xfId` links into. Each cellXfs entry is already merged with its named style, so a
 * cell reading its `s` sees the effective facets; the `xfId` link is carried through for re-write. */
export interface StyleTable {
  readonly cellXfs: ReadonlyArray<XfStyle>;
  readonly namedStyles: ReadonlyArray<NamedCellStyle>;
}

export function parseStyleTable(xml: string): StyleTable {
  if (xml === '') return {cellXfs: [], namedStyles: []};
  const fills: Array<Fill | undefined> = [];
  const fonts: Array<Partial<Font> | undefined> = [];
  const borders: Array<Border | undefined> = [];
  const numFmtCodes = new Map<number, string>();
  const xfStyles: XfStyle[] = [];
  // The named-style layer: <cellStyleXfs> holds the base formats a cell's xfId links to; <cellStyles>
  // labels them by name/builtinId. Parsed in parallel with cellXfs, then zipped and merged below.
  const namedXfs: XfStyle[] = [];
  const cellStyleNames: {xfId: number; name?: string; builtinId?: number}[] = [];
  let inFills = false;
  let inFonts = false;
  let inCellXfs = false;
  let inCellStyleXfs = false;
  let pattern = '';
  let fgColor: Color | undefined;
  let bgColor: Color | undefined;
  // A gradient fill accumulates from <gradientFill> open to close; its stops fill in as <stop>/<color>
  // pairs arrive. `fillSlotAt` marks where in `fills` the current <fill> began, so its close can keep a
  // slot even when the fill body was neither a pattern nor a gradient — index alignment is load-bearing.
  let gradientDraft: GradientDraft | null = null;
  let fillSlotAt = -1;
  let fontDraft: FontDraft | null = null;
  let borderDraft: BorderDraft | null = null;
  // Which edge of the current border a nested <color> belongs to; null between edges.
  let currentEdge: BorderEdgeName | null = null;
  // The <cellXfs>/<cellStyleXfs> xf being read; held from open to close so its <alignment>/
  // <protection> children can attach before the xf is committed. null outside an <xf>. `pendingXfTarget`
  // records which table the held xf belongs to.
  let pendingXf: XfDraft | null = null;
  let pendingXfTarget: 'cell' | 'named' | null = null;

  parseXml(
    xml,
    {
      onOpen(name, attrs) {
        switch (localName(name)) {
          case 'numFmt': {
            // <numFmts> entries are self-closing, so they are read here on open. A code with
            // no id, or the General id 0, contributes nothing.
            const id = Number(attrs.numFmtId);
            if (Number.isInteger(id) && id > 0 && attrs.formatCode !== undefined) {
              numFmtCodes.set(id, attrs.formatCode);
            }
            break;
          }
          case 'fills':
            inFills = true;
            break;
          case 'fill':
            // Mark where this <fill> starts so its close can guarantee exactly one slot — a fill body
            // that is neither <patternFill> nor <gradientFill> (or a gradient we could not parse) must
            // still consume an id, or every later fill index shifts and cells mis-resolve their fill.
            if (inFills) fillSlotAt = fills.length;
            break;
          case 'gradientFill':
            if (inFills) {
              gradientDraft = {
                fill: {
                  type: 'gradient',
                  gradient: attrs.type === 'path' ? 'path' : 'linear',
                  stops: [],
                },
                stopPosition: null,
                stopColor: undefined,
              };
              assignGradientNumbers(gradientDraft.fill, attrs);
            }
            break;
          case 'stop':
            if (gradientDraft !== null) {
              const position = Number(attrs.position);
              gradientDraft.stopPosition = Number.isFinite(position) ? position : 0;
              gradientDraft.stopColor = undefined;
            }
            break;
          case 'fonts':
            inFonts = true;
            break;
          case 'font':
            if (inFonts) fontDraft = {};
            break;
          case 'borders':
            break;
          case 'border':
            borderDraft = {};
            currentEdge = null;
            if (boolStrict(attrs.diagonalUp)) borderDraft.diagonalUp = true;
            if (boolStrict(attrs.diagonalDown)) borderDraft.diagonalDown = true;
            break;
          case 'cellXfs':
            inCellXfs = true;
            break;
          case 'cellStyleXfs':
            inCellStyleXfs = true;
            break;
          case 'cellStyle': {
            // A <cellStyle> (inside <cellStyles>) names a cellStyleXfs entry by xfId; it is self-closing,
            // so it is read here on open.
            const xfId = Number(attrs.xfId);
            if (Number.isInteger(xfId)) {
              const entry: {xfId: number; name?: string; builtinId?: number} = {xfId};
              if (attrs.name !== undefined) entry.name = attrs.name;
              if (attrs.builtinId !== undefined) {
                const builtinId = Number(attrs.builtinId);
                if (Number.isInteger(builtinId)) entry.builtinId = builtinId;
              }
              cellStyleNames.push(entry);
            }
            break;
          }
          case 'patternFill':
            if (inFills) {
              pattern = attrs.patternType ?? 'none';
              fgColor = undefined;
              bgColor = undefined;
            }
            break;
          case 'fgColor':
            if (inFills) fgColor = parseColor(attrs);
            break;
          case 'bgColor':
            if (inFills) bgColor = parseColor(attrs);
            break;
          default: {
            const local = localName(name);
            // A <font>'s children are self-closing, so they are read here on open.
            if (gradientDraft !== null && local === 'color') {
              // The colour of the open <stop>; committed to a GradientStop when the stop closes.
              gradientDraft.stopColor = parseColor(attrs);
            } else if (fontDraft !== null) {
              applyFontChild(fontDraft, local, attrs);
            } else if (borderDraft !== null) {
              // A border's edges and their <color> children are all read on open (each is
              // self-closing bar a coloured edge, whose colour child is itself self-closing).
              if (BORDER_EDGES.has(local)) {
                currentEdge = attrs.style !== undefined ? (local as BorderEdgeName) : null;
                if (currentEdge !== null)
                  borderDraft[currentEdge] = {style: attrs.style as BorderStyle};
              } else if (local === 'color' && currentEdge !== null) {
                const edge = borderDraft[currentEdge];
                if (edge !== undefined)
                  borderDraft[currentEdge] = {style: edge.style, color: parseColor(attrs)};
              }
            } else if (pendingXf !== null && local === 'alignment') {
              // An xf's <alignment> child arrives before the xf closes; attach it to the pending xf.
              const alignment = parseAlignment(attrs);
              if (alignment !== undefined) pendingXf.alignment = alignment;
            } else if (pendingXf !== null && local === 'protection') {
              // An xf's <protection> child likewise arrives before the xf closes.
              const protection = parseProtection(attrs);
              if (protection !== undefined) pendingXf.protection = protection;
            } else if ((inCellXfs || inCellStyleXfs) && local === 'xf') {
              // Both <cellXfs> and <cellStyleXfs> hold <xf> with identical structure; they differ only
              // in which table the result lands in and whether an xfId link is meaningful (only cellXfs
              // entries link to a named style).
              const fillId = Number(attrs.fillId);
              const fill = Number.isInteger(fillId) ? fills[fillId] : undefined;
              const fontId = Number(attrs.fontId);
              // Font id 0 is the workbook default font (a real Calibri-11-style face), not an
              // absence — unlike border id 0, which is a genuinely empty border. So an xf naming
              // font 0 resolves to that default face, giving every cell a concrete font to render.
              const font = Number.isInteger(fontId) ? fonts[fontId] : undefined;
              const borderId = Number(attrs.borderId);
              // Border id 0 is the empty default; only a custom border (id > 0) is an explicit one.
              const border =
                Number.isInteger(borderId) && borderId > 0 ? borders[borderId] : undefined;
              const numFmt = resolveNumFmt(attrs.numFmtId, numFmtCodes);
              const draft: XfDraft = {};
              if (fill) draft.fill = fill;
              if (numFmt !== undefined) draft.numFmt = numFmt;
              if (font) draft.font = font;
              if (border) draft.border = border;
              // The quote-prefix flag is an attribute on the xf itself (no shared sub-table); carry it
              // only when set so an ordinary cell does not gain a spurious `quotePrefix: false`.
              if (boolStrict(attrs.quotePrefix)) draft.quotePrefix = true;
              // A cellXfs entry's xfId links it to a named style; capture it only when it points beyond
              // the Normal default (0), so an ordinary cell carries no spurious named-style link.
              if (inCellXfs && attrs.xfId !== undefined) {
                const xfId = Number(attrs.xfId);
                if (Number.isInteger(xfId) && xfId > 0) draft.xfId = xfId;
              }
              // Hold the xf open until its close so an <alignment>/<protection> child can attach first;
              // a self-closing <xf/> is expanded to a close, so it commits there too, child-free.
              pendingXf = draft;
              pendingXfTarget = inCellXfs ? 'cell' : 'named';
            }
            break;
          }
        }
      },
      onClose(name) {
        switch (localName(name)) {
          case 'fills':
            inFills = false;
            break;
          case 'fonts':
            inFonts = false;
            break;
          case 'font':
            if (fontDraft !== null) {
              fonts.push(Object.keys(fontDraft).length > 0 ? fontDraft : undefined);
              fontDraft = null;
            }
            break;
          case 'border':
            if (borderDraft !== null) {
              borders.push(borderToStyle(borderDraft));
              borderDraft = null;
              currentEdge = null;
            }
            break;
          case 'cellXfs':
            inCellXfs = false;
            break;
          case 'cellStyleXfs':
            inCellStyleXfs = false;
            break;
          case 'xf':
            // A held (non-self-closing) xf commits here, with any alignment child attached, into
            // whichever table it was opened in.
            if (pendingXf !== null) {
              (pendingXfTarget === 'named' ? namedXfs : xfStyles).push(pendingXf);
              pendingXf = null;
              pendingXfTarget = null;
            }
            break;
          case 'patternFill':
            if (inFills) fills.push(toFill(pattern, fgColor, bgColor));
            break;
          case 'stop':
            if (gradientDraft !== null && gradientDraft.stopPosition !== null) {
              const stop: GradientStop = {
                position: gradientDraft.stopPosition,
                color: gradientDraft.stopColor ?? {},
              };
              gradientDraft.fill.stops = [...gradientDraft.fill.stops, stop];
              gradientDraft.stopPosition = null;
              gradientDraft.stopColor = undefined;
            }
            break;
          case 'gradientFill':
            if (gradientDraft !== null) {
              fills.push(gradientDraft.fill);
              gradientDraft = null;
            }
            break;
          case 'fill':
            // Backstop the slot: if this <fill>'s body pushed nothing (unparsed/unknown content), keep an
            // empty slot so id alignment holds and later fills still resolve to the right cells.
            if (inFills && fills.length === fillSlotAt) fills.push(undefined);
            break;
          default:
            // A coloured edge closes after its <color> child; drop the edge context so a stray
            // later <color> cannot attach to it.
            if (borderDraft !== null && BORDER_EDGES.has(localName(name))) currentEdge = null;
            break;
        }
      },
    },
    {closeEmptyElements: STYLE_EMPTY_CLOSES},
  );

  // Layer each cellXfs entry over the named style its xfId links to: a facet the direct format sets
  // wins; one it leaves unset falls through to the named style. The xfId is carried through so the
  // link survives a re-write. A draft only holds keys for facets it actually set, so the spread merge
  // takes the named base and lets the direct entry override exactly what it names.
  const cellXfs: XfStyle[] = xfStyles.map((xf) => {
    if (xf.xfId === undefined) return xf;
    const named = namedXfs[xf.xfId];
    return named === undefined ? xf : {...named, ...xf};
  });

  // Zip the resolved cellStyleXfs facets with their cellStyles name/builtinId into the model's named
  // styles, index for index (a cellStyle's xfId is its cellStyleXfs index).
  const namedStyles: NamedCellStyle[] = namedXfs.map((xf, index) => {
    const label = cellStyleNames.find((entry) => entry.xfId === index);
    const style: {-readonly [K in keyof NamedCellStyle]?: NamedCellStyle[K]} = {};
    if (xf.fill !== undefined) style.fill = xf.fill;
    if (xf.numFmt !== undefined) style.numFmt = xf.numFmt;
    if (xf.font !== undefined) style.font = xf.font;
    if (xf.border !== undefined) style.border = xf.border;
    if (xf.alignment !== undefined) style.alignment = xf.alignment;
    if (xf.protection !== undefined) style.protection = xf.protection;
    if (label?.name !== undefined) style.name = label.name;
    if (label?.builtinId !== undefined) style.builtinId = label.builtinId;
    return style;
  });

  return {cellXfs, namedStyles};
}

// A <font> child element sets one facet on the draft. Boolean flags honour their `val`: a
// bare tag or val="1"/"true" is on, val="0"/"false" is off (an explicit-false flag is not
// truthy merely because the tag is present). An unrecognised child is ignored.
export function applyFontChild(
  draft: FontDraft,
  local: string,
  attrs: {readonly [k: string]: string},
): void {
  switch (local) {
    case 'b':
      draft.bold = boolPresent(attrs.val);
      break;
    case 'i':
      draft.italic = boolPresent(attrs.val);
      break;
    case 'strike':
      draft.strike = boolPresent(attrs.val);
      break;
    case 'outline':
      draft.outline = boolPresent(attrs.val);
      break;
    case 'u':
      // A bare <u/> is a single underline; a named style (single/double/…) carries through; but
      // val="none" is the explicit ABSENCE of an underline, so it must read back falsy — not the
      // truthy string "none" that a consumer's `if (font.underline)` would mistake for underlined.
      draft.underline =
        attrs.val === undefined
          ? true
          : attrs.val === 'none'
            ? false
            : (attrs.val as UnderlineStyle);
      break;
    case 'vertAlign':
      if (attrs.val !== undefined) draft.vertAlign = attrs.val as FontVerticalAlignment;
      break;
    case 'sz': {
      const size = Number(attrs.val);
      if (Number.isFinite(size)) draft.size = size;
      break;
    }
    case 'color':
      draft.color = parseColor(attrs);
      break;
    // `<name>` in a styles `<font>`, `<rFont>` in a rich-text run's `<rPr>` — the same font face.
    case 'name':
    case 'rFont':
      if (attrs.val !== undefined) draft.name = attrs.val;
      break;
    case 'family': {
      const family = Number(attrs.val);
      if (Number.isInteger(family)) draft.family = family;
      break;
    }
    case 'charset': {
      const charset = Number(attrs.val);
      if (Number.isInteger(charset)) draft.charset = charset;
      break;
    }
    case 'scheme':
      if (attrs.val !== undefined) draft.scheme = attrs.val as Font['scheme'];
      break;
    default:
      break;
  }
}

// An xf's numFmtId resolves against the custom codes first, then the built-in table; the
// General format (id 0) and any unrecognised id mean the cell carries no explicit format.
function resolveNumFmt(
  raw: string | undefined,
  custom: ReadonlyMap<number, string>,
): string | undefined {
  if (raw === undefined) return undefined;
  const id = Number(raw);
  if (!Number.isInteger(id) || id === 0) return undefined;
  return custom.get(id) ?? BUILTIN_NUMFMTS.get(id);
}

function toFill(
  pattern: string,
  fgColor: Color | undefined,
  bgColor: Color | undefined,
): Fill | undefined {
  if (pattern === '' || pattern === 'none') return undefined;
  return {
    type: 'pattern',
    pattern: pattern as FillPatternType,
    ...(fgColor ? {fgColor} : {}),
    ...(bgColor ? {bgColor} : {}),
  };
}

// Copy the numeric <gradientFill> attributes (degree; the path insets) onto a gradient draft, keeping
// only the finite ones so an absent or malformed attribute leaves the field its OOXML default (unset).
function assignGradientNumbers(fill: GradientDraft['fill'], attrs: Record<string, string>): void {
  for (const key of ['degree', 'left', 'right', 'top', 'bottom'] as const) {
    const value = Number(attrs[key]);
    if (attrs[key] !== undefined && Number.isFinite(value)) fill[key] = value;
  }
}

// An accumulated border with no styled edge and no diagonal direction is the empty default:
// it carries nothing, so it resolves to undefined rather than an all-empty Border object.
function borderToStyle(draft: BorderDraft): Border | undefined {
  const hasEdge = (['left', 'right', 'top', 'bottom', 'diagonal'] as const).some(
    (edge: BorderEdgeName): boolean => draft[edge] !== undefined,
  );
  if (!hasEdge && draft.diagonalUp === undefined && draft.diagonalDown === undefined)
    return undefined;
  return draft;
}

// Read an <alignment> element's attributes into an Alignment, keeping only facets that differ
// from the default. Boolean flags honour their parsed value — wrapText="0" is off, so it must
// not fabricate a { wrapText: false } alignment — and an element carrying only defaults yields
// undefined rather than an empty alignment object.
function parseAlignment(attrs: {readonly [k: string]: string}): Alignment | undefined {
  const out: {-readonly [K in keyof Alignment]?: Alignment[K]} = {};
  if (attrs.horizontal !== undefined && attrs.horizontal !== 'general') {
    out.horizontal = attrs.horizontal as HorizontalAlignment;
  }
  if (attrs.vertical !== undefined) out.vertical = attrs.vertical as VerticalAlignment;
  if (attrs.textRotation !== undefined) {
    const rotation = Number(attrs.textRotation);
    if (Number.isFinite(rotation) && rotation !== 0) out.textRotation = rotation;
  }
  if (boolStrict(attrs.wrapText)) out.wrapText = true;
  if (attrs.indent !== undefined) {
    const indent = Number(attrs.indent);
    if (Number.isInteger(indent) && indent !== 0) out.indent = indent;
  }
  if (boolStrict(attrs.shrinkToFit)) out.shrinkToFit = true;
  if (attrs.readingOrder !== undefined) {
    const order = Number(attrs.readingOrder);
    if (Number.isInteger(order) && order !== 0) out.readingOrder = order;
  }
  return Object.keys(out).length > 0 ? out : undefined;
}

// Read a <protection> element into a Protection, keeping only facets that differ from the OOXML
// default. `locked` defaults to TRUE, so only an explicit `locked="0"` carries information (an
// unlocked cell) — a default or explicit-true cell must not read back as { locked: true }; `hidden`
// defaults to false, so only `hidden="1"` is carried. An element with only defaults yields undefined.
function parseProtection(attrs: {readonly [k: string]: string}): Protection | undefined {
  const out: {-readonly [K in keyof Protection]?: Protection[K]} = {};
  if (attrs.locked === '0' || attrs.locked === 'false') out.locked = false;
  if (boolStrict(attrs.hidden)) out.hidden = true;
  return Object.keys(out).length > 0 ? out : undefined;
}
