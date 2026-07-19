// The style-table reader: `xl/styles.xml` in, a flat table of resolved cell formats out. It is a
// single streaming pass over the shared sub-tables (`<numFmts>`, `<fills>`, `<fonts>`, `<borders>`)
// and the two xf tables (`<cellXfs>`, `<cellStyleXfs>`), flattening the id-indirection so a cell's
// `s` index maps straight to its facets. A construct it does not recognise is skipped, never guessed.

import {
  type Alignment,
  type Border,
  type Color,
  type Fill,
  type FillPatternType,
  type Font,
  type GradientFill,
  type GradientStop,
  isBorderStyle,
  isFontScheme,
  isFontVerticalAlignment,
  isHorizontalAlignment,
  isNamedUnderlineStyle,
  isVerticalAlignment,
  type NamedCellStyle,
  type Protection,
} from '../../core/style.ts';
import {parseColor} from './styles.ts';
import {
  boolPresent,
  boolStrict,
  closeEmptyElements,
  localName,
  type XmlAttributes,
  type XmlEvent,
  xmlEvents,
} from './xml-read.ts';

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
  let fills: ReadonlyArray<Fill | undefined> = [];
  let fonts: ReadonlyArray<Partial<Font> | undefined> = [];
  let borders: ReadonlyArray<Border | undefined> = [];
  let numFmtCodes: ReadonlyMap<number, string> = new Map();
  const xfStyles: XfStyle[] = [];
  // The named-style layer: <cellStyleXfs> holds the base formats a cell's xfId links to; <cellStyles>
  // labels them by name/builtinId. Parsed in parallel with cellXfs, then zipped and merged below.
  const namedXfs: XfStyle[] = [];
  let cellStyleNames: ReadonlyArray<CellStyleName> = [];

  // One streaming pass, but each top-level sub-table drives its own focused sub-parser over the slice
  // of events between its open and close. The schema orders the shared tables (<numFmts>, <fonts>,
  // <fills>, <borders>) before the xf tables, so their results are complete before an <xf> resolves
  // against them. Every recognised container name is plural and unique to the styleSheet root — none
  // appears inside a <dxf>'s singular <font>/<fill>/<border> children — so skipping an unrecognised
  // section here drops exactly what the old flat pass gated off with its `in*` flags.
  const events = closeEmptyElements(xmlEvents(xml), STYLE_EMPTY_CLOSES);
  let next = events.next();
  while (next.done !== true) {
    const event = next.value;
    if (event.kind === 'open' && !event.selfClosing) {
      switch (localName(event.name)) {
        case 'numFmts':
          numFmtCodes = parseNumFmts(events);
          break;
        case 'fonts':
          fonts = parseFonts(events);
          break;
        case 'fills':
          fills = parseFills(events);
          break;
        case 'borders':
          borders = parseBorders(events);
          break;
        case 'cellStyleXfs':
          namedXfs.push(
            ...parseXfTable(events, 'cellStyleXfs', {fills, fonts, borders, numFmtCodes}),
          );
          break;
        case 'cellXfs':
          xfStyles.push(...parseXfTable(events, 'cellXfs', {fills, fonts, borders, numFmtCodes}));
          break;
        case 'cellStyles':
          cellStyleNames = parseCellStyles(events);
          break;
      }
    }
    next = events.next();
  }

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

// A <cellStyle> label: the name/builtinId a cellStyleXfs entry carries, keyed by its xfId (its index).
interface CellStyleName {
  readonly xfId: number;
  readonly name?: string;
  readonly builtinId?: number;
}

// The shared sub-tables an <xf> resolves its facet ids against.
interface XfDeps {
  readonly fills: ReadonlyArray<Fill | undefined>;
  readonly fonts: ReadonlyArray<Partial<Font> | undefined>;
  readonly borders: ReadonlyArray<Border | undefined>;
  readonly numFmtCodes: ReadonlyMap<number, string>;
}

// Pull events off the shared stream up to — and consuming — the close of `container`, yielding only
// those strictly inside it. A sub-table parser loops this to completion (never breaking), so it drives
// its own small state machine over exactly its section without closing the underlying generator, and
// the outer pass resumes at the element after the container's close.
function* until(events: Iterator<XmlEvent>, container: string): Generator<XmlEvent> {
  let next = events.next();
  while (next.done !== true) {
    const event = next.value;
    if (event.kind === 'close' && localName(event.name) === container) return;
    yield event;
    next = events.next();
  }
}

// <numFmts> entries are self-closing, so they are read on open. A code with no id, or the General
// id 0, contributes nothing.
function parseNumFmts(events: Iterator<XmlEvent>): ReadonlyMap<number, string> {
  const codes = new Map<number, string>();
  for (const event of until(events, 'numFmts')) {
    if (event.kind === 'open' && localName(event.name) === 'numFmt') {
      const id = Number(event.attrs.numFmtId);
      if (Number.isInteger(id) && id > 0 && event.attrs.formatCode !== undefined) {
        codes.set(id, event.attrs.formatCode);
      }
    }
  }
  return codes;
}

function parseFonts(events: Iterator<XmlEvent>): ReadonlyArray<Partial<Font> | undefined> {
  const fonts: Array<Partial<Font> | undefined> = [];
  let fontDraft: FontDraft | null = null;
  for (const event of until(events, 'fonts')) {
    if (event.kind === 'open') {
      const local = localName(event.name);
      // A <font>'s children are self-closing, so they are read here on open.
      if (local === 'font') fontDraft = {};
      else if (fontDraft !== null) applyFontChild(fontDraft, local, event.attrs);
    } else if (event.kind === 'close' && localName(event.name) === 'font' && fontDraft !== null) {
      fonts.push(Object.keys(fontDraft).length > 0 ? fontDraft : undefined);
      fontDraft = null;
    }
  }
  return fonts;
}

function parseFills(events: Iterator<XmlEvent>): ReadonlyArray<Fill | undefined> {
  const fills: Array<Fill | undefined> = [];
  let pattern = '';
  let fgColor: Color | undefined;
  let bgColor: Color | undefined;
  // A gradient fill accumulates from <gradientFill> open to close; its stops fill in as <stop>/<color>
  // pairs arrive. `fillSlotAt` marks where in `fills` the current <fill> began, so its close can keep a
  // slot even when the fill body was neither a pattern nor a gradient — index alignment is load-bearing.
  let gradientDraft: GradientDraft | null = null;
  let fillSlotAt = -1;
  for (const event of until(events, 'fills')) {
    if (event.kind === 'open') {
      const attrs = event.attrs;
      switch (localName(event.name)) {
        case 'fill':
          // Mark where this <fill> starts so its close can guarantee exactly one slot — a fill body
          // that is neither <patternFill> nor <gradientFill> (or a gradient we could not parse) must
          // still consume an id, or every later fill index shifts and cells mis-resolve their fill.
          fillSlotAt = fills.length;
          break;
        case 'patternFill':
          pattern = attrs.patternType ?? 'none';
          fgColor = undefined;
          bgColor = undefined;
          break;
        case 'fgColor':
          fgColor = parseColor(attrs);
          break;
        case 'bgColor':
          bgColor = parseColor(attrs);
          break;
        case 'gradientFill':
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
          break;
        case 'stop':
          if (gradientDraft !== null) {
            const position = Number(attrs.position);
            gradientDraft.stopPosition = Number.isFinite(position) ? position : 0;
            gradientDraft.stopColor = undefined;
          }
          break;
        case 'color':
          // The colour of the open <stop>; committed to a GradientStop when the stop closes.
          if (gradientDraft !== null) gradientDraft.stopColor = parseColor(attrs);
          break;
      }
    } else if (event.kind === 'close') {
      switch (localName(event.name)) {
        case 'patternFill':
          fills.push(toFill(pattern, fgColor, bgColor));
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
          if (fills.length === fillSlotAt) fills.push(undefined);
          break;
      }
    }
  }
  return fills;
}

function parseBorders(events: Iterator<XmlEvent>): ReadonlyArray<Border | undefined> {
  const borders: Array<Border | undefined> = [];
  let borderDraft: BorderDraft | null = null;
  // Which edge of the current border a nested <color> belongs to; null between edges.
  let currentEdge: BorderEdgeName | null = null;
  for (const event of until(events, 'borders')) {
    if (event.kind === 'open') {
      const local = localName(event.name);
      const attrs = event.attrs;
      if (local === 'border') {
        borderDraft = {};
        currentEdge = null;
        if (boolStrict(attrs.diagonalUp)) borderDraft.diagonalUp = true;
        if (boolStrict(attrs.diagonalDown)) borderDraft.diagonalDown = true;
      } else if (borderDraft !== null) {
        // A border's edges and their <color> children are all read on open (each is self-closing bar a
        // coloured edge, whose colour child is itself self-closing). An edge whose style is absent or
        // an unrecognised token is dropped — the side simply carries no border.
        if (BORDER_EDGES.has(local)) {
          if (attrs.style !== undefined && isBorderStyle(attrs.style)) {
            currentEdge = local as BorderEdgeName;
            borderDraft[currentEdge] = {style: attrs.style};
          } else {
            currentEdge = null;
          }
        } else if (local === 'color' && currentEdge !== null) {
          const edge = borderDraft[currentEdge];
          if (edge !== undefined)
            borderDraft[currentEdge] = {style: edge.style, color: parseColor(attrs)};
        }
      }
    } else if (event.kind === 'close') {
      const local = localName(event.name);
      if (local === 'border') {
        if (borderDraft !== null) {
          borders.push(borderToStyle(borderDraft));
          borderDraft = null;
          currentEdge = null;
        }
      } else if (borderDraft !== null && BORDER_EDGES.has(local)) {
        // A coloured edge closes after its <color> child; drop the edge context so a stray later
        // <color> cannot attach to it.
        currentEdge = null;
      }
    }
  }
  return borders;
}

// Both <cellXfs> and <cellStyleXfs> hold <xf> with identical structure; they differ only in which
// table the result lands in and whether an xfId link is meaningful (only cellXfs entries link to a
// named style). One parser serves both, told by `container` which it is reading.
function parseXfTable(
  events: Iterator<XmlEvent>,
  container: 'cellXfs' | 'cellStyleXfs',
  deps: XfDeps,
): XfStyle[] {
  const xfs: XfStyle[] = [];
  const captureXfId = container === 'cellXfs';
  // The xf being read; held from open to close so its <alignment>/<protection> children can attach
  // before it is committed. null outside an <xf>.
  let pendingXf: XfDraft | null = null;
  for (const event of until(events, container)) {
    if (event.kind === 'open') {
      const local = localName(event.name);
      if (local === 'xf') {
        // Hold the xf open until its close so an <alignment>/<protection> child can attach first; a
        // self-closing <xf/> is expanded to a close, so it commits there too, child-free.
        pendingXf = resolveXf(event.attrs, deps, captureXfId);
      } else if (pendingXf !== null && local === 'alignment') {
        // An xf's <alignment> child arrives before the xf closes; attach it to the pending xf.
        const alignment = parseAlignment(event.attrs);
        if (alignment !== undefined) pendingXf.alignment = alignment;
      } else if (pendingXf !== null && local === 'protection') {
        // An xf's <protection> child likewise arrives before the xf closes.
        const protection = parseProtection(event.attrs);
        if (protection !== undefined) pendingXf.protection = protection;
      }
    } else if (event.kind === 'close' && localName(event.name) === 'xf' && pendingXf !== null) {
      xfs.push(pendingXf);
      pendingXf = null;
    }
  }
  return xfs;
}

// Resolve an <xf>'s facet ids against the shared sub-tables into a draft. `captureXfId` is set only
// for cellXfs entries, the sole table whose xfId links to a named style.
function resolveXf(attrs: XmlAttributes, deps: XfDeps, captureXfId: boolean): XfDraft {
  const fillId = Number(attrs.fillId);
  const fill = Number.isInteger(fillId) ? deps.fills[fillId] : undefined;
  const fontId = Number(attrs.fontId);
  // Font id 0 is the workbook default font (a real Calibri-11-style face), not an absence — unlike
  // border id 0, which is a genuinely empty border. So an xf naming font 0 resolves to that default
  // face, giving every cell a concrete font to render.
  const font = Number.isInteger(fontId) ? deps.fonts[fontId] : undefined;
  const borderId = Number(attrs.borderId);
  // Border id 0 is the empty default; only a custom border (id > 0) is an explicit one.
  const border = Number.isInteger(borderId) && borderId > 0 ? deps.borders[borderId] : undefined;
  const numFmt = resolveNumFmt(attrs.numFmtId, deps.numFmtCodes);
  const draft: XfDraft = {};
  if (fill) draft.fill = fill;
  if (numFmt !== undefined) draft.numFmt = numFmt;
  if (font) draft.font = font;
  if (border) draft.border = border;
  // The quote-prefix flag is an attribute on the xf itself (no shared sub-table); carry it only when
  // set so an ordinary cell does not gain a spurious `quotePrefix: false`.
  if (boolStrict(attrs.quotePrefix)) draft.quotePrefix = true;
  // A cellXfs entry's xfId links it to a named style; capture it only when it points beyond the Normal
  // default (0), so an ordinary cell carries no spurious named-style link.
  if (captureXfId && attrs.xfId !== undefined) {
    const xfId = Number(attrs.xfId);
    if (Number.isInteger(xfId) && xfId > 0) draft.xfId = xfId;
  }
  return draft;
}

// A <cellStyle> (inside <cellStyles>) names a cellStyleXfs entry by xfId; it is self-closing, so it
// is read on open.
function parseCellStyles(events: Iterator<XmlEvent>): ReadonlyArray<CellStyleName> {
  const names: CellStyleName[] = [];
  for (const event of until(events, 'cellStyles')) {
    if (event.kind === 'open' && localName(event.name) === 'cellStyle') {
      const attrs = event.attrs;
      const xfId = Number(attrs.xfId);
      if (Number.isInteger(xfId)) {
        const entry: {xfId: number; name?: string; builtinId?: number} = {xfId};
        if (attrs.name !== undefined) entry.name = attrs.name;
        if (attrs.builtinId !== undefined) {
          const builtinId = Number(attrs.builtinId);
          if (Number.isInteger(builtinId)) entry.builtinId = builtinId;
        }
        names.push(entry);
      }
    }
  }
  return names;
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
      // truthy string "none" that a consumer's `if (font.underline)` would mistake for underlined. An
      // unrecognised token keeps the "is underlined" fact but drops the unknown style (a plain true).
      draft.underline =
        attrs.val === undefined
          ? true
          : attrs.val === 'none'
            ? false
            : isNamedUnderlineStyle(attrs.val)
              ? attrs.val
              : true;
      break;
    case 'vertAlign':
      if (attrs.val !== undefined && isFontVerticalAlignment(attrs.val))
        draft.vertAlign = attrs.val;
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
      if (attrs.val !== undefined && isFontScheme(attrs.val)) draft.scheme = attrs.val;
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
  // `general` is the default and reads back as no explicit horizontal alignment; an unrecognised
  // token (like an out-of-enum vertical one) is dropped rather than trusted into the model.
  if (
    attrs.horizontal !== undefined &&
    attrs.horizontal !== 'general' &&
    isHorizontalAlignment(attrs.horizontal)
  ) {
    out.horizontal = attrs.horizontal;
  }
  if (attrs.vertical !== undefined && isVerticalAlignment(attrs.vertical))
    out.vertical = attrs.vertical;
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
