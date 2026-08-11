// The worksheet's non-cell property blocks: `<sheetPr>`, `<sheetViews>`, `<sheetProtection>`,
// `<autoFilter>`, and the print settings (`<printOptions>`, `<pageMargins>`, `<pageSetup>`,
// `<headerFooter>`, `<rowBreaks>`/`<colBreaks>`). Each renders one CT_Worksheet child (or child
// group) independently of the row/cell body `worksheet-xml.ts` orchestrates them alongside.

import {encodeAddress} from '../../core/address.ts';
import type {AutoFilter, FilterColumn, FilterCriteria} from '../../core/autofilter.ts';
import type {
  HeaderFooter,
  PageBreak,
  PageMargins,
  PageSetup,
  PrintOptions,
} from '../../core/page-setup.ts';
import {SHEET_PROTECTION_FLAGS, type SheetProtection} from '../../core/protection.ts';
import type {OutlineProperties, SheetView, Worksheet} from '../../core/worksheet.ts';
import {attr, boolAttr, escapeAttr, escapeText, numberText} from '../../xml/xml.ts';
import {colorAttrs} from './color-xml.ts';

// `<sheetViews>` holds the sheet's single view. A frozen view adds a `<pane>` recording the split
// and a `<selection>` naming the pane the split activates, exactly as Excel writes it — a normal
// view carries neither, so unfreezing leaves no leftover `<pane>` that would trip a repair prompt.
// The active pane is whichever scrolling region the freeze creates: bottom-right when both axes are
// frozen, else top-right (columns only) or bottom-left (rows only).
//
// `active` marks this sheet as the one selected on open (`tabSelected`). Exactly one sheet in a
// workbook carries it — with none, the consumer opens with no sheet view initialised; with several,
// the sheets form a *group selection*, where an edit to one is applied to all of them. The caller
// (`worksheetXml`, fed from `Workbook.activeTabIndex`) is what guarantees the "exactly one".
export function sheetViewsXml(view: SheetView, active: boolean): string {
  const selected = active ? ' tabSelected="1"' : '';
  // Excel defaults the grid on, so only an explicit `false` is worth an attribute; leaving it unset
  // keeps a sheet that never asked about gridlines byte-clean through a round-trip.
  const gridLines = view.showGridLines === false ? ' showGridLines="0"' : '';
  const xSplit = view.xSplit ?? 0;
  const ySplit = view.ySplit ?? 0;
  if (view.state !== 'frozen' || (xSplit === 0 && ySplit === 0)) {
    return `<sheetViews><sheetView${gridLines}${selected} workbookViewId="0"/></sheetViews>`;
  }
  const topLeftCell = view.topLeftCell ?? encodeAddress(xSplit + 1, ySplit + 1);
  const activePane =
    xSplit > 0 && ySplit > 0 ? 'bottomRight' : xSplit > 0 ? 'topRight' : 'bottomLeft';
  const pane =
    '<pane' +
    (xSplit > 0 ? ` xSplit="${xSplit}"` : '') +
    (ySplit > 0 ? ` ySplit="${ySplit}"` : '') +
    ` topLeftCell="${escapeAttr(topLeftCell)}" activePane="${activePane}" state="frozen"/>`;
  const selection = `<selection pane="${activePane}" activeCell="${escapeAttr(topLeftCell)}" sqref="${escapeAttr(topLeftCell)}"/>`;
  return `<sheetViews><sheetView${gridLines}${selected} workbookViewId="0">${pane}${selection}</sheetView></sheetViews>`;
}

// `<sheetPr>` carries the sheet's appearance properties: the tab colour, the outline
// summary-position flags, and the fit-to-page flag. It is the first child of `<worksheet>` in
// CT_Worksheet order; its own children follow CT_SheetPr order — `<tabColor>`, `<outlinePr>`, then
// `<pageSetUpPr>`. Omitted entirely when the sheet carries none, so an unadorned sheet stays
// byte-clean.
export function sheetPrXml(sheet: Worksheet): string {
  const children =
    (sheet.tabColor !== undefined ? `<tabColor ${colorAttrs(sheet.tabColor)}/>` : '') +
    outlinePrXml(sheet.outline) +
    pageSetUpPrXml(sheet.pageSetup);
  return children === '' ? '' : `<sheetPr>${children}</sheetPr>`;
}

// `<pageSetUpPr>` holds the fit-to-page toggle, which lives on the sheet properties rather than on
// `<pageSetup>` — Excel reads it from here to decide whether the `fitToWidth`/`fitToHeight` counts
// or the fixed `scale` govern printing. Emitted only when the author set the flag.
function pageSetUpPrXml(pageSetup: PageSetup): string {
  return pageSetup.fitToPage ? '<pageSetUpPr fitToPage="1"/>' : '';
}

// `<outlinePr>` carries only the summary-position flags today. Each is emitted solely when the
// caller set it, so an inverted placement (`summaryBelow="0"`) is honoured while an untouched sheet
// keeps the element out of the file entirely.
function outlinePrXml(outline: OutlineProperties): string {
  const attrs =
    boolAttr('summaryBelow', outline.summaryBelow) + boolAttr('summaryRight', outline.summaryRight);
  return attrs === '' ? '' : `<outlinePr${attrs}/>`;
}

// Each sheet-protection flag maps to a `<sheetProtection>` attribute whose value is INVERTED
// from the author-facing allow-flag: the attribute records that an operation is *forbidden*
// ("1"), so `allow: true` serialises as "0". Only a value that differs from OOXML's per-
// attribute default (see SHEET_PROTECTION_FLAGS) is written — most editing operations default
// to forbidden under protection, while selecting cells defaults to permitted.
//
// <sheetProtection> is what makes the per-cell locked/hidden flags bite. `sheet="1"` marks the
// sheet protected; the password credential (when present) guards lifting it; the flag attributes
// carve out the operations that stay available. base64 salt/hash use only XML-safe characters.
export function sheetProtectionXml(protection: SheetProtection | undefined): string {
  if (protection === undefined) return '';
  const {flags, credential} = protection;
  let attrs = '';
  if (credential !== undefined) {
    attrs +=
      ` algorithmName="${credential.algorithmName}"` +
      ` hashValue="${credential.hashValue}"` +
      ` saltValue="${credential.saltValue}"` +
      ` spinCount="${credential.spinCount}"`;
  }
  attrs += ' sheet="1"';
  for (const {key, defaultForbidden} of SHEET_PROTECTION_FLAGS) {
    const allow = flags[key];
    if (allow === undefined) continue;
    const forbidden = !allow;
    if (forbidden === defaultForbidden) continue;
    attrs += boolAttr(key, forbidden);
  }
  return `<sheetProtection${attrs}/>`;
}

// The sheet's autofilter: `<autoFilter ref="A1:C10"/>` when it only draws dropdowns, or with nested
// `<filterColumn>` children when columns carry criteria. Its companion `_FilterDatabase` defined name
// (the range Excel derives filtering from) is written in the workbook part, so a sheet with no filter
// emits nothing here and nothing there.
export function autoFilterXml(filter: AutoFilter | undefined): string {
  if (filter === undefined) return '';
  const ref = escapeAttr(filter.ref);
  if (filter.columns.length === 0) return `<autoFilter ref="${ref}"/>`;
  return `<autoFilter ref="${ref}">${filter.columns.map(filterColumnXml).join('')}</autoFilter>`;
}

function filterColumnXml(column: FilterColumn): string {
  return `<filterColumn colId="${column.colId}">${filterCriteriaXml(column.criteria)}</filterColumn>`;
}

// A values filter is `<filters>` with a `<filter val>` per allowed value (and `blank="1"` to admit
// empty cells); a custom filter is `<customFilters>` with one or two `<customFilter operator val>`
// predicates, `and="1"` when they are AND-combined rather than OR.
function filterCriteriaXml(criteria: FilterCriteria): string {
  if (criteria.kind === 'values') {
    const blankAttr = criteria.blank ? ' blank="1"' : '';
    const filters = criteria.values.map((value) => `<filter val="${escapeAttr(value)}"/>`).join('');
    return `<filters${blankAttr}>${filters}</filters>`;
  }
  const andAttr = criteria.and ? ' and="1"' : '';
  const predicates = criteria.predicates
    .map((p) => `<customFilter operator="${p.operator}" val="${escapeAttr(p.val)}"/>`)
    .join('');
  return `<customFilters${andAttr}>${predicates}</customFilters>`;
}

// CT_HeaderFooter child order, paired with the flag their presence gates: the even- and
// first-page variants are silently ignored by Excel unless differentOddEven / differentFirst
// are set, so the writer derives each flag from whether any variant in its class was provided.
const HF_CHILDREN = [
  {tag: 'oddHeader', key: 'oddHeader'},
  {tag: 'oddFooter', key: 'oddFooter'},
  {tag: 'evenHeader', key: 'evenHeader'},
  {tag: 'evenFooter', key: 'evenFooter'},
  {tag: 'firstHeader', key: 'firstHeader'},
  {tag: 'firstFooter', key: 'firstFooter'},
] as const;

export function headerFooterXml(hf: HeaderFooter): string {
  const children = HF_CHILDREN.filter(({key}) => hf[key] !== undefined);
  if (children.length === 0) return '';
  const differentOddEven = hf.evenHeader !== undefined || hf.evenFooter !== undefined;
  const differentFirst = hf.firstHeader !== undefined || hf.firstFooter !== undefined;
  let attrs = '';
  if (differentOddEven) attrs += ' differentOddEven="1"';
  if (differentFirst) attrs += ' differentFirst="1"';
  const body = children
    .map(({tag, key}) => `<${tag}>${escapeText(hf[key] as string)}</${tag}>`)
    .join('');
  return `<headerFooter${attrs}>${body}</headerFooter>`;
}

// Excel's "Normal" margins, in inches — the defaults Excel writes for an untouched sheet.
const DEFAULT_MARGINS = {
  left: 0.7,
  right: 0.7,
  top: 0.75,
  bottom: 0.75,
  header: 0.3,
  footer: 0.3,
} as const;
const MARGIN_SIDES = ['left', 'right', 'top', 'bottom', 'header', 'footer'] as const;

// `<printOptions>` carries the print-toggle flags and sits just before `<pageMargins>` in
// CT_Worksheet order. Each attribute is emitted only when the model carries it — as an explicit
// `="1"`/`="0"` so a caller can force a flag off against Excel's default — and an untouched sheet
// keeps the element out of the file entirely.
export function printOptionsXml(printOptions: PrintOptions): string {
  const attrs =
    boolAttr('horizontalCentered', printOptions.horizontalCentered) +
    boolAttr('verticalCentered', printOptions.verticalCentered) +
    boolAttr('headings', printOptions.headings) +
    boolAttr('gridLines', printOptions.gridLines) +
    boolAttr('gridLinesSet', printOptions.gridLinesSet);
  return attrs === '' ? '' : `<printOptions${attrs}/>`;
}

// OOXML's <pageMargins> is all-or-nothing: setting any one margin requires all six, or Excel
// repairs the file. So the element is emitted only when the caller set at least one, and the
// untouched sides fall back to the Normal-preset defaults.
export function pageMarginsXml(margins: PageMargins): string {
  if (MARGIN_SIDES.every((side) => margins[side] === undefined)) return '';
  const attrs = MARGIN_SIDES.map(
    (side) => `${side}="${numberText(margins[side] ?? DEFAULT_MARGINS[side])}"`,
  ).join(' ');
  return `<pageMargins ${attrs}/>`;
}

// `<pageSetup>` carries the print-scaling attributes (all but `fitToPage`, which is a `<sheetPr>`
// flag). It sits between `<pageMargins>` and `<headerFooter>` in CT_Worksheet order. Each attribute
// is emitted only when the author set it, so an untouched sheet keeps the element out of the file
// and a partially-set one never fabricates the counts Excel would otherwise default. A non-null
// `printerSettingsRelId` links the sheet's opaque printer-settings blob and forces the element out
// even when no scaling attribute is set — the reference is the only thing the model has to carry.
export function pageSetupXml(pageSetup: PageSetup, printerSettingsRelId: string | null): string {
  const attrs =
    attr('paperSize', pageSetup.paperSize) +
    attr('scale', pageSetup.scale) +
    attr('fitToWidth', pageSetup.fitToWidth) +
    attr('fitToHeight', pageSetup.fitToHeight) +
    (pageSetup.pageOrder !== undefined ? ` pageOrder="${pageSetup.pageOrder}"` : '') +
    (pageSetup.orientation !== undefined ? ` orientation="${pageSetup.orientation}"` : '') +
    (printerSettingsRelId !== null ? ` r:id="${printerSettingsRelId}"` : '');
  return attrs === '' ? '' : `<pageSetup${attrs}/>`;
}

// Manual page breaks (`<rowBreaks>`/`<colBreaks>`): one `<brk>` per row/column the layout splits
// before. Excel records both the running total (`count`) and the manual subset (`manualBreakCount`);
// every break the model carries is a manual, author-set one, so the two counts coincide. `max` bounds
// the break across the other axis (Excel writes the last row/column index); a break without one is
// emitted bare. Row and column breaks share this shape, differing only in the wrapping element.
export function pageBreaksXml(
  breaks: readonly PageBreak[],
  element: 'rowBreaks' | 'colBreaks',
): string {
  if (breaks.length === 0) return '';
  const brks = breaks
    .map((brk) => {
      const maxAttr = brk.max !== undefined ? ` max="${brk.max}"` : '';
      return `<brk id="${brk.id}"${maxAttr} man="1"/>`;
    })
    .join('');
  return `<${element} count="${breaks.length}" manualBreakCount="${breaks.length}">${brks}</${element}>`;
}
