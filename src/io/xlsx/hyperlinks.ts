// Cell hyperlinks — the sheet-level `<hyperlinks>` element, its external relationships, and the
// reader that folds a link back onto its cell's value.
//
// A hyperlink is not stored inside the cell in OOXML: the `<c>` holds only the visible label (a
// normal string value), while a separate `<hyperlink>` child of `<worksheet>` ties an A1 reference
// to a destination. An EXTERNAL destination (a URL) is reached indirectly, through a sheet
// relationship carrying `TargetMode="External"` that the `<hyperlink>` names by `r:id`. An INTERNAL
// destination (a location inside the same workbook, which the author writes as a `#`-prefixed value)
// is held directly in a `location` attribute with NO relationship — emitting an internal target as an
// external relationship makes a strict consumer resolve both the rel and the location and render the
// destination doubled.

import {decodeRange} from '../../core/address.ts';
import {type HyperlinkValue, isHyperlinkValue, isRichTextValue} from '../../core/value.ts';
import type {Worksheet} from '../../core/worksheet.ts';
import {localName, parseXml} from './xml-read.ts';
import {escapeAttr} from './xml.ts';

/** A hyperlink gathered from a sheet for serialisation: the cell it sits on, its target, and an
 * optional tooltip. The visible label is the cell's own value and is serialised as that value. */
export interface CollectedHyperlink {
  readonly ref: string;
  readonly target: string;
  readonly tooltip?: string;
}

/** A hyperlink resolved for serialisation. An external target carries a `relId` (the sheet
 * relationship holding the URL) plus that `target`; an internal target carries a `location` (the
 * in-workbook reference). Exactly one of `relId`/`location` is ever set. */
export interface PlannedHyperlink {
  readonly ref: string;
  readonly relId?: string;
  readonly target?: string;
  readonly location?: string;
  readonly tooltip?: string;
}

/** Gather every hyperlink cell on a sheet, in row-major order. */
export function collectHyperlinks(sheet: Worksheet): CollectedHyperlink[] {
  const links: CollectedHyperlink[] = [];
  for (const {cells} of sheet.rows()) {
    for (const cell of cells) {
      const value = cell.value;
      if (isHyperlinkValue(value)) {
        links.push({
          ref: cell.address,
          target: value.hyperlink,
          ...(value.tooltip !== undefined ? {tooltip: value.tooltip} : {}),
        });
      }
    }
  }
  return links;
}

/** Split collected links into internal (location, no rel) and external (relationship) forms,
 * numbering external relationship ids from `relIdBase + 1` so they follow every other sheet-local
 * relationship and adding one never renumbers an id already threaded into the sheet XML. */
export function planHyperlinks(
  links: readonly CollectedHyperlink[],
  relIdBase: number
): PlannedHyperlink[] {
  let external = 0;
  return links.map((link) => {
    const tooltip = link.tooltip !== undefined ? {tooltip: link.tooltip} : {};
    // A '#'-prefixed target is an internal document location: held verbatim in `location`, with no
    // relationship. Everything else is an external URL reached through a relationship.
    if (link.target.startsWith('#')) {
      return {ref: link.ref, location: link.target.slice(1), ...tooltip};
    }
    external += 1;
    return {ref: link.ref, relId: `rId${relIdBase + external}`, target: link.target, ...tooltip};
  });
}

/** The `<hyperlinks>` element, or '' when the sheet has none. Attribute order follows CT_Hyperlink:
 * `ref`, `r:id`, `location`, `tooltip`. */
export function hyperlinksXml(links: readonly PlannedHyperlink[]): string {
  if (links.length === 0) return '';
  const items = links
    .map((link) => {
      const rid = link.relId !== undefined ? ` r:id="${link.relId}"` : '';
      const location = link.location !== undefined ? ` location="${escapeAttr(link.location)}"` : '';
      const tooltip = link.tooltip !== undefined ? ` tooltip="${escapeAttr(link.tooltip)}"` : '';
      return `<hyperlink ref="${escapeAttr(link.ref)}"${rid}${location}${tooltip}/>`;
    })
    .join('');
  return `<hyperlinks>${items}</hyperlinks>`;
}

/** A hyperlink parsed from a sheet: its cell reference plus whichever of `rid`/`location`/`tooltip`
 * the `<hyperlink>` element carried. */
interface ParsedHyperlink {
  readonly ref: string;
  readonly rid?: string;
  readonly location?: string;
  readonly tooltip?: string;
}

/** Parse every `<hyperlink>` element out of a worksheet part. */
export function parseSheetHyperlinks(xml: string): ParsedHyperlink[] {
  const links: ParsedHyperlink[] = [];
  parseXml(xml, {
    onOpen(name, attrs) {
      if (localName(name) !== 'hyperlink') return;
      const ref = attrs.ref;
      if (ref === undefined) return;
      links.push({
        ref,
        ...(attrs['r:id'] !== undefined ? {rid: attrs['r:id']} : {}),
        ...(attrs.location !== undefined ? {location: attrs.location} : {}),
        ...(attrs.tooltip !== undefined ? {tooltip: attrs.tooltip} : {}),
      });
    },
    onText() {},
    onClose() {},
  });
  return links;
}

/** Fold parsed hyperlinks onto a sheet's cells, wrapping each cell's existing value (its visible
 * label) into a {@link HyperlinkValue}. `rels` maps a relationship id to its target URL. */
export function applyHyperlinks(
  sheet: Worksheet,
  links: readonly ParsedHyperlink[],
  rels: Map<string, string>
): void {
  for (const link of links) {
    const target = resolveTarget(link, rels);
    if (target === undefined) continue;
    // A hyperlink may span a range (`ref="D1:H1"`); Excel anchors the link at the range's top-left
    // cell. Decode to that anchor so a multi-cell link folds onto one cell instead of asking the
    // sheet for a range address it cannot resolve. A ref that does not decode is skipped, not fatal.
    const anchor = hyperlinkAnchor(link.ref);
    if (anchor === undefined) continue;
    const cell = sheet.getCell(anchor);
    // The visible label is the cell's own value: a plain string, or rich text when the label
    // carried per-run formatting. Any other value kind has no textual label, so it reads as empty.
    const cellValue = cell.value;
    const text =
      typeof cellValue === 'string' ? cellValue : isRichTextValue(cellValue) ? cellValue : '';
    const value: HyperlinkValue = {
      hyperlink: target,
      text,
      ...(link.tooltip !== undefined ? {tooltip: link.tooltip} : {}),
    };
    cell.value = value;
  }
}

// The single cell a hyperlink is anchored at: its `ref` verbatim when it names one cell, or the
// top-left corner when it spans a range (`D1:H1` → `D1`). Returns undefined for a ref that does not
// decode, so a malformed hyperlink is dropped rather than crashing the load.
function hyperlinkAnchor(ref: string): string | undefined {
  try {
    return decodeRange(ref).tl.address;
  } catch {
    return undefined;
  }
}

function resolveTarget(link: ParsedHyperlink, rels: Map<string, string>): string | undefined {
  if (link.rid !== undefined) {
    const base = rels.get(link.rid);
    if (base === undefined) return undefined;
    // A foreign file may split an external URL's fragment into the `location` attribute, apart from
    // the relationship Target; rejoin them so the whole URL survives. Our own writer keeps the
    // fragment in the Target, so a link we wrote never carries both.
    return link.location !== undefined ? `${base}#${link.location}` : base;
  }
  // No relationship: an internal ('#'-prefixed) target held verbatim in `location`.
  return link.location !== undefined ? `#${link.location}` : undefined;
}
