// The print/page-layout data shapes a worksheet carries: how it scales and orients on paper, which
// print toggles are set, where manual page breaks fall, the margins, and the header/footer text. Each
// is a pure data shape mapping onto an OOXML print element (`<pageSetup>`, `<printOptions>`, `<brk>`,
// `<pageMargins>`, `<headerFooter>`); the model stores only what an author or source file set, so an
// unset field is omitted and a round-trip never fabricates one.

import {tokenSet} from '../token-set.ts';
import type {AssertNever} from './internal.ts';

/**
 * Paper orientation, as `<pageSetup orientation>` carries it.
 *
 * `ST_Orientation` has a third member, `default`, which means "whatever the printer decides" and is
 * indistinguishable from the attribute being absent. The model spells that absence as an unset field,
 * so a file carrying `default` reads back with no orientation and writes back without the attribute.
 */
export type PageOrientation = 'portrait' | 'landscape';

/** Narrow a raw `<pageSetup orientation>` token to a known {@link PageOrientation}. */
export const isPageOrientation = tokenSet<PageOrientation>({portrait: true, landscape: true});

/** The order pages are numbered and printed in across a sheet wider and taller than one page. */
export type PageOrder = 'downThenOver' | 'overThenDown';

/** Narrow a raw `<pageSetup pageOrder>` token to a known {@link PageOrder}. */
export const isPageOrder = tokenSet<PageOrder>({downThenOver: true, overThenDown: true});

/**
 * Print-scaling and orientation settings. These map onto two OOXML elements: `fitToPage` is the
 * `<pageSetUpPr>` flag (a `<sheetPr>` child) that switches Excel from fixed-zoom to fit-to-page
 * scaling, while the rest are `<pageSetup>` attributes. Excel honours `scale` only when `fitToPage`
 * is off and the `fitToWidth`/`fitToHeight` page counts only when it is on, but the model carries
 * whatever the author set: an unset field is omitted so a round-trip never fabricates one. An
 * empty object emits neither element.
 */
export interface PageSetup {
  /** Switch to fit-to-page scaling. Emitted as `<pageSetUpPr fitToPage="1">`. */
  fitToPage?: boolean;
  /** Pages wide to fit onto; `0` means "unbounded" (fit only by height). */
  fitToWidth?: number;
  /** Pages tall to fit onto; `0` means "unbounded" (fit only by width). */
  fitToHeight?: number;
  /** Fixed print zoom as a percentage; Excel honours it only when `fitToPage` is off. */
  scale?: number;
  /** Paper orientation. */
  orientation?: PageOrientation;
  /** Order pages are numbered/printed in across a multi-page sheet. */
  pageOrder?: PageOrder;
  /**
   * Paper size as Excel's 1-based enumeration index (e.g. `9` = A4, `1` = US Letter). Carried as an
   * opaque integer: the model does not map it to physical dimensions, only preserves whatever the
   * author or source file set.
   */
  paperSize?: number;
  /**
   * The printer-settings blob a source file bound to this sheet's `<pageSetup>` via an `r:id`
   * relationship, held verbatim. Excel stores the platform-specific `DEVMODE` (paper tray, duplex,
   * DPI, …) in this opaque binary part; the model does not interpret it, only round-trips the exact
   * bytes so re-writing a file that carried one does not silently drop the user's print configuration.
   */
  printerSettings?: Uint8Array;
}

/**
 * How one `<pageSetup>` attribute encodes: which model key it is and what kind of value it carries.
 *
 * Format-blind on purpose, the same way {@link AlignmentFacet} is: `PageSetup` is a core type and
 * the layering gate forbids core importing a serialisation, so the table states what an attribute
 * *is* and each codec supplies the reading and the writing off the `kind`. The OOXML attribute name
 * is the model key throughout, so it is not restated.
 */
export type PageSetupFacet =
  | {
      readonly key: 'paperSize' | 'scale' | 'fitToWidth' | 'fitToHeight';
      /** A non-negative integer: a page count, a percentage, or a paper-size id. */
      readonly kind: 'count';
    }
  | {
      readonly key: 'pageOrder' | 'orientation';
      readonly kind: 'token';
      /** The enumeration guard, and what to call it in the error when a value fails it. */
      readonly isValid: (value: string) => boolean;
      readonly label: string;
    };

/**
 * The six `<pageSetup>` attributes, declared once, in CT_PageSetup order. Both directions key off
 * this list, so an attribute written but not read (it survives a re-write and vanishes on load) or
 * read but not written is a compile error rather than something a reviewer has to notice.
 */
export const PAGE_SETUP_FACETS = [
  {key: 'paperSize', kind: 'count'},
  {key: 'scale', kind: 'count'},
  {key: 'fitToWidth', kind: 'count'},
  {key: 'fitToHeight', kind: 'count'},
  {key: 'pageOrder', kind: 'token', isValid: isPageOrder, label: 'page order'},
  {key: 'orientation', kind: 'token', isValid: isPageOrientation, label: 'page orientation'},
] as const satisfies readonly PageSetupFacet[];

/**
 * Compile-time proof that {@link PAGE_SETUP_FACETS} covers every `<pageSetup>` attribute.
 *
 * `fitToPage` and `printerSettings` are excluded because neither is one: `fitToPage` is a
 * `<sheetPr>` child's flag and `printerSettings` is the blob behind an `r:id`, so both are written
 * and read somewhere else entirely and a table entry for them would describe nothing.
 */
export type EveryPageSetupFacetIsDeclared = AssertNever<
  Exclude<
    keyof PageSetup,
    'fitToPage' | 'printerSettings' | (typeof PAGE_SETUP_FACETS)[number]['key']
  >
>;

/**
 * Print-toggle flags from the `<printOptions>` element. Each maps to a boolean OOXML attribute that
 * defaults false, except `gridLinesSet`, which defaults true and gates whether `gridLines` is
 * honoured. The model stores only what the source or caller set, so an unset flag is omitted and a
 * round-trip never fabricates one; an empty object emits no element at all.
 */
export interface PrintOptions {
  /** Centre the printed content horizontally on the page. */
  horizontalCentered?: boolean;
  /** Centre the printed content vertically on the page. */
  verticalCentered?: boolean;
  /** Print the row and column headings (the `1,2,3…` / `A,B,C…` gutters). */
  headings?: boolean;
  /** Print the cell gridlines. */
  gridLines?: boolean;
  /** Whether the `gridLines` flag is authoritative; when `false`, Excel ignores `gridLines`. */
  gridLinesSet?: boolean;
}

/** The `<printOptions>` flags, in CT_PrintOptions attribute order. Each is a plain OOXML boolean, so
 * the list is the whole of what either direction needs to know. */
export const PRINT_OPTION_FLAGS = [
  'horizontalCentered',
  'verticalCentered',
  'headings',
  'gridLines',
  'gridLinesSet',
] as const satisfies readonly (keyof PrintOptions)[];

/** Compile-time proof that {@link PRINT_OPTION_FLAGS} covers every {@link PrintOptions} flag. */
export type EveryPrintOptionFlagIsDeclared = AssertNever<
  Exclude<keyof PrintOptions, (typeof PRINT_OPTION_FLAGS)[number]>
>;

/**
 * A manual page break (`<brk>`). For a row break, `id` is the row the layout splits *before*; for a
 * column break it is the column. `max` bounds the break's extent across the other axis (Excel writes
 * the last row/column index) and `man` marks it author-set rather than automatic. The model preserves
 * whatever the source carried so a round-trip reproduces the break's span exactly.
 */
export interface PageBreak {
  /** The row (or column) the break precedes. */
  readonly id: number;
  /** The break's far extent across the other axis, if the source declared one. */
  readonly max?: number;
  /** `true` when the break is manual (author-set); Excel-authored breaks always are. */
  readonly man?: boolean;
}

/**
 * Print margins, in inches. OOXML's `<pageMargins>` requires all six to be present, but
 * the model stores only what the caller set; the writer fills the untouched ones with
 * valid defaults. An empty object means the element is omitted entirely.
 */
export interface PageMargins {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

/** The `<pageMargins>` sides, in the order CT_PageMargins declares them. */
export const MARGIN_SIDES = [
  'left',
  'right',
  'top',
  'bottom',
  'header',
  'footer',
] as const satisfies readonly (keyof PageMargins)[];

/** Compile-time proof that {@link MARGIN_SIDES} covers every {@link PageMargins} side. */
export type EveryMarginSideIsDeclared = AssertNever<
  Exclude<keyof PageMargins, (typeof MARGIN_SIDES)[number]>
>;

/**
 * Page header/footer text, one string per page class. Excel only honours the even- and
 * first-page variants when the writer also sets the gating flags (`differentOddEven`,
 * `differentFirst`); the writer derives those from which variants are present. An empty
 * object means the element is omitted entirely.
 */
export interface HeaderFooter {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
}

/** The `<headerFooter>` children, in CT_HeaderFooter child order. The element name is the model key
 * throughout, so one list serves the reader's capture, the reader's commit, and the writer. */
export const HEADER_FOOTER_ELEMENTS = [
  'oddHeader',
  'oddFooter',
  'evenHeader',
  'evenFooter',
  'firstHeader',
  'firstFooter',
] as const satisfies readonly (keyof HeaderFooter)[];

/** Compile-time proof that {@link HEADER_FOOTER_ELEMENTS} covers every {@link HeaderFooter} slot. */
export type EveryHeaderFooterElementIsDeclared = AssertNever<
  Exclude<keyof HeaderFooter, (typeof HEADER_FOOTER_ELEMENTS)[number]>
>;
