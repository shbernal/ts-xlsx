// The print/page-layout data shapes a worksheet carries: how it scales and orients on paper, which
// print toggles are set, where manual page breaks fall, the margins, and the header/footer text. Each
// is a pure data shape mapping onto an OOXML print element (`<pageSetup>`, `<printOptions>`, `<brk>`,
// `<pageMargins>`, `<headerFooter>`); the model stores only what an author or source file set, so an
// unset field is omitted and a round-trip never fabricates one.

/**
 * Print-scaling and orientation settings. These map onto two OOXML elements: `fitToPage` is the
 * `<pageSetUpPr>` flag (a `<sheetPr>` child) that switches Excel from fixed-zoom to fit-to-page
 * scaling, while the rest are `<pageSetup>` attributes. Excel honours `scale` only when `fitToPage`
 * is off and the `fitToWidth`/`fitToHeight` page counts only when it is on, but the model carries
 * whatever the author set — an unset field is omitted so a round-trip never fabricates one. An
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
  orientation?: 'portrait' | 'landscape';
  /** Order pages are numbered/printed in across a multi-page sheet. */
  pageOrder?: 'downThenOver' | 'overThenDown';
  /**
   * Paper size as Excel's 1-based enumeration index (e.g. `9` = A4, `1` = US Letter). Carried as an
   * opaque integer — the model does not map it to physical dimensions, only preserves whatever the
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
 * Print-toggle flags from the `<printOptions>` element. Each maps to a boolean OOXML attribute that
 * defaults false — except `gridLinesSet`, which defaults true and gates whether `gridLines` is
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

/**
 * A manual page break (`<brk>`). For a row break, `id` is the row the layout splits *before*; for a
 * column break it is the column. `max` bounds the break's extent across the other axis (Excel writes
 * the last row/column index) and `man` marks it author-set rather than automatic — the model preserves
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
