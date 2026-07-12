// Anchored images: the model for a picture pinned to a worksheet's grid.
//
// The image *bytes* live once on the workbook (a small media registry, addressed by a numeric id);
// a worksheet then anchors that image to a rectangle of cells. Storing the bytes centrally means the
// same picture used on two sheets — a logo in a header band, say — is one media part, not two.

/** A point in the drawing grid: a 0-based column and row, plus an EMU offset into that cell.
 * The offsets default to zero, pinning the point to the cell's top-left corner. */
export interface AnchorPoint {
  /** 0-based column index (column A is 0). */
  readonly col: number;
  /** 0-based row index (row 1 is 0). */
  readonly row: number;
  /** Horizontal offset into the cell, in EMUs (914400 per inch). Defaults to 0. */
  readonly colOff?: number;
  /** Vertical offset into the cell, in EMUs. Defaults to 0. */
  readonly rowOff?: number;
}

/** A two-cell anchor: the image's top-left (`from`) and bottom-right (`to`) grid points. The image
 * fills the rectangle between them and reflows as the intervening rows/columns resize. */
export interface ImageAnchor {
  readonly from: AnchorPoint;
  readonly to: AnchorPoint;
}

/** An image pinned to a worksheet: which workbook media it shows (`imageId`) and where. */
export interface AnchoredImage {
  /** Index into the workbook's media registry (the id {@link Workbook.addImage} returned). */
  readonly imageId: number;
  readonly anchor: ImageAnchor;
}

/** A picture's bytes and its file kind, as held in the workbook's media registry. */
export interface WorkbookImage {
  /** Lower-case file extension without a dot — `"png"`, `"jpeg"`, `"gif"`. Drives the media part's
   * name and content type. */
  readonly extension: string;
  readonly data: Uint8Array;
}
