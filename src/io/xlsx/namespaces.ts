/**
 * Canonical OOXML namespace URIs and well-known extension GUIDs.
 *
 * These are wire-format constants: Excel keys its parsing off the exact URI or
 * GUID, so a producer must reproduce each one byte-for-byte. Centralizing them
 * keeps the writer and reader from drifting apart and retires the
 * `NS_MAIN`/`MAIN_NS`/`main` naming fork that had grown across the io modules.
 */

/**
 * SpreadsheetML main namespace — the default `xmlns` of the workbook,
 * worksheet, styles, sharedStrings, comments, table and pivot parts.
 */
export const SPREADSHEETML_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';

/** Scopes `r:id` relationship references carried inside a part's body. */
export const RELATIONSHIPS_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

/** The `xmlns` of every `.rels` package relationships part. */
export const PKG_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** DrawingML shared graphics namespace (`a:`). */
export const DRAWINGML_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** Spreadsheet-drawing anchor namespace (`xdr:`) used by the worksheet drawing part. */
export const XDR_NS = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';

/**
 * The 2009 Microsoft extension namespace. `x14` scopes the feature elements
 * Excel tucks inside `<ext>` blocks (conditional formatting, data validation,
 * slicers); it is declared inline on those elements exactly as Excel writes
 * them, so a worksheet root never needs an extra namespace declaration.
 */
export const X14_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';

/** Scopes the `<xm:sqref>`/`<xm:f>` references the x14 feature elements carry. */
export const XM_NS = 'http://schemas.microsoft.com/office/excel/2006/main';

/**
 * Well-known `<ext uri=…>` GUIDs. Each `<ext>` block is opaque to a consumer
 * that does not recognize its GUID, so a producer must emit these exact values
 * for Excel to rediscover the feature.
 */
export const CF_EXT_URI = '{78C0D931-6437-407d-A8EE-F0AAD7539E65}';
export const DATABAR_LINK_EXT_URI = '{B025F937-C7B1-47D3-B67F-A62EFF666E3E}';
export const DATA_VALIDATION_EXT_URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}';
export const SLICER_LIST_EXT_URI = '{A8765BA9-456A-4dab-B4F3-ACF838C121DE}';
export const SLICER_CACHES_EXT_URI = '{BBE1A952-AA13-448e-AADC-164F8A28A991}';
