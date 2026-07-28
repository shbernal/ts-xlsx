/**
 * The namespace URIs the Open Packaging Conventions layer itself owns — the ones that describe a
 * *package* rather than the spreadsheet inside it. Every OOXML package carries these regardless of
 * which serialisation its office document uses, so an `.xlsb` and an `.xlsx` agree on them exactly.
 *
 * The SpreadsheetML vocabulary (the main namespace, DrawingML, the Microsoft extension URIs) is a
 * property of the XML codec and lives in `../xlsx/namespaces.ts`.
 */

/** The `xmlns` of every `.rels` package relationships part. */
export const PKG_RELS_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

/** The `xmlns` of the package's `[Content_Types].xml`. */
export const CONTENT_TYPES_NS = 'http://schemas.openxmlformats.org/package/2006/content-types';

/**
 * Scopes both the relationship *type* URIs a `.rels` part declares and the `r:id` references a part's
 * body carries. Part-level rather than package-level in ECMA-376's own split, but it is the vocabulary
 * the relationship graph is written in, so it belongs with the graph and not with either codec.
 */
export const RELATIONSHIPS_NS =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
