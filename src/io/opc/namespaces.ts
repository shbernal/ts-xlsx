/**
 * The namespace URIs the Open Packaging Conventions layer itself owns: the ones that describe a
 * *package* rather than the spreadsheet inside it. Every OOXML package carries these regardless of
 * which serialisation its office document uses, so an `.xlsb` and an `.xlsx` agree on them exactly.
 *
 * The SpreadsheetML vocabulary (the main namespace, DrawingML, the Microsoft extension URIs) is a
 * property of the XML codec and lives in `../xlsx/namespaces.ts`.
 */

import type {NamespaceScope} from '../../xml/xml-namespaces.ts';
import type {XmlAttributes} from '../../xml/xml-scan.ts';

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

/**
 * A relationship attribute (`r:id`, `r:embed`) resolved by namespace rather than by prefix.
 *
 * A workbook may bind {@link RELATIONSHIPS_NS} to any prefix it likes, and real toolchains do: reading
 * the literal `attrs['r:id']` lost every sheet relationship in such a file, so the sheets loaded
 * permanently empty with no error.
 *
 * Falls back to the conventional `r:` spelling when the prefix is not bound anywhere in scope. That is
 * not laxity for its own sake: a part carrying `r:embed` without declaring `r` is malformed XML that
 * this scanner does not reject, it was read before, and refusing it now would lose a file to a fix
 * meant to gain files. Resolution by namespace comes first, so a file that binds `r:` to something
 * else is read correctly rather than by its prefix.
 */
export function relAttr(
  scope: NamespaceScope,
  attrs: XmlAttributes,
  local: string,
): string | undefined {
  return scope.attr(attrs, RELATIONSHIPS_NS, local) ?? attrs[`r:${local}`];
}
