// The write side of the relationship graph: a single `<Relationship>` element and the `.rels` part
// envelope that wraps a set of them. Pure OPC: a `.rels` part has the same shape whatever the parts
// it wires together are spelled in, so this is stated once here rather than inside a codec.
//
// The reader's half of the same graph (resolving targets, walking a part closure) is in `read-opc.ts`.

import {escapeAttr, XML_DECLARATION} from '../../xml/xml.ts';
import {PKG_RELS_NS} from './namespaces.ts';

/**
 * A single `<Relationship>`. An `external` target lives outside the package (a hyperlink URL), so the
 * element carries `TargetMode="External"`; a package-internal target (the default) omits it.
 *
 * All three values are escaped here, unconditionally. The obligation used to sit in a comment saying
 * the caller escaped a target it did not control, which made the safety of this function a property
 * of its call sites rather than of itself, on the one boundary in the tree where `xml.ts` says that is
 * explicitly not how it is done. Escaping a writer-controlled package path is the identity, so the
 * generated chains emit the same bytes and the guarantee holds by construction instead of by
 * inspecting every caller. `TargetMode` stays a fixed token, not a value.
 */
export function relationship(
  id: string,
  type: string,
  target: string,
  options?: {external?: boolean},
): string {
  const mode = options?.external ? ' TargetMode="External"' : '';
  return (
    `<Relationship Id="${escapeAttr(id)}" Type="${escapeAttr(type)}" ` +
    `Target="${escapeAttr(target)}"${mode}/>`
  );
}

// Wrap a part's `<Relationship>` elements in the OPC `.rels` envelope (XML declaration + the namespaced
// `<Relationships>` root). Every `.rels` part the writer emits shares this envelope; only the elements
// inside differ, so each caller builds its own list of {@link relationship} strings and hands them here.
export function relationshipsPart(relationships: readonly string[]): string {
  return `${XML_DECLARATION}<Relationships xmlns="${PKG_RELS_NS}">${relationships.join('')}</Relationships>`;
}

// A whole `.rels` part from relationship records: the common case, where the caller has the three
// values rather than pre-rendered elements. `external` is optional and defaults to a package-internal
// target, which is what a generated part chain (pivot table → cache definition → cache records) wants
// and what a preserved foreign relationship carries explicitly.
export function relsPartXml(
  rels: readonly {id: string; type: string; target: string; external?: boolean}[],
): string {
  return relationshipsPart(
    rels.map((rel) =>
      relationship(rel.id, rel.type, rel.target, rel.external ? {external: true} : {}),
    ),
  );
}
