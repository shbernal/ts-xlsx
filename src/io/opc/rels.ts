// The write side of the relationship graph: a single `<Relationship>` element and the `.rels` part
// envelope that wraps a set of them. Pure OPC — a `.rels` part has the same shape whatever the parts
// it wires together are spelled in, so this is stated once here rather than inside a codec.
//
// The reader's half of the same graph (resolving targets, walking a part closure) is in `read-opc.ts`.

import {escapeAttr, XML_DECLARATION} from '../../xml/xml.ts';
import {PKG_RELS_NS} from './namespaces.ts';

// A single `<Relationship>`. An `external` target lives outside the package (a hyperlink URL), so the
// element carries `TargetMode="External"`; a package-internal target (the default) omits it. The caller
// escapes the target when it is not a writer-controlled package path.
export function relationship(
  id: string,
  type: string,
  target: string,
  options?: {external?: boolean},
): string {
  const mode = options?.external ? ' TargetMode="External"' : '';
  return `<Relationship Id="${id}" Type="${type}" Target="${target}"${mode}/>`;
}

// Wrap a part's `<Relationship>` elements in the OPC `.rels` envelope (XML declaration + the namespaced
// `<Relationships>` root). Every `.rels` part the writer emits shares this envelope; only the elements
// inside differ, so each caller builds its own list of {@link relationship} strings and hands them here.
export function relationshipsPart(relationships: readonly string[]): string {
  return `${XML_DECLARATION}<Relationships xmlns="${PKG_RELS_NS}">${relationships.join('')}</Relationships>`;
}

export function preservedRelsXml(
  rels: readonly {id: string; type: string; target: string; external?: boolean}[],
): string {
  return relationshipsPart(
    rels.map((rel) =>
      relationship(rel.id, rel.type, escapeAttr(rel.target), rel.external ? {external: true} : {}),
    ),
  );
}

// A `.rels` part for a generated part chain (pivot table → cache definition → cache records). Targets
// are writer-controlled package paths, so no attribute escaping is needed.
export function relsPartXml(rels: readonly {id: string; type: string; target: string}[]): string {
  return relationshipsPart(rels.map((rel) => relationship(rel.id, rel.type, rel.target)));
}
