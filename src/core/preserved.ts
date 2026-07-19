// Package content the model does not interpret, captured verbatim so a round-trip re-emits it intact.
// A preserved reference names a worksheet- or workbook-level relationship (a chart, vector drawing,
// slicer, pivot table) the model does not model; its transitive closure of parts is held as raw bytes
// with their content types and rewired relationships, and re-emitted unchanged on write.

/**
 * One outbound relationship of a {@link PreservedPart}: the id it carries inside its own rels part,
 * the relationship Type URI, and the resolved package path of the (internal) target part. Only
 * package-internal relationships are preserved — an external target (a URL) is not part of the
 * closure and is dropped.
 */
export interface PreservedRelationship {
  readonly id: string;
  readonly type: string;
  readonly targetPath: string;
}

/**
 * A package part the model does not interpret, captured verbatim so a round-trip re-emits it intact.
 * `bytes` are the raw part contents, `contentType` how the source package declared it, and `rels` its
 * outbound relationships (empty when the part references nothing). The writer re-numbers the part to a
 * fresh, collision-proof path and rewires `rels` accordingly, but never touches `bytes`.
 */
export interface PreservedPart {
  readonly path: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
  readonly rels: readonly PreservedRelationship[];
}

/**
 * A worksheet-level reference to package content the model does not model — preserved verbatim across
 * a round-trip instead of being silently dropped. `element` is the worksheet child that wires the
 * reference (`<drawing>` for a vector-shape drawing, `<legacyDrawingHF>` for a header/footer image),
 * or `null` when the sheet wires it by relationship alone (a pivot table or slicer Excel discovers by
 * scanning the sheet's rels, with no worksheet child pointing at it). `relType` is the relationship
 * Type URI to re-emit; `entryPath` is the part it points at; `parts` is the transitive closure of
 * parts that reference reaches (the entry included), each re-emitted with its relationships rewired.
 */
export interface PreservedWorksheetReference {
  readonly element: 'drawing' | 'legacyDrawingHF' | null;
  readonly relType: string;
  readonly entryPath: string;
  readonly parts: readonly PreservedPart[];
}
