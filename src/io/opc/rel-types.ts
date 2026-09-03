// What a relationship Type says, asked once.
//
// Its own module, importing nothing, because both directions ask it. The reader resolves a part
// through its `.rels`; the writer decides which of a preserved sheet's relationships it is re-emitting.
// Put beside the reader, this one-line test dragged the whole read path -- the bounded inflater
// included -- into the writer's module closure, which the per-entry size budget is there to notice.

/**
 * Does a relationship Type name this class of part?
 *
 * A Type is a URI whose last segment names the part class, and the namespace in front of it varies by
 * relationship family and by the schema version a producer wrote against. So the match is on that
 * last segment.
 *
 * **The separator is the load-bearing character.** `/slicer` and `slicer` are not the same test:
 * without the slash, `slicer` matches `slicerCache`, `table` matches `pivotTable`, and a lookup
 * quietly gathers a part class nobody asked for. Every caller used to spell the slash itself, ten
 * times over, with nowhere to say why it was there.
 */
export function isRelType(type: string, name: string): boolean {
  return type.endsWith(`/${name}`);
}

/** {@link isRelType} against several part classes at once, for the predicates that partition
 * relationships into "modeled", "preserved verbatim", and "dropped". */
export function isAnyRelType(type: string, ...names: readonly string[]): boolean {
  return names.some((name) => isRelType(type, name));
}
