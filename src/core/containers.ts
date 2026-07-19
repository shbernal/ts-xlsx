// Identity-preserving replacement of a container's contents.
//
// Several model fields are `readonly` fields holding a mutable object or array: a caller (or a live
// getter) may hold a durable reference to the container, so importing a model must replace what the
// container *holds* without swapping the container itself. These two primitives do that — one for an
// object's keys, one for an array's elements — so the field's identity survives a wholesale reload.

// Replace an object's contents in place: clear every own key, then copy `source`'s keys over. Used for
// fields like a worksheet's `properties`/`pageSetup`/`headerFooter`, which are `readonly` fields holding
// mutable objects, so importing a model must overwrite them rather than reassign — and clear any stale
// key the incoming model does not carry. `Reflect` deletes each existing key without pretending the
// object carries a string index signature.
export function overwrite<T extends object>(target: T, source: T): void {
  for (const key of Reflect.ownKeys(target)) Reflect.deleteProperty(target, key);
  Object.assign(target, source);
}

// Replace an array's contents in place: clear it, then append `next` element by element. Appends
// individually rather than spreading `next` into a single `push(...next)` call, whose argument count the
// JS engine bounds by its call-stack limit — these arrays are filled from parsed files, so a hostile
// input must not be able to overflow that limit.
export function replaceContents<T>(array: T[], next: readonly T[]): void {
  array.length = 0;
  for (const item of next) array.push(item);
}
