// Identity-preserving replacement of a container's contents.
//
// Several model fields are `readonly` fields holding a mutable object or array: a caller (or a live
// getter) may hold a durable reference to the container, so importing a model must replace what the
// container *holds* without swapping the container itself. These two primitives do that, one for an
// object's keys and one for an array's elements, so the field's identity survives a wholesale reload.

// Replace an object's contents in place: clear every own key, then copy `source`'s keys over. Used for
// fields like a worksheet's `properties`/`pageSetup`/`headerFooter`, which are `readonly` fields holding
// mutable objects, so importing a model must overwrite them rather than reassign, and clear any stale
// key the incoming model does not carry. `Reflect` deletes each existing key without pretending the
// object carries a string index signature.
export function overwrite<T extends object>(target: T, source: T): void {
  for (const key of Reflect.ownKeys(target)) Reflect.deleteProperty(target, key);
  Object.assign(target, source);
}

// Replace an array's contents in place: clear it, then append `next` element by element. Appends
// individually rather than spreading `next` into a single `push(...next)` call, whose argument count the
// JS engine bounds by its call-stack limit. These arrays are filled from parsed files, so a hostile
// input must not be able to overflow that limit.
export function replaceContents<T>(array: T[], next: readonly T[]): void {
  array.length = 0;
  for (const item of next) array.push(item);
}

/**
 * Copy one key from `source` onto `target`, leaving `target` untouched when `source` omits it.
 *
 * One key at a time is the whole point, and the reason is a limit of the checker rather than a
 * preference: over a union key the compiler cannot correlate `source[key]`'s type with `target[key]`'s,
 * so a version taking the union needs a cast, and that cast is the one place a facet could be written
 * into the wrong slot with nothing to notice. Bound to a single member here, the assignment is checked.
 *
 * Generic over the object as well as the key because these two lines had been written out once for a
 * cell's formatting and once for a cell's content, each under its own paragraph making this point.
 */
export function copyKeyIfPresent<T extends object, K extends keyof T>(
  target: T,
  source: Readonly<T>,
  key: K,
): void {
  const value = source[key];
  if (value !== undefined) target[key] = value;
}
