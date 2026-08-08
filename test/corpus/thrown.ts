// Turning a caught value into the string a case asserts on.
//
// A `catch` binding is `unknown`, and the corpus has one settled answer for what to report from it:
// the thrown value's `message` if it has one, the value itself otherwise. That answer was written out
// by hand 46 times as `String((e as any)?.message || e)`, where the cast was load-bearing — the only
// way to reach `.message` off `unknown` without narrowing. Naming the operation removes every one of
// those casts, because the widening happens once, here, behind a signature that says `unknown` in and
// `string` out.
//
// It is deliberately lenient about what it was handed. An adapter method reports a *failure* as data
// for a case to assert on, so it must survive being given a string, a null, or a plain object as
// readily as an `Error`; narrowing with `instanceof Error` and giving up on the rest would silently
// turn a non-`Error` throw into a uselessly generic report, which is the opposite of what a
// regression case needs from it.

/**
 * The message a caught value reports, or the value stringified when it carries none.
 *
 * `messageOf(new Error('x'))` → `'x'`; `messageOf('x')` → `'x'`; `messageOf(null)` → `'null'`.
 * An empty `message` falls through to the value itself, so a throw of `{message: ''}` is reported as
 * `'[object Object]'` rather than as no error at all.
 */
export const messageOf = (thrown: unknown): string =>
  String((thrown as {message?: unknown} | null | undefined)?.message || thrown);

/**
 * As {@link messageOf}, but preferring a system `code` when the value carries one.
 *
 * For failures that originate in the OS rather than in the library: `ENOENT` is the assertable fact
 * about a bad destination, and the `message` wrapped around it varies by platform and Node version.
 */
export const codeOrMessageOf = (thrown: unknown): string => {
  const carrier = thrown as {code?: unknown; message?: unknown} | null | undefined;
  return String(carrier?.code || carrier?.message || thrown);
};
