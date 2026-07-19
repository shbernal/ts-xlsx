// Compile-time assertions for the public API's type surface.
//
// These types have no runtime footprint: they are verified by `npm run typecheck`
// (tsc over the whole src/ tree) and excluded from the published build. When a
// public type drifts, the matching `Expect<Equal<...>>` stops resolving to `true`
// and the typecheck gate fails — the type-level analogue of a red test.

/**
 * Bivariance-safe structural equality: `true` only when `A` and `B` are mutually
 * assignable *and* identically shaped. The two-function-parameter trick is what
 * distinguishes `{a: string}` from `{a: string} | {a: string; b: number}` that a
 * plain `A extends B ? ... : ...` would wave through.
 */
export type Equal<A, B> =
  (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;

/** Resolves only when its argument is exactly `true`; the workhorse of a case list. */
export type Expect<T extends true> = T;

/** `true` when `A` is assignable to `B` — a directional (subtype) check. */
export type Extends<A, B> = A extends B ? true : false;
