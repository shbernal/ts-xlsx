// Defensive copies that the compiler proves are complete.
//
// Four overlay types are copied on the way into the model so a stored rule never aliases the caller's
// object, and all four were hand-rolled: a spread, then a hand-listed re-spread per nested field.
// That is the one shape in `core/` with no exhaustiveness proof behind it, and the failure it invites
// is quiet. Add an eighth nested field to a conditional-formatting rule and the spread carries it *by
// reference*: the stored entry then aliases the caller's object, a later mutation of theirs reaches
// into the sheet, and nothing fails to build. It is the merge-loss class of failure one indirection
// out, which is why every other copy loop here is list-driven and compiler-checked.
//
// So each type declares a plan: one entry per field saying how deep the copy of that field goes, with
// an `AssertNever` proof that the plan names every field the type has. A new field then does not
// compile until someone says what copying it means.

/**
 * How far a copy of one field goes.
 *
 * - `'value'`: assign it. Correct for a scalar, and for anything genuinely immutable.
 * - `'record'`: `{...v}`. A flat object whose own fields are scalars.
 * - `'values'`: `[...v]`. An array of scalars.
 * - `'records'`: `v.map(x => ({...x}))`. An array of flat objects.
 * - a function: the field's own clone, for a nested type that has a plan of its own.
 */
export type CloneStrategy<V> = 'value' | 'record' | 'values' | 'records' | ((value: V) => V);

/**
 * One strategy per field of `T`. The `Record` is what makes a missing field a compile error.
 *
 * A field's strategy sees the field's type without `undefined`, because {@link cloneWith} skips an
 * absent field rather than handing it over. Under `exactOptionalPropertyTypes` a facet declared
 * `font?: Font | undefined` keeps that `undefined` through `Required`, so without the exclusion a
 * nested type's own clone function would not satisfy its own field's strategy.
 */
export type ClonePlan<T> = {
  readonly [K in keyof Required<T>]-?: CloneStrategy<Exclude<Required<T>[K], undefined>>;
};

/**
 * Copy `source` according to `plan`, leaving a field absent when the source leaves it absent.
 *
 * Absent rather than `undefined`: every one of these types declares its optional fields under
 * `exactOptionalPropertyTypes`, where a present-but-undefined key is a different shape from a missing
 * one, and a round-trip that fabricated the first would report formatting a cell does not have.
 */
export function cloneWith<T extends object>(source: T, plan: ClonePlan<T>): T {
  const out: Record<string, unknown> = {};
  for (const key of Object.keys(plan) as (keyof T & string)[]) {
    const value = source[key];
    if (value === undefined) continue;
    out[key] = applyStrategy(value, plan[key] as CloneStrategy<unknown>);
  }
  return out as T;
}

function applyStrategy(value: unknown, strategy: CloneStrategy<unknown>): unknown {
  if (typeof strategy === 'function') return strategy(value);
  switch (strategy) {
    case 'value':
      return value;
    case 'record':
      return {...(value as object)};
    case 'values':
      return [...(value as readonly unknown[])];
    case 'records':
      return (value as readonly object[]).map((entry) => ({...entry}));
  }
}

// Each plan pairs with an `AssertNever<Exclude<keyof Required<T>, keyof typeof PLAN>>` alias beside
// it, the same proof `EveryColumnPropertyIsMirrored` and `EveryWorksheetModelFieldHasAFacet` carry.
// It has to be instantiated where the plan is: a generic `EveryFieldPlanned<T, Plan>` cannot resolve
// the constraint for an unresolved `T`, so the alias would prove nothing until it was used.
