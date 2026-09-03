// Type-level tests over the internal facet tables: the declarations that exist so a reader and a
// writer cannot fall out of step over the same set of fields.
//
// Each table pairs with an `AssertNever<Exclude<keyof T, table keys>>` alias beside it, which fails
// to resolve the moment the type gains a field the table does not name. That alias is already a
// compile error where it is declared; asserting it here is what states, in one place, which tables
// carry the guarantee, so a table added without one is visible as an absence rather than assumed.

import type {EveryColumnPropertyIsMirrored} from '../core/column.ts';
import type {
  EveryDifferentialStyleFieldIsCloned,
  EveryRuleFieldIsCloned,
} from '../core/conditional-formatting.ts';
import type {EveryDataValidationFieldIsCloned} from '../core/data-validation.ts';
import type {EveryAlignmentFacetIsDeclared} from '../core/style.ts';
import type {EveryTableStyleInfoFieldIsCloned} from '../core/table.ts';
import type {EveryWorksheetModelFieldHasAFacet} from '../core/worksheet-model.ts';
import type {Equal, Expect} from './expect.ts';

export type FacetTableExhaustiveness = [
  // `<alignment>`'s seven facets, read and written from one declaration.
  Expect<Equal<EveryAlignmentFacetIsDeclared, never>>,
  // A column's properties, mirrored between the handle and the stored record.
  Expect<Equal<EveryColumnPropertyIsMirrored, never>>,
  // Every worksheet model field, covered by a getter/setter facet.
  Expect<Equal<EveryWorksheetModelFieldHasAFacet, never>>,
  // The four defensive copies: a field added to one of these types without a copy strategy would
  // otherwise be carried by reference into the model, aliasing the caller's object silently.
  Expect<Equal<EveryRuleFieldIsCloned, never>>,
  Expect<Equal<EveryDifferentialStyleFieldIsCloned, never>>,
  Expect<Equal<EveryDataValidationFieldIsCloned, never>>,
  Expect<Equal<EveryTableStyleInfoFieldIsCloned, never>>,
];
