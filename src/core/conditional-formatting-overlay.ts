// The conditional-formatting overlay a Worksheet owns: an insertion-ordered, defensively-copied list
// of range-bound rule sets. Kept as its own class, the sibling to {@link DataValidationOverlay}, so
// Worksheet delegates the collection's storage and cloning rather than managing the array itself.

import {type ConditionalFormatting, cloneConditionalFormatting} from './conditional-formatting.ts';
import {replaceContents} from './containers.ts';
import type {AxisSplice} from './grid-shift.ts';
import {shiftSqref} from './merge.ts';

export class ConditionalFormattingOverlay {
  readonly #entries: ConditionalFormatting[] = [];

  /**
   * Attach a conditional formatting to a target range. `formatting.ref` is an OOXML `sqref`: one
   * range (`"A1:A10"`), a whole column, or several space-separated areas (`"A1:C1 A3:C3"`) sharing one
   * rule set. The block is stored once against the range, defensively copied so the getter never hands
   * back a reference into the caller's object.
   */
  add(formatting: ConditionalFormatting): void {
    this.#entries.push(cloneConditionalFormatting(formatting));
  }

  /** The conditional formattings on this sheet, each bound to its target range, in insertion order. */
  get entries(): readonly ConditionalFormatting[] {
    return this.#entries;
  }

  /**
   * Re-anchor every rule set through a row or column splice, so a highlight keeps covering the cells
   * it was written for. A rule set whose every target area fell inside a deleted span goes with them.
   */
  shift(splice: AxisSplice): void {
    const entries: ConditionalFormatting[] = [];
    for (const entry of this.#entries) {
      const ref = shiftSqref(entry.ref, splice);
      if (ref !== undefined) entries.push({...entry, ref});
    }
    replaceContents(this.#entries, entries);
  }

  /** Drop every conditional formatting, leaving the overlay empty. */
  clear(): void {
    this.#entries.length = 0;
  }
}
