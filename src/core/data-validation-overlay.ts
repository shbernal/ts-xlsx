// The data-validation overlay a Worksheet owns: an insertion-ordered list of range-bound rules, plus
// the decoded rectangles a point-in-rect lookup ({@link DataValidationOverlay.at}) tests against. Kept
// as its own class — the sibling to {@link GridEdits} that owns splice arithmetic — rather than inline
// on Worksheet, since a validation's storage (a rule plus its decoded ranges) is a self-contained unit
// Worksheet only ever adds to, reads, or clears wholesale.

import {
  cloneDataValidation,
  type DataValidation,
  type DataValidationEntry,
} from './data-validation.ts';
import {decodeSqrefRects, type MergeRect} from './merge.ts';

export class DataValidationOverlay {
  readonly #entries: DataValidationEntry[] = [];
  readonly #rects: {rects: readonly MergeRect[]; rule: DataValidation}[] = [];

  /**
   * Attach a rule to a target range (`"B2:B20"`, a whole column `"B2:B1048576"`, or a space-separated
   * `sqref` of several ranges). The rule is stored once against the range, not copied per covered
   * cell, so a whole-column dropdown stays a single entry. A cell inside the range reports the rule
   * through {@link at}.
   *
   * Pass `{extended: true}` to mark a rule that belongs in the 2009 extension form
   * (`<x14:dataValidation>`) — the carrier Excel uses for a list source on another sheet and other
   * shapes the standard element cannot express. The reader sets it for a rule found in that form so a
   * round-trip writes it back there instead of silently corrupting the cross-sheet reference.
   */
  add(sqref: string, rule: DataValidation, options: {extended?: boolean} = {}): void {
    // One defensive copy, shared by the serialisable entry and the lookup index, so the getter never
    // hands back a reference into the caller's object.
    const stored = cloneDataValidation(rule);
    const entry: DataValidationEntry = {sqref, rule: stored};
    if (options.extended) entry.extended = true;
    this.#entries.push(entry);
    this.#rects.push({rects: decodeSqrefRects(sqref), rule: stored});
  }

  /** The data validations on this sheet, each bound to its target range, in insertion order. */
  get entries(): readonly DataValidationEntry[] {
    return this.#entries;
  }

  /**
   * The validation covering the 1-based `col`/`row`, or `undefined` when none does. The first added
   * rule whose range contains the cell wins, mirroring how a spreadsheet resolves overlapping
   * validations.
   */
  at(col: number, row: number): DataValidation | undefined {
    for (const {rects, rule} of this.#rects) {
      for (const rect of rects) {
        if (col >= rect.left && col <= rect.right && row >= rect.top && row <= rect.bottom) {
          return rule;
        }
      }
    }
    return undefined;
  }

  /** Drop every validation, leaving the overlay empty. */
  clear(): void {
    this.#entries.length = 0;
    this.#rects.length = 0;
  }
}
