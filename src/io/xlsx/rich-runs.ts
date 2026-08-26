// Rich-text run accumulation, shared by the two readers that parse `<r>` runs identically: an inline
// string's `<is>` in a worksheet body and a pooled `<si>` in sharedStrings.xml. One `<r>` opens a run
// (resetting its font and text so an unformatted run inherits nothing from the last), an `<rPr>` opens
// the run's font bundle whose self-closing children each set one facet, and a `<t>` appends the run's
// text; the run commits on `</r>`. The surrounding parser owns only whether it is inside an `<is>`/`<si>`
// and where a bare (non-run) `<t>` goes.
//
// One invariant spans both readers, and it is {@link RunAccumulator.beginContainer} that names it: an
// accumulator is emptied when a *container* opens, never when one closes. It has to be that way round,
// because the runs are read after the container closes (a `<c>` decodes its value at `</c>`, well past
// the `</is>` that ended the runs) so draining them at the close would take them before the consumer
// arrives. The cost is that every caller that opens a container must say so, which is why the method is
// named for the container rather than for what it does to the state.

import type {Font} from '../../core/style.ts';
import type {RichTextRun} from '../../core/value.ts';
import type {XmlAttributes} from '../../xml/xml-read.ts';
import {applyFontChild, type FontDraft} from './read-styles.ts';

export class RunAccumulator {
  #runs: RichTextRun[] = [];
  #font: FontDraft | null = null;
  #text = '';
  #inRun = false;

  // Open a string container (an `<is>`, an `<si>`, or the `<c>` that may hold an `<is>`) by
  // discarding whatever the last one gathered. Every container must open here or it inherits the
  // previous container's runs, which is the one way this accumulator can be misused.
  //
  // A new array is installed rather than the existing one emptied, so a value already built from the
  // last container's runs keeps its own array.
  beginContainer(): void {
    this.#runs = [];
    this.#font = null;
    this.#inRun = false;
  }

  /** The runs gathered so far. */
  get runs(): RichTextRun[] {
    return this.#runs;
  }

  // Open a run: reset the per-run font and text so an unformatted run inherits nothing from the last.
  beginRun(): void {
    this.#inRun = true;
    this.#font = null;
    this.#text = '';
  }

  // Open the run's `<rPr>` font bundle; its children stream in via {@link applyProperty}.
  beginProperties(): void {
    if (this.#inRun) this.#font = {};
  }

  // Apply one `<rPr>` child (`<b/>`, `<sz>`, `<color>`, `<rFont>`, …) to the open run's font; a no-op
  // when no `<rPr>` is open.
  applyProperty(local: string, attrs: XmlAttributes): void {
    if (this.#font !== null) applyFontChild(this.#font, local, attrs);
  }

  // Append a `<t>`'s text to the open run. Returns false when no run is open, so the caller routes the
  // text to its surrounding plain/inline string instead.
  appendText(text: string): boolean {
    if (!this.#inRun) return false;
    this.#text += text;
    return true;
  }

  // Commit the open run to the runs list, with its font only if that font set at least one facet.
  endRun(): void {
    if (!this.#inRun) return;
    const run: {text: string; font?: Font} = {text: this.#text};
    if (this.#font !== null && Object.keys(this.#font).length > 0) run.font = this.#font;
    this.#runs.push(run);
    this.#inRun = false;
  }
}
