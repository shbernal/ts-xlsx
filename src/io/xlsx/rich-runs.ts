// Rich-text run accumulation, shared by the two readers that parse `<r>` runs identically: an inline
// string's `<is>` in a worksheet body and a pooled `<si>` in sharedStrings.xml. One `<r>` opens a run
// (resetting its font and text so an unformatted run inherits nothing from the last), an `<rPr>` opens
// the run's font bundle whose self-closing children each set one facet, and a `<t>` appends the run's
// text; the run commits on `</r>`. A `<t>` no run claimed is the container's own plain text.
//
// It drives itself, the way `CellAccumulator` and `TextCapture` do. {@link open}/{@link text}/
// {@link close} are the element machine that the shared-string reader and the inline-string reader
// used to spell out identically, six raw state calls each, over a grammar the two only *asserted*
// was the same. A drift there means a rich string reads one way when Excel pooled it and another
// way when it is inline: the same cell content, two answers, decided by an encoding choice the
// author never made. The only thing that genuinely differs is which element name opens the
// container, which is now a constructor argument rather than a hand-copied switch.
//
// Each caller keeps only what is actually its own: what committing an `<si>` means, and whether the
// runs are read at all (the row streamer deliberately reads none, so a rich inline string flattens
// to its concatenated text there).
//
// One invariant spans both, and it is {@link RunAccumulator.beginContainer} that names it: an
// accumulator is emptied when a *container* opens, never when one closes. It has to be that way round,
// because the runs are read after the container closes (a `<c>` decodes its value at `</c>`, well past
// the `</is>` that ended the runs) so draining them at the close would take them before the consumer
// arrives. The cost is that a caller with a container of its own (a `<c>` that may hold an `<is>`)
// must say so, which is why the method is named for the container rather than for what it does.

import type {Font} from '../../core/style.ts';
import type {RichTextRun} from '../../core/value.ts';
import {decodeSpreadsheetText, TextCapture, type XmlAttributes} from '../../xml/xml-read.ts';
import {applyFontChild, type FontDraft} from './read-styles.ts';

/**
 * What {@link RunAccumulator.close} did with an element: `'container'` means the `<si>`/`<is>` itself
 * closed and the caller should commit whatever it makes of the gathered runs, the one step the two
 * readers do differently; `'claimed'` means the machine handled it; `'other'` leaves it to the caller.
 *
 * Three values rather than a boolean because both callers need to tell "the machine dealt with it"
 * from "your string is complete", and a boolean would mean something different to each of them.
 */
export type RunClose = 'claimed' | 'container' | 'other';

export class RunAccumulator {
  readonly #container: string;
  readonly #readRuns: boolean;
  // `<t>` is the only element this machine captures text for. Owning a capture rather than a raw
  // flag is what keeps a self-closing `<t/>`, legal and firing no close, from arming a capture that
  // the next element's text then falls into.
  readonly #capture = new TextCapture('t');

  #runs: RichTextRun[] = [];
  // The open run's font, which outlives its `</rPr>` because the run commits it at `</r>`; and,
  // separately, whether an `<rPr>` is open *right now*. Two flags rather than one because the
  // "unrecognised element is a font facet" branch must ask the second question: a draft still held
  // for a run that has not committed is not a licence to absorb the next element in the document.
  #font: FontDraft | null = null;
  #inProperties = false;
  #runText = '';
  #plain = '';
  #inRun = false;
  #isRich = false;
  #inContainer = false;

  /**
   * @param options.container the element that opens one string: `'si'` for the shared-string pool,
   * `'is'` for an inline string in a worksheet body.
   * @param options.readRuns whether `<r>`/`<rPr>` open a rich-text run. The buffered readers read
   * them; the row stream deliberately does not, and with runs off every `<t>` falls through to
   * {@link plainText}, which is exactly the documented flattening.
   */
  constructor(options: {readonly container: 'si' | 'is'; readonly readRuns: boolean}) {
    this.#container = options.container;
    this.#readRuns = options.readRuns;
  }

  /** The runs gathered so far. */
  get runs(): RichTextRun[] {
    return this.#runs;
  }

  /** Whether the container held at least one `<r>`: what decides rich text from a plain string. */
  get isRich(): boolean {
    return this.#isRich;
  }

  /** The container's own text: every `<t>` no run claimed, concatenated and already unescaped. */
  get plainText(): string {
    return this.#plain;
  }

  /**
   * Open a string container by discarding whatever the last one gathered. Called for the machine's
   * own container automatically; a caller whose own element may *hold* one (a `<c>` around an
   * `<is>`) calls it too, or that container inherits the previous one's runs, which is the one way
   * this accumulator can be misused.
   *
   * A new array is installed rather than the existing one emptied, so a value already built from the
   * last container's runs keeps its own array.
   */
  beginContainer(): void {
    this.#runs = [];
    this.#font = null;
    this.#inProperties = false;
    this.#runText = '';
    this.#plain = '';
    this.#inRun = false;
    this.#isRich = false;
  }

  /** Drive one element open, and return whether it was one of this machine's own. */
  open(local: string, attrs: XmlAttributes, selfClosing: boolean): boolean {
    switch (local) {
      case this.#container:
        this.#inContainer = true;
        this.beginContainer();
        return true;
      case 'r':
        if (!this.#readRuns || !this.#inContainer) return false;
        this.#isRich = true;
        // Reset the per-run font and text so an unformatted run inherits nothing from the last.
        this.#inRun = true;
        this.#font = null;
        this.#inProperties = false;
        this.#runText = '';
        return true;
      case 'rPr':
        if (!this.#readRuns || !this.#inRun) return false;
        this.#font = {};
        this.#inProperties = true;
        return true;
      case 't':
        this.#capture.open(local, selfClosing);
        return true;
      default:
        // An unrecognised element is one facet of the open run's font (`<b/>`, `<sz>`, `<color>`,
        // `<rFont>`, …), and only while an `<rPr>` is open. Both readers used to spell this as
        // "anything I do not recognise is a run property", which silently absorbed an element that
        // was someone else's; the guard is what lets this say no.
        if (!this.#inProperties || this.#font === null) return false;
        applyFontChild(this.#font, local, attrs);
        return true;
    }
  }

  /** Feed one run of character data. Ignored unless a `<t>` is open. */
  text(chunk: string): void {
    this.#capture.text(chunk);
  }

  /** Drive one element close. See {@link RunClose}. */
  close(local: string): RunClose {
    switch (local) {
      case 't': {
        // A `<t>` inside a run is that run's text; one directly in the container is its plain text.
        // A run takes precedence, since a run is also inside the container.
        //
        // The `_xHHHH_` decode happens here, on one whole `<t>`. It cannot move up into the SAX text
        // callback: that fires once per run of character data and an entity splits a run, so `_x00`
        // and `01_` can arrive separately and a per-chunk decode would miss the escape in exactly
        // those strings that happen to contain an `&`.
        const text = this.#capture.close(local);
        if (text !== undefined) {
          const decoded = decodeSpreadsheetText(text);
          if (this.#inRun) this.#runText += decoded;
          else if (this.#inContainer) this.#plain += decoded;
        }
        return 'claimed';
      }
      case 'rPr':
        if (!this.#readRuns || !this.#inProperties) return 'other';
        this.#inProperties = false;
        return 'claimed';
      case 'r': {
        if (!this.#readRuns || !this.#inRun) return 'other';
        this.#inProperties = false;
        // Commit the run, with its font only if that font set at least one facet.
        const run: {text: string; font?: Font} = {text: this.#runText};
        if (this.#font !== null && Object.keys(this.#font).length > 0) run.font = this.#font;
        this.#runs.push(run);
        this.#inRun = false;
        return 'claimed';
      }
      case this.#container:
        this.#inContainer = false;
        return 'container';
      default:
        return 'other';
    }
  }
}
