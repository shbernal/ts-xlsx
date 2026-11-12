// Which rows of a streamed sheet are collapsed outline summaries, decided as the rows are written rather
// than over the whole sheet at the end (ADR 0046, decision 5). The rule is the buffered writer's
// (`scanRowOutline` in `worksheet-xml.ts`): a row is a collapsed summary when its adjacent detail run,
// the contiguous rows on the summary side with a deeper level, is non-empty and every row in it is
// hidden. What differs is which rows can be seen when a summary row is written.

/** One row's outline facts, as the writer reads them off the model before evicting it. */
export interface OutlineRow {
  readonly number: number;
  readonly level: number;
  readonly hidden: boolean;
}

/**
 * The trailing detail runs of the rows written so far, one per outline threshold, so a summary row
 * below its detail can be answered without keeping the rows above it. `#runs[t]` describes the
 * contiguous run of rows ending at `#last` whose level is at least `t`: absent when the last row is
 * shallower than `t`, otherwise whether every row in it is hidden. Its size is the deepest level seen,
 * not the row count.
 */
export class StreamedOutline {
  readonly #summaryBelow: boolean;
  #last = 0;
  #runs: (boolean | undefined)[] = [];

  constructor(summaryBelow: boolean) {
    this.#summaryBelow = summaryBelow;
  }

  /**
   * The collapsed summaries among `batch`, the rows written together in ascending order. With summary
   * rows below, every row above a batch has already been folded into the running state. With summary
   * rows above, the detail run is looked for inside the batch only, because the rows after it have not
   * been written: a group whose rows are committed after its summary is not derived.
   */
  collapsedIn(batch: readonly OutlineRow[]): Set<number> {
    const collapsed = new Set<number>();
    batch.forEach((row, index) => {
      if (this.#summaryBelow ? this.#closesHiddenRun(row) : hiddenRunFollows(batch, index)) {
        collapsed.add(row.number);
      }
      this.#note(row);
    });
    return collapsed;
  }

  #closesHiddenRun(row: OutlineRow): boolean {
    return row.number === this.#last + 1 && this.#runs[row.level + 1] === true;
  }

  #note(row: OutlineRow): void {
    // A gap in the numbering ends every group, as an absent row does in the buffered walk.
    if (row.number !== this.#last + 1) this.#runs = [];
    for (let threshold = 1; threshold <= row.level; threshold++) {
      const run = this.#runs[threshold];
      this.#runs[threshold] = (run ?? true) && row.hidden;
    }
    this.#runs.length = row.level + 1;
    this.#last = row.number;
  }
}

function hiddenRunFollows(batch: readonly OutlineRow[], index: number): boolean {
  const summary = batch[index];
  if (summary === undefined) return false;
  let previous = summary.number;
  let sawDetail = false;
  for (let next = index + 1; next < batch.length; next++) {
    const detail = batch[next];
    if (detail === undefined || detail.number !== previous + 1 || detail.level <= summary.level) {
      break;
    }
    if (!detail.hidden) return false;
    sawDetail = true;
    previous = detail.number;
  }
  return sawDetail;
}
