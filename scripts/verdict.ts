// How a check in this directory reports what it found.
//
// Four of them - layering, entries, browser-safe, source-text - ended with the same fifteen lines:
// a one-line summary on success, and on failure a counted headline, the findings indented and
// blank-line separated, and a non-zero exit. Four copies that had already drifted in the two ways
// that matter to a reader. Three exited with `process.exit(1)` and one with `process.exitCode`, and
// only the one that had a compiler process to hand back had noticed the difference; `process.exit`
// truncates a pipe that has not flushed, which is how a gate loses the diagnostic it just printed
// on exactly the runs where someone needs it. And the failure headline's noun was per-check prose
// ("break(s) in the browser boundary") while the *success* line was per-check prose too, so a reader
// scanning nine gates had nine shapes to recognise instead of one.
//
// So: the shape is fixed here and the prose stays with the check, which is the half that actually
// carries information. `process.exitCode` always, never `process.exit`.

/**
 * Report a check's result and set the exit code.
 *
 * @param options.gate the check's name, which prefixes both lines. Matches the `*:check` package
 * script and the gate name `scripts/verify.ts` reports, so a failure names the command to re-run.
 * @param options.problems one entry per finding, already formatted. Indent continuation lines by
 * four spaces; they are printed blank-line separated, so a finding may be several lines.
 * @param options.ok the one-line summary printed when there are none. Say what was checked and how
 * much of it, not merely that it passed: the number is what makes a gate that quietly stopped
 * looking visible.
 * @param options.failure what the count counts, as a noun phrase (`'import(s) cross a layer
 * boundary'`). Printed as `<gate>: <n> <failure>.`
 */
export function verdict(options: {
  readonly gate: string;
  readonly problems: readonly string[];
  readonly ok: string;
  readonly failure: string;
}): void {
  const {gate, problems, ok, failure} = options;
  if (problems.length === 0) {
    console.log(`${gate}: ${ok}`);
    return;
  }
  console.error(`\n${gate}: ${problems.length} ${failure}.\n`);
  console.error(`${problems.join('\n\n')}\n`);
  process.exitCode = 1;
}

/**
 * A bad invocation, not a failing check: one legible line, no stack.
 *
 * The distinction is the whole point of the class. A stack trace answers "where did this break",
 * which is the wrong question when the answer is "you passed `--jobs banana`", and it buries the one
 * line that would have told the reader so.
 */
export class UsageError extends Error {}

/**
 * The tail every script's top-level `main().catch` needs: a {@link UsageError} as its own message,
 * anything else with its stack, and a non-zero exit either way.
 */
export function reportCrash(gate: string, err: unknown): void {
  if (err instanceof UsageError) console.error(`${gate}: ${err.message}`);
  else {
    console.error(
      `${gate} failed: ${err instanceof Error ? (err.stack ?? err.message) : String(err)}`,
    );
  }
  process.exitCode = 1;
}
