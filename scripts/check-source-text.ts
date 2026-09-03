#!/usr/bin/env node
// Authored text stays legible to the tools that read text.
//
// A literal NUL in a .ts file costs nothing at compile time. It is valid UTF-8, and tsc does not
// care. Every other text tool flips into binary mode. grep answers `Binary file … matches`
// instead of the matching lines, so the file stops appearing in searches while still appearing to
// have been searched. That is how `src/core/range.ts` hid two `@throws` tags from a doc audit: the
// grep declined to quote them, the tally came back one short, and nothing anywhere reported a
// problem. A check that fails loudly is worth more than a file that lies quietly.
//
// The sentinel it was spelling is fine; writing it as the character rather than as the escape was
// not. Those two are identical to the compiler and opposite to everything else: exactly the kind
// of difference no reviewer catches by eye, and so exactly the kind a machine should catch.
//
// Two classes are refused:
//
//   1. C0 controls other than tab/LF/CR, plus DEL. The tool-breaking class above.
//   2. Bidirectional overrides (U+202A–202E, U+2066–2069), which reorder how source *renders*
//      without changing what it *means*, so a reviewer can approve the opposite of what compiles
//      (CVE-2021-42574). We parse untrusted spreadsheets; we need not also read untrusted source.
//
// Scope is an allowlist of authored text extensions, not a denylist of binary ones. The corpus is
// full of .xlsx fixtures, so the failure mode of guessing wrong is a build that fails on a file for
// being itself. An unlisted extension goes unchecked instead, so add it here when we start
// authoring one.
//
// Which *files* are authored is git's answer, not a hand-kept list of directories to walk past.
// A walk that skipped by directory name held for as long as the derived trees were all called
// `dist` or `node_modules`; the moment `www` joined the roots it started reading
// `www/.vitepress/cache/`, which is a bundler's scratch, is gitignored, and is full of the very
// control characters this gate exists to refuse. So it failed for anyone who had run the site
// locally and passed in CI, where the cache does not exist, and a gate that disagrees with itself
// depending on what you ran last teaches people to ignore it. Tracked files plus untracked files
// git would not ignore *is* the set of things a human or agent wrote here, so we ask for exactly
// that and let one definition of "derived" serve both git and this check.
//
//   node scripts/check-source-text.ts

import {spawnSync} from 'node:child_process';
import {existsSync, readFileSync} from 'node:fs';
import {extname, join} from 'node:path';

import {ROOT} from './repo.ts';
import {verdict} from './verdict.ts';

/** The trees we author by hand or generate into. `skills` is published, so it is source too. */
const ROOTS = ['src', 'scripts', 'test', 'tools', 'docs', 'skills', 'www'];

const TEXT_EXTENSIONS = new Set([
  '.ts',
  '.js',
  '.mjs',
  '.cjs',
  '.json',
  '.jsonc',
  '.md',
  '.yml',
  '.yaml',
  '.xml',
  '.txt',
  '.ps1',
]);

const CONTROL_NAMES = new Map([
  [0x00, 'NUL'],
  [0x07, 'BEL'],
  [0x08, 'BS'],
  [0x0b, 'VT'],
  [0x0c, 'FF'],
  [0x1a, 'SUB'],
  [0x1b, 'ESC'],
  [0x7f, 'DEL'],
]);

const CONTROL_HINT =
  "a literal control character reads as binary to grep and friends; write it as an escape ('\\u0000')";
const BIDI_HINT =
  'a bidirectional override reorders how the line renders without changing what it means';

const codePoint = (code: number) => `U+${code.toString(16).toUpperCase().padStart(4, '0')}`;

interface Offence {
  readonly label: string;
  readonly hint: string;
}

/** What is wrong with this code point, or undefined if nothing is. */
function classify(code: number): Offence | undefined {
  // Tab, LF and CR are how text is shaped; CR in particular is ordinary in a Windows checkout.
  if (code === 0x09 || code === 0x0a || code === 0x0d) return undefined;
  if (code < 0x20 || code === 0x7f) {
    return {
      label: `${CONTROL_NAMES.get(code) ?? 'control'} (${codePoint(code)})`,
      hint: CONTROL_HINT,
    };
  }
  if ((code >= 0x202a && code <= 0x202e) || (code >= 0x2066 && code <= 0x2069)) {
    return {label: `bidi override (${codePoint(code)})`, hint: BIDI_HINT};
  }
  return undefined;
}

interface Problem extends Offence {
  readonly file: string;
  readonly line: number;
  readonly column: number;
}

function scan(file: string, problems: Problem[]): void {
  const lines = readFileSync(join(ROOT, file), 'utf8').split('\n');
  for (const [index, line] of lines.entries()) {
    // Iterating the string yields whole code points, so an astral character counts as one column
    // and cannot be mistaken for a pair of surrogates that happen to look like controls.
    let column = 0;
    for (const character of line) {
      column += 1;
      const code = character.codePointAt(0);
      if (code === undefined) continue;
      const offence = classify(code);
      if (offence !== undefined) problems.push({file, line: index + 1, column, ...offence});
    }
  }
}

/** Every authored text file under `ROOTS`, repo-relative: tracked, or untracked and not ignored. */
function collect(): string[] {
  const result = spawnSync(
    'git',
    ['ls-files', '--cached', '--others', '--exclude-standard', '-z', '--', ...ROOTS],
    {cwd: ROOT, encoding: 'utf8', maxBuffer: 64 * 1024 * 1024},
  );
  if (result.status !== 0) {
    throw new Error(`git ls-files failed (${result.status ?? 'signal'}): ${result.stderr.trim()}`);
  }
  // A NUL-delimited listing so a path with a space or a quote in it arrives whole and unquoted.
  const paths = result.stdout.split('\u0000').filter((path) => path !== '');
  // `--cached` still lists a tracked file deleted from the working tree, and a path can arrive
  // from both halves of the listing, so drop what is not on disk and dedupe what is.
  const authored = paths.filter(
    (path) => TEXT_EXTENSIONS.has(extname(path).toLowerCase()) && existsSync(join(ROOT, path)),
  );
  return [...new Set(authored)].sort();
}

const files = collect();

const problems: Problem[] = [];
for (const file of files) scan(file, problems);

verdict({
  gate: 'source-text',
  problems: problems.map(
    (problem) =>
      `  ${problem.file}:${problem.line}:${problem.column}  ${problem.label}\n      ${problem.hint}`,
  ),
  ok: `${files.length} files, no control characters or bidi overrides`,
  failure: 'problem(s) in authored text',
});
