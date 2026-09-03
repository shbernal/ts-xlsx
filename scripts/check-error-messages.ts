#!/usr/bin/env node
// One spelling for a name inside an error message.
//
// `src/errors.ts` states the rule and gives it a name: almost every message this library throws
// names something the caller or the file chose - a sheet, a table, a defined name, a part path - and
// the tree had grown three ways of setting that name off from the prose (`"…"`, `'…'`, and
// `JSON.stringify`), split by directory rather than by intent. Only the third survives a name that
// itself contains a quote, a newline, or a zero-width character, which for a library whose input is
// untrusted is the difference between a diagnostic and a decoy. So `quoted()` is that third one,
// and the point is not the algorithm - it *is* `JSON.stringify` - but that every throw site reaches
// the same one.
//
// It did not hold. Twenty-three modules called `quoted()` and nine called `JSON.stringify` inline,
// two of them in a file that did both and one of those in the same function, four lines apart. That
// is not a rule anyone broke on purpose; it is what a rule with no mechanism decays into. The
// convention is invisible at the throw site - both spellings produce identical output today - so
// review cannot see the difference and a reader has no way to know one of them is the one.
//
// oxlint has no `no-restricted-syntax`, so this is the mechanism instead: every `JSON.stringify` in
// production source is either one of the two named below or a finding.
//
//   node scripts/check-error-messages.ts

import {readFileSync} from 'node:fs';
import {join} from 'node:path';

import {sourceFiles, withoutComments} from './module-graph.ts';
import {ROOT} from './repo.ts';
import {reportCrash, verdict} from './verdict.ts';

const GATE = 'error-messages';

/**
 * The calls that are not a message at all, each with the reason it is not.
 *
 * An allowlist rather than a heuristic about what "inside a throw" means: a message is often built a
 * statement above the `throw`, so a syntactic proximity rule would miss the cases that matter and
 * argue about the ones that do not. Two entries is small enough to read, and a third would be a
 * decision worth making explicitly.
 */
const ALLOWED = new Map([
  ['src/errors.ts', "`quoted`'s own implementation: this is the one spelling"],
  ['src/core/range.ts', 'the replacer form, serialising a value rather than naming it'],
]);

function main(): void {
  const problems: string[] = [];
  let files = 0;
  for (const file of sourceFiles(join(ROOT, 'src'), '.ts')) {
    const relative = file.slice(`${ROOT.replaceAll('\\', '/')}/`.length);
    files += 1;
    if (ALLOWED.has(relative)) continue;
    // Comments are blanked rather than removed, so an offset still maps to the line it came from -
    // and the rule is about code, where `src/errors.ts` names `JSON.stringify` in its own prose.
    const source = withoutComments(readFileSync(file, 'utf8'));
    for (const match of source.matchAll(/\bJSON\.stringify\b/g)) {
      const line = source.slice(0, match.index).split('\n').length;
      problems.push(
        `${relative}:${line}\n    JSON.stringify in an error message; call quoted() from src/errors.ts\n` +
          '    (or invalidToken() for the "not a value the OOXML enumeration allows" template)',
      );
    }
  }
  verdict({
    gate: GATE,
    problems,
    ok: `every name in a message goes through quoted(), across ${files} source file(s)`,
    failure: 'inline JSON.stringify call(s) bypass the one spelling',
  });
}

try {
  main();
} catch (error) {
  reportCrash(GATE, error);
}
