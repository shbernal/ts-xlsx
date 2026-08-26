// Banned characters in authored prose, checked by charcheck.
//
// The em dash is the one character this repo actually argues about. It is the punctuation a
// language model reaches for when a sentence has two ideas in it and no one decided which is
// the main one, so a document full of them reads as unedited rather than as emphatic. The fix
// is almost never a different dash: it is a full stop, a colon, or a pair of commas, and which
// one it is depends on the sentence. That is why this rule declares no `fix`. An automatic
// rewrite to `-` would pass the check and leave the prose worse than it found it, which is the
// opposite of the point.
//
// U+2015 HORIZONTAL BAR rides along because it renders identically at every size a reviewer
// reads at, so banning one without the other bans nothing. U+2013 EN DASH is deliberately NOT
// banned: `docs/knowledge/specs/` uses it for numeric ranges, which is what it is for.
//
// Scope is `raw`, not `markdown`, so fenced code blocks are read too. The `markdown` scope
// would be the more precise instrument, but it needs `micromark` and its ~25 transitive
// packages to skip regions that today hold not one banned character, and CLAUDE.md section 2
// keeps the dependency tree small.
//
// `raw` is also the only shape the source rule below could take, and that one is not a
// preference. charcheck's scopes are `raw`, `strings`, `markup`, `markdown` and `html`; none
// of them reads comments, and `strings` reads the opposite. So a rule that guards a doc
// comment guards the whole file, which is why the string literals were recast too rather
// than left as the one place in `src/` the character still lives.
//
// Know the exit before you need it, because it is narrower than it looks. A suppression
// marker inside a fenced block is ignored in a .md file under every scope, by design, so that
// a page documenting the syntax does not silence itself. Under `raw` the fence is still
// scanned. So a banned character in quoted tool output cannot be silenced line by line at
// all: only `charcheck-disable-file` reaches it, and that blinds the rule to the same file's
// prose, which is the part worth checking. On the day that happens, fix the quotation, or
// exclude the one file, or move this rule to `markdown` and pay for micromark. Do not reach
// for a file-level disable, which reads like a local exception and is not one.
//
//   node node_modules/charcheck/dist/cli.js            # the whole tree
//   node node_modules/charcheck/dist/cli.js --staged   # what this commit would land

import {defineConfig} from 'charcheck/config';

/**
 * Em dash and its lookalike, built from code points rather than typed. A config that spelled
 * its banned characters literally would be the one file in the repo guaranteed to contain
 * them, so it could never widen its own `include` to cover itself, and a reviewer would have
 * to tell U+2014 from U+2013 by eye in the one place where the difference is the whole point.
 */
const EM_DASHES = [String.fromCodePoint(0x2014), String.fromCodePoint(0x2015)];

export default defineConfig({
  rules: [
    {
      id: 'no-em-dash-in-docs',
      chars: EM_DASHES,
      message:
        'em dash in authored prose: recast the sentence (full stop, colon, or commas) rather than swapping the dash',
      include: ['docs/**/*.md'],
    },
    {
      id: 'no-em-dash-in-source',
      chars: EM_DASHES,
      message:
        'em dash in source prose: recast the sentence (full stop, colon, or commas) rather than swapping the dash',
      // Comments and string literals alike. An error message is prose too, read by someone
      // under stress, and it is the half a comments-only rule would have had to leave out.
      // `scripts/`, `test/` and `tools/` are not in yet: they still carry the character, and a
      // rule aimed at a tree that trips it reports on every run until someone learns to ignore
      // the output. Widen this glob in the change that cleans them, not before.
      include: ['src/**/*.ts'],
    },
  ],
});
