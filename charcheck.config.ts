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
// than left as the one place in `src/` the character still lives. A `comments` scope is
// asked for upstream at charcheck#22; even if it ships, this rule keeps `raw`, since the
// include below also reaches `.json` and `.ps1`, which no TypeScript scope reads.
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
      // Every include names a directory, so no root-level `.md` is in scope. That is a decision for
      // the files it is usually about: an ephemeral plan and its follow-up list live at the root
      // (`.git/info/exclude` names them), and they are scaffolding deleted when the work lands, so
      // holding them to the prose bar of a page someone will read next year buys nothing, and a
      // finding on a file git is already told to ignore reads as noise.
      //
      // It is not yet a decision for `README.md`, `CLAUDE.md` and `CHANGELOG.md`, which the sweep
      // that seeded this rule never reached and which hold 161 findings between them. Widening the
      // glob before recasting that prose would only redden the gate. Do the sweep first, then add
      // those three by name; do not reach for a bare `*.md`, which would pull the plans back in.
      include: ['docs/**/*.md', 'test/**/*.md', 'tools/**/*.md'],
    },
    {
      id: 'no-em-dash-in-source',
      chars: EM_DASHES,
      message:
        'em dash in source prose: recast the sentence (full stop, colon, or commas) rather than swapping the dash',
      // Comments and string literals alike. An error message is prose too, read by someone
      // under stress, and it is the half a comments-only rule would have had to leave out.
      // The `tsconfig.json` files are here for the same reason their `"//"` keys exist: they
      // are a paragraph of prose that happens to live in a JSON value. `.ps1` is in because a
      // COM driver's header comment is where its guardrails are explained, and that is prose
      // whoever next has to debug a hung Excel will read. `charcheck.config.ts` is in because
      // a config that bans a character and then contains one is the one file nobody would
      // check; building `EM_DASHES` from code points is what makes that possible.
      include: [
        'charcheck.config.ts',
        'src/**/*.ts',
        'scripts/**/*.ts',
        'scripts/**/*.json',
        'test/**/*.ts',
        'test/**/*.json',
        'test/**/*.ps1',
        'tools/**/*.ts',
        'tools/**/*.json',
        'tools/**/*.ps1',
      ],
      // A probe record is evidence, not prose we may reword later. `verdict` and `description`
      // are what the author concluded on the day Excel was asked, and both the probe spec and
      // the corpus fixture it seeded carry the same sentences; recasting one and not the other
      // invents a drift, recasting both edits the record. 16 findings, declined on purpose.
      //
      // The character survives in two other places under `test/`, and neither is an exception
      // to the rule: both are cell content a case asserts on byte-for-byte. The .ts writes the
      // character as a unicode escape, the same fix `check-source-text.ts` asks for elsewhere;
      // the PowerShell fixture author, whose single-quoted strings admit no escape, carries a
      // line suppression with the reason beside it.
      exclude: ['tools/excel-oracle/probes/**', 'test/corpus/fixtures/excel-oracle/**'],
    },
  ],
});
