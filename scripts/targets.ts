// Which files a tool runs over, said once.
//
// Four of these lists lived in two or three places each, and two of them carried a comment asking
// the next editor to keep the copies in step. `check-constitution.ts` makes the argument against
// exactly that: duplication is only safe when drift is impossible, and the way drift is made
// impossible here is a machine check, not a promise to remember. A promise had already been broken
// once: `format.ts`'s targets claimed to track the tsconfig include lists and had gained `www/`,
// which no root-family project covers.
//
// The consumers are `scripts/verify.ts`, the thin package scripts beside this file, and
// `scripts/check-tsconfig-coverage.ts`, which asserts that every tree named here is claimed by a
// `tsconfig.json` so the linter reads this repo's options rather than TypeScript's defaults.

/**
 * The top-level trees `tsconfig.test.json` spans, each of which must carry its own `tsconfig.json`
 * for the type-aware linter's config search. `src/` is not here: the root `tsconfig.json` claims it.
 */
export const HARNESS_TREES = ['test', 'scripts', 'tools'] as const;

/**
 * What `lint` covers. Wider than what `tsc` covers, deliberately: `www/` is a real TypeScript tree
 * with its own project, and a rule that stopped at the library would stop at the place a broken
 * sample is most visible.
 */
export const LINT_TARGETS = ['src', ...HARNESS_TREES, 'www', 'charcheck.config.ts'] as const;

/**
 * What the formatter rewrites. Every tree any project claims, plus the one config file that lives at
 * the root; formatting is not a type-aware question, so this follows the lint targets rather than the
 * compiler's include lists, which is what the comment here used to claim and was already wrong about.
 */
export const FORMAT_TARGETS = [
  'src/**/*.ts',
  ...HARNESS_TREES.map((tree) => `${tree}/**/*.ts`),
  'www/**/*.ts',
  'charcheck.config.ts',
] as const;

/** The unit suite, as `node --test` selects it. The corpus is its own runner and its own gate. */
export const UNIT_SUITE_GLOB = 'src/**/*.test.ts';

/**
 * What lives under `src/` and is not the library: the tests, the test support, and the type-level
 * assertions. `tsconfig.build.json` excludes exactly these from the emit, and `scripts/coverage.ts`
 * excludes them from the measurement, for one reason said once: what is not shipped is not the
 * library, and an instrument that reports on itself flatters itself.
 */
export const NON_LIBRARY_GLOBS = [
  'src/**/*.test.ts',
  'src/**/*.test-support.ts',
  'src/type-tests/**',
] as const;

// CLAUDE.md admits no warnings, and oxlint exits 0 on them. Nothing in .oxlintrc.jsonc is set to
// "warn" today, so this changes no current outcome. It is here so that the first rule adopted at
// warning severity, to stage a migration, is still a gate rather than a message.
export const LINT_STRICT = '--deny-warnings';
// A suppression that has outlived its cause is worse than none: it reads as a live hazard and
// silences nothing.
export const LINT_UNUSED_DIRECTIVES = '--report-unused-disable-directives';
export const LINT_TYPE_AWARE = '--type-aware';
