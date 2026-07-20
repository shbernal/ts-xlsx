# Declaration of Principles

> This file is the constitution of the project. Every agent — human or AI — reads
> it before touching the code. `AGENTS.md` is a symlink to this file so that all
> runtimes converge on the same rules. If a decision is not covered here, decide
> in the *spirit* of this document, act, and record what you did.

---

## 0. What this is

This is a **hard fork** of ExcelJS with the explicit goal of becoming a **fully
independent, modern, TypeScript-first library** for reading and writing
spreadsheet (OOXML / `.xlsx`, and adjacent) documents.

We forked because upstream `exceljs/exceljs` is effectively unmaintained: no merge
to `master` since January 2024, no release since v4.4.0 (October 2023), ~650 open
issues and ~140 open PRs rotting on the vine — while the package still serves tens
of millions of downloads a month. There is enormous accumulated value trapped in
that backlog and enormous accumulated debt in that codebase. We intend to **extract
the value and discard the debt.**

We are not a compatibility fork. We do not exist to keep other people's imports
working. We exist to be the *best* spreadsheet library for a world where most code
is written and maintained by autonomous agents alongside humans.

---

## 1. Non-negotiable stances

1. **This is a clean break.** We aim for full independence from upstream. The harvest is
   complete (see `docs/architecture.md`) and we no longer track upstream.

2. **No backwards-compatibility guarantee.** Ever. The old API is a *reference*,
   not a contract. If a better shape exists, we take it.

3. **Breaking changes are welcome, not tolerated.** A breaking change that makes
   the library clearer, safer, or faster is a *good day*. We version honestly
   (SemVer major bumps are cheap and expected during the rebuild) and document the
   break — but we never contort the design to avoid one.

4. **Nothing is sacred.** No comment, no naming convention, no file layout, no
   abstraction survives on the grounds that "it was already there." Legacy status
   is not a reason to keep something. It is, if anything, a reason to re-examine it.

5. **We do not keep legacy code.** Code is either *useful and modern* or it is
   *deleted*. There is no "leave it, it works" tier. If it works but is ugly,
   untyped, untested, or unclear — it is not done, it is debt, and debt gets paid
   or removed.

6. **AI-first, human-minimal.** This project is optimized so that autonomous agents
   can do the overwhelming majority of the work with high confidence and low human
   supervision. The human's time is the scarcest resource; spend it only where
   judgment is genuinely irreducible. **Do not clobber the human with questions.**
   Make the reasonable call, act, and note the assumption so it can be reverted.

---

## 2. Quality is the moat

Because humans intervene rarely, **the machine-checkable safety net is the primary
guarantor of correctness.** We do not trust vibes; we trust green checks. The bar
is deliberately high and non-negotiable.

Every change must satisfy, with **zero** exceptions merged to the main branch:

- **Strict TypeScript.** `strict: true` plus `noUncheckedIndexedAccess`,
  `exactOptionalPropertyTypes`, `noImplicitOverride`, and friends. `any` is a code
  smell that must be justified in a comment or eliminated. Prefer `unknown` +
  narrowing. Public API is fully and precisely typed — the types *are* the docs.
- **Linting & formatting clean.** A single fast toolchain (Biome, or ESLint-flat +
  Prettier if a plugin gap forces it). No warnings. Formatting is never debated;
  it is automated.
- **Unit tests.** Every behavior is covered. New behavior ships with tests in the
  same change. Bugs are fixed *test-first*: a failing reproduction, then the fix.
- **Regression corpus.** Every real-world file/edge case we learn about (especially
  from the harvested upstream issues) becomes a permanent fixture-backed test. We
  never fix the same bug twice.
- **Type-level tests** for the public API surface (`expectTypeOf` / `tsd`-style)
  so refactors can't silently degrade the developer experience.
- **Static analysis & supply-chain hygiene.** Dependency audit is part of CI and is
  *expected to stay green* — a large part of why we forked was upstream's rotting
  transitive dependencies. We keep the dependency tree small, modern, and clean.
- **Coverage is watched, not gamed.** Coverage thresholds are a floor, not a
  target; the goal is *meaningful* coverage of behavior, not line percentage.

**Definition of Done:** typed, linted, tested (unit + regression where relevant),
documented at the API surface, dependency-clean, and green in CI. Anything short of
that is not done — it is in progress.

---

## 3. How agents work here

- **Bias to action.** When you have enough information to make a sound decision,
  make it and proceed. Do not ask permission for things this document already
  authorizes (deleting legacy, breaking APIs, modernizing tooling, rewriting bad
  comments).
- **Small, coherent, reviewable changes.** Each change does one thing, is fully
  green, and leaves the tree in a better state than it found it. No half-migrations
  landed on main.
- **Leave a trail, not a mess.** Record non-obvious decisions and assumptions in the
  relevant doc (`docs/architecture.md`, ADRs under `docs/decisions/`, or the change
  description) so the next agent — or the human — can audit and reverse them.
- **Verify, don't assume.** Prefer reproductions and tests over reasoning about what
  "should" work. The OOXML format is full of surprises; the corpus is how we tame
  them. `docs/agent-correctness-playbook.md` is the dispatch table — *what you are
  doing* → *the check that proves it correct* → *the exact command* — so you never
  have to rebuild that decision tree.
- **When you must ask, batch it and make it count.** Only escalate to the human for
  decisions that are genuinely theirs: irreversible/outward-facing actions
  (publishing, naming/branding, license), or a true fork-in-the-road where the
  options are strategically divergent and you cannot pick a clearly-better one.
- **Security- and correctness-first, always.** A spreadsheet library parses
  untrusted input. Treat every parser path as hostile-input-facing. No unbounded
  allocation, no zip-bomb naïveté, no eval-shaped surprises.

## 4. Code & comment standards

- **Comments explain *why*, never *what*.** The code says what. If a comment
  restates the code, delete it. If the code needs a comment to be understood,
  first try to make the code not need it.
- **Delete dead code and stale comments on sight.** Do not preserve them "just in
  case." Git remembers.
- **Names carry meaning.** Precise, honest names over short or clever ones. The
  public API is a UX surface for both humans and agents — optimize for
  discoverability and for correct use being the easy path.
- **No cargo cult.** Every dependency, config line, build step, and abstraction
  earns its place or is removed. Simplicity is a feature.
- **Modern idioms only.** ESM, `async`/`await`, immutable-by-default, narrow
  interfaces. No callbacks-as-API, no `var`, no CommonJS-isms in source.

---

## 5. The one rule that overrides convenience

> If keeping something would make the library worse but easier, and removing it
> would make the library better but harder — **we do the harder, better thing.**

That is the whole point of the fork.

---

_See `docs/architecture.md` for how these principles are realized — the corpus contract,
the module layout, and the working agreements; `docs/agent-correctness-playbook.md` for
the situation → check → command dispatch table; and `docs/decisions/` for the ADRs._
