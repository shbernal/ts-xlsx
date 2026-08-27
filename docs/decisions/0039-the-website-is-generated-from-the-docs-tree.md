# ADR 0039: The website is generated from the docs tree, and drift is a build error

**Status:** Accepted 2026-08-27

## Context

`docs/` is 247 markdown files: three hand-written top-level pages, 44 API pages generated
from the public types (ADR-0006), 39 records under `decisions/`, and 161 behaviour notes
under `knowledge/specs/` (ADR-0007). It is written for someone standing in a checkout. It
carries no frontmatter, it links to its neighbours by relative path, and it links out to
files a website would not publish, such as `CLAUDE.md` and `test/corpus/`.

A published site needs a different shape from all three of those: a title and description
per page, links that resolve to routes, and a navigation order. The question is where that
shape is applied.

Two arrangements were available, both proven in sibling repositories. Author the pages
inside the site and let the repository read them from there; or keep the repository tree
authoritative and generate the site from it.

## Decision

`docs/` is the source. `www/` is the site, and `www/docs/` is a generated mirror of `docs/`
that is git-ignored and rewritten whole on every build. Nothing in the mirror is edited,
and each page says so in a comment naming the file it came from.

**Drift is an error, never a warning.** `www/scripts/docs-source.ts` reads the tree,
validates it and throws on: a page no navigation group reaches; a manifest entry with no
file; a page claimed twice; a link inside `docs/` that resolves to nothing; a link leaving
`docs/` that points at a path which does not exist; a page with no `#` heading; and markdown
the site's Vue compiler cannot parse. `node www/scripts/check.ts` is that validation with
nothing written, and it runs in the `invariants` gate of every `verify` in about a second.

**Navigation is a manifest, and coverage is total by construction.** `docs/docs.json`
carries the reading order and nothing else. A group either names its pages, which is what
puts three overview pages in the order a reader should meet them, or declares a `tree`,
which enumerates the rest of a directory in sorted order. Every `.md` under `docs/` must be
claimed exactly once. That is what lets the manifest stay 30 lines while 247 pages remain
reachable: a page added to `knowledge/specs/` appears in the sidebar, and a page added to a
directory no group covers fails the build rather than being published unlinked.

**No frontmatter schema, against both sibling repositories' practice.** Titles come from the
first `#` heading, which all 247 pages already have, and a page without one throws.
Descriptions come from the first paragraph that contains a sentence terminator, which is the
one test that separates an opening sentence from the three things that reliably precede one
here: a `Cluster: images` metadata line, a `<sub>interface</sub>` marker on a generated API
page, and a status line of dates and separators. One page in 247 never reaches prose, and it
gets no description rather than an invented one; `check.ts` prints that list so the number
cannot creep. Adding a schema to 247 files to serve a `<title>` tag would be a migration paid
for by every page written afterwards, and it would buy checkability that the derivation
already has.

**A directory's index keeps its filename in the mirror.** `docs/api/README.md` is mirrored as
`www/docs/api/README.md` and served at `/docs/api/` through VitePress's `rewrites`, rather
than being renamed to `index.md`. That keeps a page's path below the site root identical to
its path in the repository, which is what lets the theme's edit link point at the tracked
file with no mapping table of its own.

## Consequences

**Two dev dependencies and 158 packages.** `vitepress` and `vue`, neither of which the
library ships. esbuild's postinstall is approved in `pnpm-workspace.yaml` because Vite cannot
transform anything without its platform binary.

**One more gate.** `site:check` joins `invariants`, which is where the sub-second whole-tree
checks live. It is not the site build: the build is half a minute and pulls in Vite, and
everything it would catch about `docs/` this catches first, with a message that names the
line rather than the place a parser gave up.

**Two constructs are now refused in `docs/`.** A bare `<Tag>` in prose, which is invisible on
GitHub today and fails the site build; and an inline code span wrapped across a newline where
the continuation line starts with `<`, which markdown keeps as code and the Vue compiler does
not. Three pages carried one of these and were reflowed. Inline code is rendered with `v-pre`
so a `{{ ... }}` inside it stays literal, which the note on template placeholders needs.

**The site gets no size budget.** `scripts/size-budget.ts` guards what a consumer installs,
where a byte is a byte someone did not ask for. A documentation site is not installed, and
its largest artefact by far is a 1.5 MB local search index over 247 pages that is fetched
only when a reader opens search. A budget there would report on the size of the
documentation, which is not a thing we want to hold down. Declined on purpose. Vite's own
500 kB chunk warning is raised to 2000 kB for the same reason, and only to just above what
that index costs, so the next chunk to cross the line still says so.

**A generated tree is not scanned twice.** charcheck's markdown rule takes `www/*.md`, one
level deep, because `www/docs/` is a copy of a tree the same rule already reads. The mirror
is git-ignored, so nothing else reaches it either.
