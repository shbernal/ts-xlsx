# www/ (the website)

The published site: https://shbernal.github.io/ts-xlsx/

This one directory is both the VitePress root and the site's application code. The two
sibling projects this site borrows from split those in half, because in each of them the
docs tree *is* the site root, so a component would have landed inside a tree the docs kit
validates. Ours does not: the site root is its own directory, and a split here would buy a
re-export shim and nothing else.

```
www/
  .vitepress/
    config.ts       the site config; reads facts, restates none
  scripts/
    repo.ts         package.json facts: name, description, repository and blob urls
    docs-source.ts  reads and validates docs/; throws on any drift
    sync-docs.ts    writes www/docs/, the mirror
    check.ts        the drift gate: the validation with nothing written
  index.md          the home page
  docs/             GENERATED from docs/. Git-ignored. Never edit a file here.
```

## Working on it

```bash
pnpm run site:dev        # sync, then hot-reload at http://localhost:5173/ts-xlsx/
pnpm run site:build      # sync, then what CI publishes
pnpm run site:preview    # serve the built site
pnpm run site:check      # the drift gate, on its own
pnpm run typecheck:site  # tsc over www/**/*.ts
```

## Where the documentation comes from

The tracked `docs/` tree is the source and `www/docs/` is a mirror of it, rewritten whole on
every build. `docs/docs.json` is the reading order, and every `.md` under `docs/` must be
claimed by exactly one group in it. Titles and descriptions are derived from the prose, so no
page carries frontmatter. Anything that would let the two disagree, a page nothing reaches, a
link that resolves to nothing, a heading that is missing, fails `site:check` rather than
reaching a reader. ADR-0039 records why it is arranged this way.

## The rules this directory follows

- **It is inside the gates.** `lint`, `format`, `source-text:check` and `chars:check` all
  reach `www/`, and `typecheck:site` is a step of the `typecheck` gate. Nothing here is a
  tree the machine does not read.
- **No `.vue` files.** Components are plain `.ts` with `h()` render functions. `tsc` does
  not read single-file components and neither does oxlint, so a template would be the one
  part of this repo nothing checks. We pay verbosity to keep every line inside the tools.
  A markdown page still mounts a component with a tag, and that tag is the only template
  in the site.
- **Repository facts are read, not typed.** Anything true of the repository comes from
  `scripts/repo.ts`, which reads `package.json`.
