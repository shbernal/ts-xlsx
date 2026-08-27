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
    config.ts     the site config; reads facts from package.json, restates none
  scripts/
    repo.ts       package.json facts: name, description, repository and blob urls
  index.md        the home page
```

## Working on it

```bash
pnpm run site:dev        # hot-reloaded at http://localhost:5173/ts-xlsx/
pnpm run site:build      # what CI publishes
pnpm run site:preview    # serve the built site
pnpm run typecheck:site  # tsc over www/**/*.ts
```

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
