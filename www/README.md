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
    theme/
      index.ts      extends the default theme; the one place components are registered
      style.css     the palette and the lattice
  scripts/
    repo.ts         package.json facts: name, description, repository and blob urls
    docs-source.ts  reads and validates docs/; throws on any drift
    sync-docs.ts    writes www/docs/, the mirror
    check.ts        the drift gate: the validation with nothing written
  playground/       what the playground does, as pure modules with a test beside each
  index.md          the home page
  docs/             GENERATED from docs/. Git-ignored. Never edit a file here.
```

## Working on it

```bash
pnpm run site:dev        # sync, then hot-reload at http://localhost:5173/ts-xlsx/
pnpm run site:build      # sync, then what CI publishes
pnpm run site:preview    # serve the built site
pnpm run site:check      # the drift gate, on its own
pnpm run test:site       # the playground's own tests
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
- **The default theme is extended, never replaced.** The documentation is the bulk of the
  site and the default layout is what makes 248 pages navigable. Colour is expressed through
  VitePress's own custom properties; a class name of theirs is reached for only where nothing
  else will do, and only for something cosmetic enough that a rename costs a background
  rather than a layout.

## The playground

`playground/` holds everything the playground does and none of how it looks: the sample
workbooks as builder functions, the write/read/round-trip lanes, the emitted package as a
listable part tree, a worksheet as a bounded grid model, and the three quantities the page
puts on screen. Every module is pure, every one has a test beside it, and none of them names
a DOM type. That is what makes the components boring, which is the point: a component here
cannot be typechecked past its markup, so anything that could be wrong belongs below it.

Two rules the modules keep, because breaking either would put a lie on the page:

- **A lane returns a result, it does not throw.** A thrown error in a page is a blank panel,
  and a reader who learns that this library refuses their file by name has learned something
  true about it.
- **A timing measures the library call and nothing else.** A number that includes the
  caller's own work is a number that lies.

## The look

The motif is the grid, because a spreadsheet already is one: a rectangular lattice with
typographic content, which is also what a page layout is. Four rules follow from that, and
they are here rather than in the stylesheet because three of them have nothing to style yet.

- **The lattice is drawn, never fetched.** Two repeating gradients, at the pitch and weight
  of a cell. `--xlsx-lattice` is the recipe; the home page is where it is used.
- **A number is set like a numeric cell.** Monospace, tabular figures, right-aligned in its
  container. That covers the counts and the byte sizes a page quotes.
- **A divider is the weight of a cell border**, not a bar.
- **No screenshot of Excel, and no picture of a rendered workbook.** The playground renders a
  real one. A picture would make the same claim with the evidence removed.

One accent and one warning hue, no more, both defined in light and dark with the measured
contrast beside them. A token defined only under `.dark` renders as a browser default in the
other theme, so every one of them is stated twice on purpose.
