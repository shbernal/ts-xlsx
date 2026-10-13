# ADR 0044: Doc comments are TSDoc

**Status:** Accepted (2026-10-14)

## Context

The doc comments were mostly TSDoc already: no `{type}` annotations, `{@link}` for references,
`@param`, `@returns`, `@throws` and `@example` as the only standard tags. Nothing said which
dialect they were in, and nothing checked. Run through `@microsoft/tsdoc`, the reference parser,
the 1,909 doc comments across `src/`, `scripts/`, `test/`, `tools/` and `www/` produced 324
findings, all JSDoc habits:

- `@throws {RangeError} …`, 135 of them. To TSDoc the braces open a malformed inline tag.
- `@param name description` with no hyphen, 19 of them.
- `@param options.key` for one field of an options object, 6 of them. TSDoc has no form for this.
- A code span wrapped across a line break, which TSDoc does not allow, and a bare `<`, `>`, `{` or
  `}` in prose that TSDoc takes for markup. About 20 in all.

## Decision

Doc comments are TSDoc. `scripts/check-tsdoc.ts` (`pnpm run tsdoc:check`, in the `invariants`
gate) parses every comment TypeScript attaches to a declaration in any of the three typechecked
projects and fails on every TSDoc diagnostic, including an undefined tag. It reads the comments
from the compiler's AST, so a `/**` inside a string is not mistaken for one.

`@microsoft/tsdoc` is a devDependency. It has no dependencies of its own. `@microsoft/tsdoc-config`
was not taken, because it pulls in a JSON schema validator to read a file whose only content here
is one tag definition.

Tags beyond the standard set are declared in `tsdoc.json`, the file editor plugins and API
Extractor read. Today that is `@unpublished`, the block tag `check-public-types.ts` honours.

Two conventions follow from the parsers, not from taste:

- **`@throws` opens with the error type as a plain word:** `@throws RangeError if …`. TSDoc
  suggests `@throws {@link RangeError}`, but TypeScript still reads braces after `@throws` as a
  JSDoc type expression. `{@link …}` is not one, so the parse eats the rest of the comment and
  editor hovers show a lone `{`. A plain first word survives both parsers. `gen-docs.ts` treats it
  as the subject, links it when it can, and fails on a `@throws` that does not open with one.
- **An options object documents its fields on the type**, as property doc comments, in place of
  `@param options.key`. Hovers show these on the property, which is where a caller is looking.

## Consequences

- A JSDoc-ism fails `verify --full` and CI with a file and line, not a review comment.
- The generated reference renders `@param name - description` without the hyphen, and a `@throws`
  followed by a comma no longer gets a stray space before it.
- The check takes about 0.8 s. It opens the same three projects as `typecheck`.
- ADR 0006 still calls the comments "JSDoc" where it names TypeScript's API, which keeps that name.
  The prose convention is this ADR.
