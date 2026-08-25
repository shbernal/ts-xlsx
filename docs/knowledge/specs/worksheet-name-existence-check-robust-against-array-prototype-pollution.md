# Robust duplicate-worksheet-name detection

Cluster: core-model

Reported as "Error: Worksheet name already exists: My Sheet".

## Desired behavior

Adding a worksheet with a name that does not already exist in the workbook must always succeed, regardless of what other code in the same process has done to built-in prototypes. Adding a worksheet whose name collides, case- and whitespace-normalized per the sheet-name rules, with an existing sheet must fail with a clear error.

Existence should be determined by an explicit boolean predicate, not by coercing the return value of an array search method. Use a membership set or a `some()`-style predicate, or compare a found result explicitly against `undefined`, and never `if (list.find(pred))`.

## Prior art / root cause

An early implementation used `if (worksheets.find(ws => ws && ws.name === name)) throw ...`. That is correct only if `Array.prototype.find` returns `undefined` on no match. Some third-party libraries, including an old montagejs "collections" shim, globally redefine `Array.prototype.find` to return `-1` when nothing is found. Since `-1` is truthy, the guard misfired on every add, throwing "Worksheet name already exists" even for the first sheet in an empty workbook. The bug reproduced across many versions and disappeared when the polluting dependency was upgraded or removed, confirming the library's own logic was the fragile part, not the input.

## Design guidance for the fork

- Treat global built-in prototypes as potentially hostile or polluted. Do not depend on the exact "not found" sentinel of `Array.find` or `indexOf` for correctness-critical branches; prefer predicates that yield an unambiguous boolean, such as `.some(...)` or a `Set`/`Map` lookup with `.has(...)`.
- Prefer maintaining an internal index of used sheet names, normalized, so add and rename are O(1) and never walk a monkeypatchable array method for existence.
- Add a regression test that stubs a polluted `Array.prototype.find` returning `-1` and asserts a unique-named worksheet still adds cleanly, and that a true duplicate still throws.

## Open questions

- Exact normalization for "same name": Excel sheet names are compared case-insensitively and are length- and character-constrained. The collision check and the tests should share one normalization function.
