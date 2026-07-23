# Vba Errors

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `VbaAuthorError`

<sub>class</sub>

Thrown when authoring a VBA project (synthesizing a `vbaProject.bin` from module source) is asked to
produce something that cannot be encoded to a well-formed container — a stream name longer than the
[MS-CFB] 31-character limit, a duplicate stream name, or a project so large it would exceed the
writer's single-header DIFAT bound. This is a caller-side contract violation, distinct from
`VbaParseError` (which reports a malformed blob *read* from an untrusted file).

```ts
class VbaAuthorError extends Error {
  override readonly name = 'VbaAuthorError';
}
```

---

### `VbaParseError`

<sub>class</sub>

Thrown when a VBA project (`vbaProject.bin`) is present but cannot be parsed — a malformed
compound file, a corrupt compressed stream, or a `dir` record that does not conform to [MS-OVBA].
A workbook with no macros never produces this: `Workbook.vbaProject` is `undefined` instead.

The parser treats the blob as hostile input (a spreadsheet library parses untrusted files), so every
malformed structure fails closed with this error rather than crashing, hanging, or over-allocating.

```ts
class VbaParseError extends Error {
  override readonly name = 'VbaParseError';
}
```
