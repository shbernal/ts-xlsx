# Errors

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `AuthoringError`

<sub>class</sub>

Thrown when the caller describes a document that cannot exist: a pivot table with no row field, a
table whose columns do not span its range, a merge that overlaps another, a workbook with no
worksheets. The document model, not a single argument, is what is wrong.

**Where the line falls against native errors.** A single scalar that is out of range, unparseable,
or the wrong type stays a native `RangeError` / `SyntaxError` / `TypeError` — those types exist for
exactly that, every caller already reads them, and wrapping them would make this taxonomy a
re-implementation of the language's. `AuthoringError` starts where a *composite* is internally
inconsistent, or contradicts something a workbook can express. `getColumn(0)` is a `RangeError`;
a table that names a column twice is an `AuthoringError`.

It is always the calling code that is wrong, never the input file — a malformed file raises a
`'malformed-input'` error instead.

```ts
class AuthoringError extends XlsxError {
  override readonly name = 'AuthoringError';
  override readonly code = 'authoring';
}
```

---

### `InternalError`

<sub>class</sub>

Thrown where an invariant the library itself maintains turns out not to hold — typically an index
that a preceding pass proved to be in range, re-checked because `noUncheckedIndexedAccess` makes
the possibility of `undefined` explicit and casting it away would be worse.

No caller can provoke one, so it is not a failure mode to handle: seeing it means the bug is ours.
It exists as a distinct type so that "unreachable" is *stated* rather than implied by a bare
`Error`, which reads identically to a throw nobody has classified yet.

It is the one class in the taxonomy that rewrites its own message, appending `REPORT_NOTICE`
below the invariant that broke. The constructor is where that lives so a throw site added later
inherits it — the alternative, a notice pasted at each of the throw sites, is one every future
site can forget. Every other class leaves `message` exactly as given: `'malformed-input'` is a
routine outcome for a library that reads untrusted files, and a "report this" banner on each
corrupt input would train callers to ignore the one banner that always means something.

```ts
class InternalError extends XlsxError {
  override readonly name = 'InternalError';
  override readonly code = 'internal';
}
```

---

### `XlsxError`

<sub>class</sub>

The common ancestor of every error this library raises deliberately. `catch (e) { if (e instanceof
XlsxError) … }` is the one-line answer to "was that us?", and `code` then says which kind of
failure it was without naming a single subclass.

Abstract on purpose: a failure always has a kind, so there is never a reason to throw the base.
Every subclass fixes `code` to a literal, which makes the class hierarchy a discriminated
union — narrowing on `error.code` narrows the type.

The constructor is inherited from `Error`, so every subclass accepts `{cause}`; layers that wrap a
lower-level failure are expected to pass it rather than flatten it into the message.

```ts
class XlsxError extends Error {
  abstract readonly code: XlsxErrorCode;
}
```

---

### `XlsxErrorCode`

<sub>type</sub>

What kind of failure an [`XlsxError`](./errors.md#xlsxerror) reports. This is the branch most callers want, and it is
deliberately coarse: the four answers are the four different things a caller would *do* next.

- `'unsupported-format'` — the input is not a container this library reads at all (a legacy `.xls`,
  a blob that is not a spreadsheet). Nothing is wrong with the file; it is the wrong file *for us*.
- `'malformed-input'` — a part we do read is corrupt or does not conform to its specification. The
  file is broken, or hostile.
- `'authoring'` — the caller described a document that cannot exist. The bug is in the calling code.
- `'internal'` — an invariant this library maintains did not hold. It should be unreachable; if it
  fires, the bug is ours.

There is deliberately no "not implemented yet" code. Every candidate for one turned out to be an
unreachable exhaustiveness guard (so: `'internal'`), and the one genuine feature gap — a binary
`.xlsb` cannot be row-streamed — is already reported by [`UnsupportedFormatError`](./opc-errors.md#unsupportedformaterror)'s `format`
branch. A code with no throw site would be a promise the library does not keep.

**Which of these is worth reporting upstream.** `'internal'` always is, and says so at runtime.
`'unsupported-format'` and `'malformed-input'` are worth reporting when the file in hand opens
cleanly in Excel — that combination means we are the ones who cannot read it, which is a gap, not
a corrupt input. `'authoring'` is worth reporting only if the document it refused is one a real
workbook can express. See `skills/ts-xlsx-upstream` for how to file one.

```ts
type XlsxErrorCode = 'unsupported-format' | 'malformed-input' | 'authoring' | 'internal';
```
