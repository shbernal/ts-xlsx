# Vba Project Editor

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `addVbaReference`

<sub>function</sub>

Add a registered (COM type-library) reference to an existing `vbaProject.bin`, returning new bytes
that carry every existing module, reference, and host-info record unchanged. It grows the project's
`dir` stream by one `REFERENCENAME` + `REFERENCEREGISTERED` record pair, positioned immediately before
`MODULES_COUNT` (references have no count field of their own — `MODULES_COUNT` simply marks where the
reference array ends). It needs no change to `PROJECT`/`PROJECTwm`: a real Excel-authored `PROJECT`
stream carries no `Reference=` line at all — references live only in `dir` (confirmed against a genuine
Excel-authored project).

```ts
function addVbaReference(bin: Uint8Array, ref: VbaLibraryReference): Uint8Array;
```

**Throws** — [`VbaParseError`](./vba-errors.md#vbaparseerror) if `bin` is not a parseable VBA project (validated before any edit).
**Throws** — [`VbaAuthorError`](./vba-errors.md#vbaauthorerror) if any field of `ref` is invalid (see [`VbaLibraryReference`](./vba-project-editor.md#vbalibraryreference)), or the
assembled reference text has a character the project's code page cannot represent.

---

### `removeVbaModule`

<sub>function</sub>

Remove a standard module from an existing `vbaProject.bin`, returning new bytes that carry every
remaining module, reference, and host-info record unchanged. It drops the module's `VBA/<name>`
stream, its MODULE record block in `dir` (decrementing `MODULES_COUNT`), and its `Module=`/`Class=` +
workspace lines in `PROJECT`/`PROJECTwm`.

Only `procedural` and `class` modules can be removed this way — removing a `document` module (e.g.
`ThisWorkbook`) or a `designer` module (a UserForm) would leave the host referencing code that no
longer exists, since their names are tied to a worksheet/workbook `codeName` or a designer storage
this project-level primitive has no visibility into. Editing such a module's code-behind is a job for
the offline `tools/vba-compiler` (in-place mode), which drives the real host.

```ts
function removeVbaModule(bin: Uint8Array, name: string): Uint8Array;
```

**Throws** — [`VbaParseError`](./vba-errors.md#vbaparseerror) if `bin` is not a parseable VBA project (validated before any edit).
**Throws** — [`VbaAuthorError`](./vba-errors.md#vbaauthorerror) if `name` is not in the project, or names a `document`/`designer` module.

---

### `VbaLibraryReference`

<sub>interface</sub>

A registered (COM Automation type-library) reference to add to an existing VBA project — the shape of
a real "add a reference to Microsoft Scripting Runtime" call. Project references (to another VBA
project) and control references (to an ActiveX control library) are out of scope — see
[`addVbaReference`](./vba-project-editor.md#addvbareference).

```ts
interface VbaLibraryReference {
  /**
   * The reference's namespace name in the VBA editor — what a qualified reference like
   * `Scripting.Dictionary` resolves through. Must be a valid VBA identifier, at most 31 characters, as
   * real type libraries use (e.g. `Scripting`, `Office`, `stdole`).
   */
  readonly name: string;
  /**
   * The friendly name shown in the References dialog, e.g. `Microsoft Scripting Runtime`. Real projects
   * usually keep this distinct from {@link name}; defaults to {@link name} if omitted.
   */
  readonly displayName?: string;
  /** The type library's GUID, e.g. `{420B2830-E718-11CF-893D-00A0C9054228}` (braces optional). */
  readonly guid: string;
  /** The type library's major version — an integer in `[0, 0xFFFF]` ([MS-OVBA] `LibidMajorVersion`). */
  readonly majorVersion: number;
  /** The type library's minor version — an integer in `[0, 0xFFFF]` ([MS-OVBA] `LibidMinorVersion`). */
  readonly minorVersion: number;
  /**
   * The type library's LCID — an integer in `[0, 0xFFFFFFFF]`. Defaults to `0` (locale-neutral), the
   * overwhelming common case (every reference in a real project observed while building this had `0`).
   */
  readonly lcid?: number;
  /** Absolute Windows path to the type library file, e.g. `C:\Windows\System32\scrrun.dll`. */
  readonly path: string;
}
```
