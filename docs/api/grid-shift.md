# Grid Shift

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `AxisSplice`

<sub>interface</sub>

One structural edit to an axis: `count` lines removed at the 1-based `start`, with everything
after the removed span moved by `delta`. `delta` is the *net* movement, so a pure insert of two
lines is `{count: 0, delta: 2}` and a pure delete of three is `{count: 3, delta: -3}`.

The four travel as one value because they describe one edit, and because three of them are
`number`: passed positionally, two adjacent ones transposed compiles, typechecks, lints, and moves
a merge, a dropdown, a highlight or a comment anchor onto cells the author never chose -- silently,
since nothing about the resulting file is malformed. A named field cannot be transposed.

```ts
interface AxisSplice {
  /** The axis the lines were spliced on. */
  readonly axis: 'row' | 'col';
  /** 1-based first line of the removed span. */
  readonly start: number;
  /** How many lines were removed. */
  readonly count: number;
  /** Net movement of every line after the removed span. */
  readonly delta: number;
}
```
