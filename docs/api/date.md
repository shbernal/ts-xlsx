# Date

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `DateEpoch`

<sub>type</sub>

Which date system a workbook counts its serials in: the `date1904` flag of `<workbookPr>`.

Carried as the epoch year rather than as a boolean because it is what every conversion here
actually needs, and a boolean would mean a `? 1904 : 1900` at each of them: five chances to write
the ternary backwards, on a value whose wrongness is silent by construction (a date read under the
wrong system is still a perfectly good date, four years off).

```ts
type DateEpoch = 1900 | 1904;
```
