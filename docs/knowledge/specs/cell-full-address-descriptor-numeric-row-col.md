# Bad type in Cell.fullAddress

## Cell full-address descriptor: row and col are numbers

### Desired behavior
When a cell exposes a full-location descriptor (sheet name + address + row + column), the row and column fields MUST be 1-based integers matching the cell's position, and the public TypeScript type MUST declare them as `number`.

Example runtime shape for the top-left cell:

```
{ sheetName: "Sheet1", address: "A1", row: 1, col: 1 }
```

### Prior art
The legacy library shipped a declaration that typed these as the `Row` and `Column` object types:

```
fullAddress: {
  sheetName: string;
  address: Address;
  row: Row;    // wrong — runtime value is a number
  col: Column; // wrong — runtime value is a number
};
```

The runtime value has always been plain numbers; the cell's own `row`/`col` accessors return numbers as well. The declaration was simply out of sync with reality.

### Guidance for the rewrite
- Whatever we name the per-cell location descriptor in ts-xlsx, its `row` and `col` must be typed as `number` (1-based), and the address as the A1-style string type.
- Cover this with a type-level test (`expectTypeOf`) so the descriptor's `row`/`col` can never silently regress to an object type again.
- Ensure runtime and type agree: a decoded `"A1"` yields `row: 1, col: 1`.

### Open questions
- Naming: whether ts-xlsx keeps a single "full address" descriptor or splits sheet-scoped vs. cell-scoped location info.
- Whether to also expose the object-typed row/column handles under a distinctly named field, so numeric coordinates and rich handles never share a field and confuse consumers.
