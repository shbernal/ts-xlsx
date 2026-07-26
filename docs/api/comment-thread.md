# Comment Thread

<!-- Generated from the public types by `pnpm run docs`. Do not edit by hand. -->

### `Comment`

<sub>interface</sub>

One message of a `CommentThread` — what a single person wrote, once.

```ts
interface Comment {
    readonly id: string;
    readonly author?: Person;
    readonly personId?: string;
    readonly date?: string;
    readonly text: string;
    readonly mentions: readonly Mention[];
}
```

---

### `CommentThread`

<sub>interface</sub>

A conversation anchored to one cell: what was asked, every reply, and whether it was resolved.

```ts
interface CommentThread {
    readonly ref: string;
    readonly resolved: boolean;
    readonly comments: readonly Comment[];
}
```

---

### `Mention`

<sub>interface</sub>

An `@mention` inside a message: who was named, and the run of `Comment.text` that renders as the
mention chip.

The offsets are only meaningful against that exact text — shift either and a spreadsheet app
highlights the wrong words.

```ts
interface Mention {
    readonly person?: Person;
    readonly personId: string;
    readonly mentionId?: string;
    readonly startIndex: number;
    readonly length: number;
}
```

---

### `Person`

<sub>interface</sub>

A registered identity a threaded comment can point at — an author, or someone `@mentioned` in a
message. One `<person>` of the workbook's `xl/persons/person.xml` registry.

A single human legitimately has **several** entries: Excel registers a mentioned identity separately
from that person's authoring identity, with the same `displayName` and `userId` but a
different `id` and a different `providerId`. The id is therefore the only identity —
see `Workbook.getPerson`.

```ts
interface Person {
    readonly id: string;
    readonly displayName: string;
    readonly userId?: string;
    readonly providerId?: string;
}
```
