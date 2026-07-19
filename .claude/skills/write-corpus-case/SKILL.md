---
name: write-corpus-case
description: Author an implementation-blind regression corpus case under test/corpus/cases. Use when distilling a bug/reproduction into a corpus case, adding a regression test that must survive the rewrite, writing a *.case.mjs file, extending the adapter contract, or setting a behavior baseline. Pairs with the harvest-triage skill.
---

# Writing a corpus case

The corpus is **the product's spine** (`docs/architecture.md`). A case encodes "correct
behavior" as implementation-blind assertions that run against *any* implementation
through a thin adapter — so it survives the Phase 3 rewrite and proves the new code
is at least as correct as the old, plus everything the old one got wrong. Full
reference: `test/corpus/README.md`.

## The three rules that make a case durable

1. **Implementation-blind.** A case NEVER imports `lib/` (or any implementation). It
   calls only the adapter's contract vocabulary (`api.decodeAddress(...)` etc.). The
   adapter is the *only* file allowed to know how an implementation is shaped.
2. **Describe the real-world scenario, not the source thread.** The durable knowledge
   is "real `.xlsx` files declare whole-column defined names like `$A:$A`", not "issue
   140". Write the scenario into `description` and each behavior `name`. Do **not**
   put upstream issue/PR numbers in that durable text (they die with the fork). A
   number may sit in the optional `provenance` block as a disposable trace only.
3. **`baseline` records what legacy does *today*.** It is not what you wish were true
   — it is the observed current-code result, so the runner can tell a known-open bug
   from a fresh regression without a red build.

## Shape

`test/corpus/cases/<descriptive-slug>.case.mjs` (a slug like
`whole-column-defined-names` — no number prefix) default-exports:

```js
export default {
  id: 'whole-column-defined-names',
  cluster: 'address-decoding',                 // one of the known clusters
  description: 'Defined names referencing whole rows/columns must decode without ' +
    'crashing or leaking undefined/NaN into serialized addresses.',
  provenance: {source: 'upstream-issue'},      // OPTIONAL, disposable — never the identity
  behavior: [
    {
      name: 'decodeRange("$1:$1") — a full-row range — resolves its known row bounds',
      baseline: 'pass',                         // what legacy does TODAY (see below)
      expect(api, assert) {
        const range = api.decodeRange('$1:$1');
        assert.strictEqual(range.top, 1);
        assert.strictEqual(range.bottom, 1);
      },
    },
  ],
};
```

- **`behavior[]`** — each entry is one assertion about *observable* behavior.
  `expect(api, assert)` gets the adapter and Node's strict `assert`; throw to fail,
  return to pass. Keep each behavior single-purpose so a baseline flip is unambiguous.
- **`cluster`** — group by theme (`address-decoding`, tables, styles, streaming,
  pivot, images, conditional-formatting, dates, formulas, csv, types, security/deps).

## Setting the baseline (do this by running, not guessing)

1. Write the behavior with the assertion for *correct* behavior.
2. Run `pnpm run corpus` and read the actual result for that behavior.
3. Set `baseline` to what legacy actually did:
   - legacy **passes** → `baseline: 'pass'` (a **regression lock** — guards a behavior
     that already works).
   - legacy **fails** → `baseline: 'fail'` (a **known-open bug** the rewrite must fix).

The build stays green (`○` known-open is fine); it only goes red on a true
regression (`✗` = baseline `pass`, now failing). A `↑` means legacy started passing —
flip that baseline to `pass`.

## Fixtures

If the behavior needs a sample spreadsheet, promote it out of the disposable harvest
attachments into the durable corpus and reference it by path:

```
test/corpus/fixtures/<case-slug>/<file>.xlsx
```

Load it inside `expect` via a capability the adapter provides (add one if needed —
see below), never by reaching into an implementation's reader directly.

## Extending the adapter contract

If a case needs a capability the contract doesn't have yet:

1. Add the method to **every** adapter under `test/corpus/adapters/` (today: just
   `current.mjs`). Each method returns plain JSON-serializable data — corners,
   dimensions, cell values — never implementation objects.
2. Document it in the contract table in `test/corpus/README.md`.
3. Keep it minimal and behavior-shaped; the vocabulary grows only as cases demand.

When the rewrite lands, a `rewrite.mjs` adapter binds the same vocabulary to the new
code and every existing case runs unchanged — the corpus does not move, the
implementation does.

## Validate

```
pnpm run corpus
```

Green (`✓`) and known-open (`○`) are both healthy. Zero `✗` regressions is the bar.
