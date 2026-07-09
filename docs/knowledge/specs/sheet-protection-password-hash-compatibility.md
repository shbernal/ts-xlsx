# Sheet-protection password hashing and cross-application compatibility

Cluster: security

## Scenario

A user protects a worksheet with a password. The library writes a `sheetProtection` element with
a modern hash: `algorithmName="SHA-512"`, a `hashValue`, a `saltValue`, and a `spinCount`. This
is valid and honored by current Excel. But some older spreadsheet applications (and some
consumers) only understand the legacy 16-bit `password="XXXX"` hash attribute and silently ignore
the modern algorithm form — so on those consumers the sheet appears unprotected. The user expects
the protection to be enforced everywhere they open the file.

## The design question

This is not a case of wrong output — the SHA-512 `sheetProtection` the library emits is
spec-valid and enforced by modern Excel. The question is a **compatibility policy**: whether, and
how, to also satisfy older consumers that only read the legacy hash.

## Desired behavior (to decide)

- Protecting a worksheet with a non-empty password writes a `sheetProtection` that records enough
  to enforce it. Decide whether that means:
  1. **Modern only** (current): `algorithmName`/`hashValue`/`saltValue`/`spinCount`. Correct and
     strong for modern Excel; ignored by legacy-only consumers.
  2. **Legacy only**: the 16-bit `password` attribute. Broad compatibility, cryptographically
     weak, deprecated.
  3. **Both**: emit the modern hash and the legacy attribute so every consumer enforces
     *something*. Maximizes compatibility at the cost of shipping the weak hash too.
- Each protection permission passed to `protect()` (objects, scenarios, `selectLockedCells`,
  `formatCells`, `insertRows`, …) is serialized as its corresponding `sheetProtection` flag.
- Protecting with **no** password still writes the permission flags but fabricates **no** hash.
- The protection state round-trips: reading the produced file reports the sheet as protected with
  the same permission flags.

## Open questions

- Which policy (modern / legacy / both) is the fork's default, and is it configurable per call?
- If "both", document explicitly that the legacy attribute is weak and present only for
  compatibility — do not let its presence imply strong protection.
- Worksheet protection is not encryption (the sheet data is still readable in the zip); the API
  docs must not imply confidentiality. This belongs alongside any future workbook-encryption work.

## Prior art

The modern `sheetProtection` hash (SHA-512 + salt + spin count) is what current Excel writes and
reads; the legacy 16-bit hash predates it and is what older applications key on. The gap is purely
which forms are emitted.
