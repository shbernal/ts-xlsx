# Whole-file password encryption of workbook documents

Cluster: security

## Scenario

A user producing spreadsheets with sensitive data wants the generated file itself to require a
password to open — true document-level encryption, not a zip password and not sheet/workbook
structure protection. The distinction matters: sheet protection and zip passwords are trivially
stripped or extracted to disk unencrypted, whereas an Excel-encrypted file keeps its contents
encrypted at rest and demands the password before any data is readable. They also want the inverse:
reading/decrypting such a file back into the library.

> Spec note, not a corpus case: this is a substantial unbuilt cryptographic feature. The durable
> value is the format facts (CFB/MS-OFFCRYPTO), the API shape, and the hostile-input constraints —
> distinct from the much lighter sheet/workbook *protection* flags (see
> `sheet-protection-permits-requested-operations`).

## Desired behavior

- Produce, and ideally consume, **password-encrypted spreadsheet documents at the whole-file level**,
  distinct from OOXML protection flags (removable) and any zip-level password.
- **Format facts**: Excel encryption does not encrypt inside the zip. The entire OOXML zip package is
  placed as an encrypted stream inside a **CFB / OLE2 (Compound File Binary)** container per
  ECMA-376 / MS-OFFCRYPTO. Two schemes: "Standard" (AES + SHA-1 key derivation, fixed structure) and
  the newer "Agile" (an `EncryptionInfo` XML descriptor selecting cipher/hash/salt/spin count,
  AES-CBC with HMAC integrity). Writing: build the normal `.xlsx` zip, derive the key from the
  password, encrypt the package, emit a CFB container with `EncryptionInfo` + `EncryptedPackage`
  streams. Reading: detect the CFB magic, parse `EncryptionInfo`, derive the key, verify the password
  against the verifier, decrypt back to the zip, parse normally.
- **API shape** (durable, to refine): an encryption descriptor on the write path accepting a password
  (and optionally scheme/cipher, defaulting to Agile + modern AES), and symmetrically a password
  option on read/load that decrypts before parsing and surfaces a **typed error** distinguishing
  wrong password from malformed container.
- **Hostile-input bounded**: this is a hard-input-facing parser path — cap spin counts, bound
  allocations, and use vetted crypto primitives.

## Open questions

- Ship both Standard and Agile for writing, or only Agile for writing while tolerating both on read?
- Where does CFB/OLE2 container support live — a small internal module vs a dependency, given
  supply-chain hygiene goals?
- Gate the crypto weight behind an opt-in entry point so it is not pulled into every bundle?
- Interaction with streaming writers, since encryption needs the full package before sealing.
- Keep the API naming clearly separate from sheet/workbook protection flags.

Related: `sheet-protection-permits-requested-operations`, `cell-protection-locked-flag-and-sheet-protection`,
`unsupported-input-format-typed-error`, `lean-zip-container-strategy`.
