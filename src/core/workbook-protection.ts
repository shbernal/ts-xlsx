// Workbook-level protection: the `<workbookProtection>` element in CT_Workbook that locks the
// workbook's *structure* (sheets cannot be added, deleted, reordered, or unhidden) and/or its
// *windows*. It is distinct from worksheet-level `<sheetProtection>` — that guards a single sheet's
// cells; this guards the workbook shell. A weak integrity signal, but a real one: dropping it on a
// passthrough save silently unlocks a file the author locked, so the model preserves it faithfully.

/**
 * The password/agile-hash attributes a `<workbookProtection>` element may carry, preserved verbatim
 * so a hash-guarded workbook round-trips without the library ever interpreting the credential. Both
 * the legacy 16-bit `*Password` hashes and the modern agile `*AlgorithmName`/`*HashValue`/
 * `*SaltValue`/`*SpinCount` quartets appear here, for the structure guard (`workbook*`) and the
 * revisions guard (`revisions*`). The reader accepts only these names, so a hostile or unknown
 * attribute is never echoed back into the output.
 */
export const WORKBOOK_PROTECTION_CREDENTIAL_ATTRS: readonly string[] = [
  'workbookPassword',
  'workbookAlgorithmName',
  'workbookHashValue',
  'workbookSaltValue',
  'workbookSpinCount',
  'revisionsPassword',
  'revisionsAlgorithmName',
  'revisionsHashValue',
  'revisionsSaltValue',
  'revisionsSpinCount',
];

/**
 * A workbook's structure/window protection. The three lock flags each default to `false` (absent),
 * matching OOXML: an omitted attribute leaves that aspect unlocked. The optional {@link credentials}
 * bag carries the opaque password attributes verbatim — the library never verifies a password, it
 * only refuses to lose one.
 */
export interface WorkbookProtection {
  /** Lock the workbook structure — no adding, deleting, reordering, or unhiding sheets. */
  readonly lockStructure?: boolean;
  /** Lock the workbook window geometry. */
  readonly lockWindows?: boolean;
  /** Lock the revision-tracking state. */
  readonly lockRevision?: boolean;
  /** Preserved password/agile-hash attributes, keyed by their OOXML attribute name. */
  readonly credentials?: Readonly<Record<string, string>>;
}
