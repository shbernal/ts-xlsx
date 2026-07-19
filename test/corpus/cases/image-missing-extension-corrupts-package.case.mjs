// Cluster: images
//
// Real-world scenario: a user embeds an image but supplies a descriptor whose extension is
// absent or not a real image extension (a misspelled property so it resolves to undefined, or
// an empty string). Today the writer threads the bad extension straight through: the media part
// becomes media/imageN.undefined, the drawing relationship targets that filename, and
// [Content_Types].xml gains a malformed <Default> element carrying ContentType "image/undefined"
// with no Extension attribute at all. The result is a structurally invalid .xlsx — a Default
// content-type element MUST carry an Extension, and image/undefined is not a real media type —
// so strict consumers refuse to open it. A missing/invalid image extension must never produce a
// malformed content-type declaration.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

// An image whose extension is missing entirely — the exact shape a mistyped descriptor produces.
const SPEC = {sheets: [{name: 'S', images: [{range: 'B2:C3', extension: undefined}]}]};

export default {
  id: 'image-missing-extension-corrupts-package',
  provenance: {source: 'upstream-issue', repo: 'exceljs/exceljs', ref: 2300},
  cluster: 'images',
  description:
    'An image added with a missing or invalid extension must not produce a package with a ' +
    'malformed content-type declaration — no <Default> element lacking an Extension attribute ' +
    'and no bogus "image/undefined" media type — so the file opens without a repair step.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'every content-type Default declaration carries an Extension attribute',
      baseline: 'pass',
      async expect(api, assert) {
        const {contentTypeDefaults} = await api.inspectPackage(SPEC);
        const extensionless = contentTypeDefaults.filter(d => !d.extension);
        assert.deepStrictEqual(
          extensionless,
          [],
          `no <Default> may omit its Extension; got ${JSON.stringify(extensionless)}`
        );
      },
    },
    {
      name: 'no content-type declares the bogus "image/undefined" media type',
      baseline: 'pass',
      async expect(api, assert) {
        const {contentTypeDefaults} = await api.inspectPackage(SPEC);
        const bogus = contentTypeDefaults.filter(d => /undefined/.test(d.contentType || ''));
        assert.deepStrictEqual(bogus, [], `no bogus media type may be declared; got ${JSON.stringify(bogus)}`);
      },
    },
  ],
};
