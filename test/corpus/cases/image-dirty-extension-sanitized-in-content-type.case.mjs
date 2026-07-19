// Cluster: images
//
// Real-world scenario: a caller derives an image's extension from a URL and passes it to addImage.
// When the URL carries query parameters (e.g. a photo link ending in ".png?alt=media&token=…"), the
// extension string is not a clean image extension. Adding the image succeeds, but the media/drawing
// serializer threads the dirty string straight into [Content_Types].xml — producing a <Default>
// whose Extension attribute is "png?alt=media&token=…" and a ContentType of "image/png?alt=media&…".
// A content-type Extension must be a bare token (letters/digits); the query-string garbage yields an
// invalid content-type declaration, so strict consumers refuse to open the file (and in some builds
// the write path crashes later with an opaque "Cannot read property 'name' of undefined", far from
// the addImage call). A dirty extension must be sanitized to a clean token before it reaches the
// package.
//
// This is distinct from a missing/undefined extension (which yields an extension-less or
// "image/undefined" Default): here the extension is non-empty but carries invalid characters, which
// passes a "has an Extension / not image/undefined" check yet is still malformed.

/** @typedef {{ name: string, baseline: 'pass'|'fail', expect: (api: any, assert: any) => Promise<void>|void }} Behavior */

const SPEC = {
  sheets: [{name: 'S', images: [{range: 'B2:C3', extension: 'png?alt=media&token=abc123'}]}],
};

export default {
  id: 'image-dirty-extension-sanitized-in-content-type',
  provenance: {source: 'upstream-issue'},
  cluster: 'images',
  description:
    'An image whose extension carries URL query-string characters must be sanitized to a bare token ' +
    'before it reaches the content-type declaration — every image <Default> Extension is ' +
    'alphanumeric, and no ContentType embeds the query string — so the package is valid.',

  /** @type {Behavior[]} */
  behavior: [
    {
      name: 'every image content-type Default extension is a bare alphanumeric token',
      baseline: 'pass',
      async expect(api, assert) {
        const {contentTypeDefaults} = await api.inspectPackage(SPEC);
        const imageDefaults = contentTypeDefaults.filter(
          d => d.contentType && d.contentType.startsWith('image/')
        );
        assert.ok(imageDefaults.length >= 1, 'the image produces at least one image content-type Default');
        const dirty = imageDefaults.filter(d => !/^[A-Za-z0-9]+$/.test(d.extension || ''));
        assert.deepStrictEqual(
          dirty,
          [],
          `image content-type extensions must be bare tokens; got ${JSON.stringify(dirty)}`
        );
      },
    },
    {
      name: 'no image content-type embeds the URL query string',
      baseline: 'pass',
      async expect(api, assert) {
        const {contentTypeDefaults} = await api.inspectPackage(SPEC);
        const leaked = contentTypeDefaults.filter(d => /[?&=]/.test(d.contentType || ''));
        assert.deepStrictEqual(leaked, [], `no ContentType may carry query-string characters; got ${JSON.stringify(leaked)}`);
      },
    },
  ],
};
