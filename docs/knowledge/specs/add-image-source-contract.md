# The image source contract: buffer/path, with URLs handled deliberately

Cluster: images

## Scenario

A user wants to embed an image by pointing at a remote HTTP(S) URL, expecting the library to fetch
the bytes and embed them. Instead the image source is only interpreted as a local file path (Node
only) or an in-memory buffer, so passing a URL as the "filename" silently fails to produce an
embedded picture — no image, no error. The community answer is always the same: fetch the bytes
yourself and pass a buffer.

> Spec note, not a corpus case: the durable question is what an "image source" *is* and how a
> URL-shaped input should fail — a typing and API-contract decision, not a malformed-output bug. The
> Node-only-ness of the filepath source is already noted in `image-by-filename-is-node-only`.

## Desired behavior

The image-adding API has a clear, well-typed contract for what a source is. Today the accepted
sources are effectively a local file path (Node only) and a raw byte buffer; a remote URL is not
supported and passing one produces nothing. Decide deliberately between two positions:

1. **Keep network I/O out of scope** (recommended default) — a spreadsheet library should not fetch
   on the caller's behalf. Precisely type the source as `buffer | filepath`, and **reject a
   URL-shaped path with a typed, actionable error** ("image source looks like a URL; fetch the bytes
   and pass a buffer") instead of silently embedding nothing.
2. **Opt-in URL support** — offer an async helper that accepts a URL, but only via an explicitly
   injected fetcher so the library core stays network-free and the caller controls the request
   (timeouts, auth, SSRF posture).

Either way the failure mode is explicit and the source type is honest, so "pass a URL" is never a
silent no-op.

## Open questions

- Is a URL ever accepted in core, or always via an injected fetcher in an optional helper?
- What exactly counts as "URL-shaped" for the reject path (scheme allowlist vs any `://`)?
- Security: any fetch path is untrusted-input-facing (SSRF, redirect, size limits) — a reason to
  keep it out of core and in a caller-controlled helper.

Related: `image-by-filename-is-node-only`, `image-missing-extension-corrupts-package`,
`load-accepts-arraybuffer-and-typed-arrays`.
