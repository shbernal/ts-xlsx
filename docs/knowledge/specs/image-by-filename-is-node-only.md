# a.Readfile is not a function when exporting excel with images

## Adding an image by filesystem path is a Node-only capability

### Observed problem
The image-add API accepts several source shapes: a filesystem path (`filename` + `extension`), a raw binary buffer (`buffer`), and a base64 data string (`base64`). Only the buffer and base64 shapes work in a browser. The filesystem-path shape reads the file from disk, which has no meaning in a browser and fails with a confusing, often minified error (`readFile is not a function`) at write time rather than at the point the bad source was supplied.

### Desired behavior
- Supplying an image by filesystem path in an environment without filesystem access should fail **fast and clearly**: the error must name the actual problem ("adding an image by file path is only supported in Node; supply `buffer` or `base64` instead") and should surface at the point the source is registered, not deep inside serialization.
- Buffer and base64 image sources must remain fully supported in browser bundles and must not pull the filesystem module into the bundle.
- The type/API surface should make the environment constraint discoverable — e.g. the filesystem-path source is documented as Node-only, and browser builds do not silently ship a broken code path.

### Prior art / workarounds users converged on
- Fetch the image over HTTP as an ArrayBuffer and pass it as `buffer`.
- Convert the image to a base64 data string and pass it as `base64`.
Both are reliable across environments today; the filesystem-path source is the only one that is environment-specific.

### Open questions
- Should the modern API keep a single overloaded image-source shape, or split "read from path" into an explicit Node-only helper that returns a buffer, leaving the core `addImage` filesystem-free by construction?
- Should there be runtime detection (throw a clear error) in addition to build-time separation, or is build-time separation sufficient?
