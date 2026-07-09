# Cannot read property 'F_OK' of undefined

## Browser-safe I/O boundary

### Problem
The file-path I/O methods (read from / write to a path on disk) depend on Node's `fs` module, including `fs.constants.F_OK` for existence checks. When the library is bundled for the browser, `fs` is either absent or a bundler-provided empty stub, so `fs.constants` is `undefined`. Any code path that dereferences `fs.constants.F_OK` throws a cryptic `TypeError: Cannot read property 'F_OK' of undefined` from inside the minified bundle. The user has no way to tell that the actual issue is "filesystem paths are not supported in the browser."

This also affects downstream libraries that wrap the file-path API (e.g. render-from-template helpers that call a path-based writer under the hood): they surface the same opaque crash to their users.

### Desired behavior
- The library must have a clear, documented boundary between environment-agnostic I/O (buffers / streams) and Node-only path-based I/O.
- Buffer/stream-based read and write (produce or consume an in-memory `ArrayBuffer`/`Uint8Array`/stream) must work in the browser with no `fs` dependency reachable at runtime.
- Path-based methods (read/write by filesystem path) are explicitly Node-only. When invoked in an environment without a real filesystem, they must fail fast with an actionable error naming the constraint (e.g. "Filesystem path I/O is not available in this environment; use the buffer/stream API") rather than crashing on an internal `undefined` dereference.
- Ideally the Node-only path methods live behind an import/subpath that a browser bundle does not pull in, so `fs` never ends up in the browser bundle at all and the buffer path has zero Node built-in dependencies.

### Prior art
- Common pattern: a core module with no Node built-ins plus a thin Node adapter exposing `readFile(path)` / `writeFile(path)` layered on top of the buffer API.
- Bundlers used to auto-polyfill Node core modules (webpack 4) and now do not (webpack 5), which is why environment-dependent code that "worked" in older toolchains now surfaces as `undefined` module shims.

### Open questions
- Should path-based methods be a separate entry point (subpath export) or a runtime-guarded no-op that throws? A separate entry point keeps `fs` out of browser bundles entirely and is the cleaner break.
- What is the canonical browser write API name and return type (ArrayBuffer vs Uint8Array vs Blob-friendly)?
- Should the environment guard detect the missing filesystem lazily (on call) or fail at import time in a browser build?
