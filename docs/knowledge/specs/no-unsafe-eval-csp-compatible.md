# Uncaught EvalError: 'unsafe-eval' is not an allowed source of script (CSP)

## CSP-safe: no eval, no `new Function`, no `Function(...)`

### Desired behavior
The library must run correctly in browser-adjacent environments (web pages, Manifest V3 extensions, Electron renderers) under a strict Content Security Policy — specifically `script-src 'self'` with **no** `'unsafe-eval'` allowance. Loading and using any public entry point (read, write, CSV, zip) must not trigger `EvalError`.

Concretely:
- The shipped source and every bundled artifact must contain **zero** dynamic code evaluation: no `eval`, no `new Function(...)`, no `Function(...)` call constructor, no `setTimeout("string", …)`, no bundler-injected module-globals via `new Function`.
- This must hold for all published entry points and bundles, not just an opt-in "bare" build. There must be no configuration where the default import path pulls in eval-based polyfills.

### Prior art / root causes seen in the wild
The eval usage in the legacy library came entirely from tooling and transitive dependencies, never from spreadsheet logic:
1. Browserify's module-globals injection wraps modules with `new Function` to synthesize `global`/`process` shims.
2. `regenerator-runtime` intentionally escapes accidental strict mode via `Function("r", "regeneratorRuntime = r")(runtime)` — only triggered when the bundle is force-strict, which browserify+babel produced.
3. `declare.js`, reached transitively through the CSV parsing dependency, used `new Function`.
4. The legacy zip dependency's browserify wrapper also emitted `new Function`.

Community workarounds (all unsatisfactory): widen CSP to allow `unsafe-eval` (defeats CSP), alias to a "bare" bundle and import `regenerator-runtime` separately, or `sed` out `"use strict"` from the shipped dist.

### Why this is (mostly) obsolete-by-construction for ts-xlsx
ts-xlsx is a TypeScript-first rewrite targeting native async/await (no regenerator runtime) with a small, modern dependency tree and a modern bundler (no browserify module-globals injection). Choosing an eval-free zip/deflate dependency and an eval-free CSV parser eliminates every historical source of the problem. The remaining work is to make CSP-safety a *guarded invariant* rather than an accident.

### Concrete requirements to carry forward
- Select dependencies (zip/inflate, CSV) that are documented CSP-safe and contain no `new Function`/`eval`. Reject any dependency (direct or transitive) that introduces dynamic evaluation.
- Do not rely on `regenerator-runtime`; target runtimes with native generators/async.
- Add a **build-time guard** in CI that scans every published artifact (all entry points/bundles) for `eval(`, `new Function`, `Function(` call-constructor usage, and fails the build if any is found. This makes CSP-safety a machine-checked green check rather than a promise.

### Open questions
- What is the minimum supported runtime baseline? (Determines whether any transpilation/polyfill is ever needed, which is where eval historically crept in.)
- Do we ship separate browser vs node bundles, and if so, is the CSP scan applied to both? (It should be.)
- Should the CSP guarantee be asserted with an actual headless-browser test under `script-src 'self'` (strongest signal) in addition to the static scan?
