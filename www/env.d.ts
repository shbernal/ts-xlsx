/**
 * What the bundler resolves and the typechecker does not.
 *
 * `noUncheckedSideEffectImports` is on across this repo, so `import './style.css'` has to
 * name a module tsc knows about or the theme fails the type gate. Declaring the shape here
 * rather than switching the flag off keeps the guarantee everywhere else: an import of a
 * module that genuinely does not exist is still an error.
 */

declare module '*.css';
