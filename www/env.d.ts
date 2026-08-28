/**
 * What the bundler resolves and the typechecker does not.
 *
 * `noUncheckedSideEffectImports` is on across this repo, so `import './style.css'` has to
 * name a module tsc knows about or the theme fails the type gate. Declaring the shapes here
 * rather than switching the flag off keeps the guarantee everywhere else: an import of a
 * module that genuinely does not exist is still an error.
 */

declare module '*.css';

/** Built by `www/scripts/virtual-modules.ts`: sample id to highlighted markup. */
declare module 'virtual:sample-sources' {
  const sources: Readonly<Record<string, string | undefined>>;
  export default sources;
}

/** Built by `www/scripts/virtual-modules.ts`: what the home page quotes, counted from the repo. */
declare module 'virtual:site-facts' {
  interface SiteFacts {
    /** Where the site is served from, so a component can prefix a url it builds itself. */
    readonly base: string;
    readonly runtimeDependencies: number;
    readonly runtimeDependencyNames: string;
    readonly corpusCases: number;
    readonly decisionRecords: number;
    readonly specNotes: number;
    readonly guidePages: number;
    readonly docsPages: number;
    /** The guide's first sample, highlighted. Not a copy of it: the same block. */
    readonly quickStartHtml: string;
  }
  const facts: SiteFacts;
  export default facts;
}
