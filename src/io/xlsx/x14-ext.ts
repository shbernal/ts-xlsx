import {X14_NS} from './namespaces.ts';

/**
 * Wrap an x14 extension body in the `<ext>` envelope every worksheet- and workbook-level x14
 * extension shares: a `uri` scoping the extension to its feature (a fixed GUID from
 * {@link namespaces.ts}) plus the `xmlns:x14` binding the extension namespace on the element itself —
 * Excel declares it per `<ext>`, never at the part root. Callers needing an `<extLst>` wrapper add it
 * around the result; the producers the worksheet serialiser gathers return this bare so several exts
 * compose under a single shared `<extLst>`.
 */
export function x14Ext(uri: string, body: string): string {
  return `<ext uri="${uri}" xmlns:x14="${X14_NS}">${body}</ext>`;
}
