/**
 * The default theme, extended in one place.
 *
 * Extended and not replaced: the documentation is the bulk of this site, and the default
 * layout is what makes 248 pages navigable. What this file is for is the stylesheet, and
 * later the components a page can mount. Restyling belongs in CSS, never here.
 */

import type {Theme} from 'vitepress';
import DefaultTheme from 'vitepress/theme';

import './style.css';

export default {
  extends: DefaultTheme,
} satisfies Theme;
