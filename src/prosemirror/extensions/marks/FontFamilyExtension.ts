/**
 * Font Family Mark Extension
 */

import { createMarkExtension } from '../create';
import { setMark, removeMark } from './markUtils';
import type { FontFamilyAttrs } from '../../schema/marks';
import type { ExtensionContext, ExtensionRuntime } from '../types';
import { resolveFontFamily } from '../../../utils/fontResolver';

/** Extract a single OOXML font name from a name or CSS stack */
function normalizeFontName(fontName: string): string {
  return fontName.split(',')[0].trim().replace(/^["']|["']$/g, '');
}

export const FontFamilyExtension = createMarkExtension({
  name: 'fontFamily',
  schemaMarkName: 'fontFamily',
  markSpec: {
    attrs: {
      ascii: { default: null },
      hAnsi: { default: null },
      asciiTheme: { default: null },
    },
    parseDOM: [
      {
        style: 'font-family',
        getAttrs: (value) => {
          const fontValue = value as string;
          const firstFont = normalizeFontName(fontValue);
          if (firstFont) {
            return { ascii: firstFont, hAnsi: firstFont };
          }
          return false;
        },
      },
    ],
    toDOM(mark) {
      const attrs = mark.attrs as FontFamilyAttrs;
      const fontName = attrs.ascii || attrs.hAnsi;
      if (!fontName) {
        return ['span', 0];
      }
      const resolved = resolveFontFamily(normalizeFontName(fontName));
      return ['span', { style: `font-family: ${resolved.cssFallback}` }, 0];
    },
  },
  onSchemaReady(ctx: ExtensionContext): ExtensionRuntime {
    return {
      commands: {
        setFontFamily: (fontName: string) => {
          const name = normalizeFontName(fontName);
          return setMark(ctx.schema.marks.fontFamily, {
            ascii: name,
            hAnsi: name,
            asciiTheme: null,
          });
        },
        clearFontFamily: () => removeMark(ctx.schema.marks.fontFamily),
      },
    };
  },
});
