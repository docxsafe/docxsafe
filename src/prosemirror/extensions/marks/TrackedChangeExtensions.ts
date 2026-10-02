/**
 * Tracked Change Mark Extensions — insertion and deletion marks
 *
 * Visible styling (green underline / red strikethrough) is applied by the layout
 * painter via CSS classes. toDOM deliberately carries no inline styles: the hidden
 * ProseMirror DOM is re-parsed on input, and inline text-decoration/color would be
 * read back as underline/strike/textColor marks.
 */

import { createMarkExtension } from '../create';

/**
 * Insertion mark — text added in tracked changes
 * Renders with green color and underline.
 */
export const InsertionExtension = createMarkExtension({
  name: 'insertion',
  schemaMarkName: 'insertion',
  markSpec: {
    attrs: {
      revisionId: { default: 0 },
      author: { default: '' },
      date: { default: null },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: 'span.docx-insertion',
        // Win over generic span/style rules when ProseMirror re-reads the DOM
        priority: 60,
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return {
            revisionId: parseInt(el.dataset.revisionId || '0', 10),
            author: el.dataset.author || '',
            date: el.dataset.date || null,
          };
        },
      },
    ],
    toDOM(mark) {
      return [
        'span',
        {
          class: 'docx-insertion',
          'data-revision-id': String(mark.attrs.revisionId),
          'data-author': mark.attrs.author,
          ...(mark.attrs.date ? { 'data-date': mark.attrs.date } : {}),
        },
        0,
      ];
    },
  },
});

/**
 * Deletion mark — text removed in tracked changes
 * Renders with red color and strikethrough.
 */
export const DeletionExtension = createMarkExtension({
  name: 'deletion',
  schemaMarkName: 'deletion',
  markSpec: {
    attrs: {
      revisionId: { default: 0 },
      author: { default: '' },
      date: { default: null },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: 'span.docx-deletion',
        priority: 60,
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return {
            revisionId: parseInt(el.dataset.revisionId || '0', 10),
            author: el.dataset.author || '',
            date: el.dataset.date || null,
          };
        },
      },
    ],
    toDOM(mark) {
      return [
        'span',
        {
          class: 'docx-deletion',
          'data-revision-id': String(mark.attrs.revisionId),
          'data-author': mark.attrs.author,
          ...(mark.attrs.date ? { 'data-date': mark.attrs.date } : {}),
        },
        0,
      ];
    },
  },
});
