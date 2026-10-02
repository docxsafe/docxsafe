/**
 * Comment Mark Extension — highlights text that has comments
 *
 * Applied to text ranges between commentRangeStart and commentRangeEnd.
 * The comment ID links to the Comment object in the document model.
 */

import { createMarkExtension } from '../create';

export const CommentExtension = createMarkExtension({
  name: 'comment',
  schemaMarkName: 'comment',
  markSpec: {
    attrs: {
      /** Comment ID (matches Comment.id) */
      commentId: { default: 0 },
    },
    inclusive: false,
    parseDOM: [
      {
        tag: 'span.docx-comment',
        priority: 60,
        getAttrs(dom) {
          const el = dom as HTMLElement;
          return {
            commentId: parseInt(el.dataset.commentId || '0', 10),
          };
        },
      },
    ],
    toDOM(mark) {
      return [
        'span',
        {
          class: 'docx-comment',
          'data-comment-id': String(mark.attrs.commentId),
          // No inline styles: they would be re-parsed as highlight marks on input.
          // The visible highlight is painted by the layout painter.
        },
        0,
      ];
    },
  },
});
