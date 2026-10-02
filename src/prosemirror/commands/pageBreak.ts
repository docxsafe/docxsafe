/**
 * Insert a page break (Ctrl+Enter in Word).
 *
 * Splits the paragraph at the cursor and starts the new paragraph on a new
 * page via w:pageBreakBefore, which the layout engine and serializer support.
 */

import { TextSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

export function insertPageBreak(view: EditorView): boolean {
  const { state } = view;
  const { $from, $to } = state.selection;
  if ($from.parent.type.name !== 'paragraph' || !$from.sameParent($to)) return false;

  const tr = state.tr.deleteSelection();
  const splitPos = tr.selection.from;
  tr.split(splitPos);

  // The new paragraph starts right after the split point
  const $new = tr.doc.resolve(splitPos + 2);
  const paragraphPos = $new.before($new.depth);
  const paragraph = tr.doc.nodeAt(paragraphPos);
  if (!paragraph) return false;
  tr.setNodeMarkup(paragraphPos, undefined, { ...paragraph.attrs, pageBreakBefore: true });
  tr.setSelection(TextSelection.create(tr.doc, paragraphPos + 1));
  view.dispatch(tr.scrollIntoView());
  return true;
}
