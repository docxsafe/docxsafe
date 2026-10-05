/**
 * Review commands — tracked changes (suggestions), comments, and Word tags
 * (content controls) on the ProseMirror document.
 */

import { Plugin, TextSelection, type EditorState, type Transaction } from 'prosemirror-state';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import { normalizeContentControlTagAttrs } from '../../types/content';
import { SUGGESTION_SKIP_META } from '../plugins/suggestionMode';

// ============================================================================
// TRACKED CHANGES
// ============================================================================

export interface RevisionInfo {
  /** Unique key: `${type}:${revisionId}` */
  key: string;
  type: 'insertion' | 'deletion';
  revisionId: number;
  author: string;
  date: string | null;
  /** Affected text (images/other inline nodes shown as a placeholder) */
  text: string;
  from: number;
  to: number;
}

interface MarkedRange {
  from: number;
  to: number;
}

function collectRevisionRanges(
  doc: PMNode,
  match: (type: 'insertion' | 'deletion', revisionId: number) => boolean
): Map<string, { type: 'insertion' | 'deletion'; ranges: MarkedRange[] }> {
  const result = new Map<string, { type: 'insertion' | 'deletion'; ranges: MarkedRange[] }>();
  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    for (const mark of node.marks) {
      const type = mark.type.name;
      if (type !== 'insertion' && type !== 'deletion') continue;
      const revisionId = mark.attrs.revisionId as number;
      if (!match(type, revisionId)) continue;
      const key = `${type}:${revisionId}`;
      const entry = result.get(key) ?? { type, ranges: [] };
      const last = entry.ranges[entry.ranges.length - 1];
      if (last && last.to === pos) {
        last.to = pos + node.nodeSize;
      } else {
        entry.ranges.push({ from: pos, to: pos + node.nodeSize });
      }
      result.set(key, entry);
    }
    return false;
  });
  return result;
}

/**
 * List tracked changes in document order. Adjacent runs that share a
 * revision ID are reported as one change.
 */
export function getRevisions(doc: PMNode): RevisionInfo[] {
  const revisions: RevisionInfo[] = [];
  let current: RevisionInfo | null = null;

  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    const mark = node.marks.find((m) => m.type.name === 'insertion' || m.type.name === 'deletion');
    if (!mark) {
      current = null;
      return false;
    }
    const type = mark.type.name as 'insertion' | 'deletion';
    const revisionId = mark.attrs.revisionId as number;
    const key = `${type}:${revisionId}`;
    const text = node.isText ? (node.text ?? '') : node.type.name === 'image' ? '[image]' : ' ';

    if (current && current.key === key && current.to === pos) {
      current.text += text;
      current.to = pos + node.nodeSize;
    } else {
      current = {
        key,
        type,
        revisionId,
        author: (mark.attrs.author as string) || 'Unknown',
        date: (mark.attrs.date as string | null) ?? null,
        text,
        from: pos,
        to: pos + node.nodeSize,
      };
      revisions.push(current);
    }
    return false;
  });

  return revisions;
}

function resolveRevisions(
  state: EditorState,
  accept: boolean,
  match: (type: 'insertion' | 'deletion', revisionId: number) => boolean
): Transaction | null {
  const groups = collectRevisionRanges(state.doc, match);
  if (groups.size === 0) return null;

  const { insertion, deletion } = state.schema.marks;
  const tr = state.tr.setMeta(SUGGESTION_SKIP_META, true);
  const toDelete: MarkedRange[] = [];

  for (const { type, ranges } of groups.values()) {
    // Accepting an insertion / rejecting a deletion keeps the text
    const keep = (type === 'insertion') === accept;
    for (const range of ranges) {
      if (keep) {
        tr.removeMark(range.from, range.to, type === 'insertion' ? insertion : deletion);
      } else {
        toDelete.push(range);
      }
    }
  }

  for (const range of toDelete.sort((a, b) => b.from - a.from)) {
    tr.delete(tr.mapping.map(range.from), tr.mapping.map(range.to));
  }
  return tr;
}

function dispatchIf(view: EditorView, tr: Transaction | null): boolean {
  if (!tr) return false;
  view.dispatch(tr);
  return true;
}

/** Accept one tracked change (identified by RevisionInfo.key) */
export function acceptRevision(view: EditorView, key: string): boolean {
  return dispatchIf(
    view,
    resolveRevisions(view.state, true, (type, id) => `${type}:${id}` === key)
  );
}

/** Reject one tracked change (identified by RevisionInfo.key) */
export function rejectRevision(view: EditorView, key: string): boolean {
  return dispatchIf(
    view,
    resolveRevisions(view.state, false, (type, id) => `${type}:${id}` === key)
  );
}

/** Accept every tracked change in the document */
export function acceptAllRevisions(view: EditorView): boolean {
  return dispatchIf(
    view,
    resolveRevisions(view.state, true, () => true)
  );
}

/** Reject every tracked change in the document */
export function rejectAllRevisions(view: EditorView): boolean {
  return dispatchIf(
    view,
    resolveRevisions(view.state, false, () => true)
  );
}

/** Tracked changes overlapping the selection (or touching a collapsed cursor) */
export function getRevisionsAtSelection(state: EditorState): RevisionInfo[] {
  const { from, to } = state.selection;
  return getRevisions(state.doc).filter((r) => r.from <= to && r.to >= from);
}

/** Next/previous tracked change relative to the selection, wrapping around */
export function findAdjacentRevision(state: EditorState, direction: 1 | -1): RevisionInfo | null {
  const revisions = getRevisions(state.doc);
  if (revisions.length === 0) return null;
  const { from, to } = state.selection;
  if (direction === 1) {
    return revisions.find((r) => r.from >= to && !(r.from === from && r.to === to)) ?? revisions[0];
  }
  const before = revisions.filter((r) => r.to <= from && !(r.from === from && r.to === to));
  return before[before.length - 1] ?? revisions[revisions.length - 1];
}

// ============================================================================
// COMMENTS
// ============================================================================

export interface CommentAnchor {
  commentId: number;
  from: number;
  to: number;
  /** Anchored text */
  text: string;
}

/** Find the anchored range of every comment in document order */
export function getCommentAnchors(doc: PMNode): CommentAnchor[] {
  const anchors = new Map<number, CommentAnchor>();
  doc.descendants((node, pos) => {
    if (!node.isInline) return true;
    for (const mark of node.marks) {
      if (mark.type.name !== 'comment') continue;
      const commentId = mark.attrs.commentId as number;
      const existing = anchors.get(commentId);
      const text = node.isText ? (node.text ?? '') : '';
      if (existing) {
        existing.to = pos + node.nodeSize;
        existing.text += text;
      } else {
        anchors.set(commentId, { commentId, from: pos, to: pos + node.nodeSize, text });
      }
    }
    return false;
  });
  return [...anchors.values()];
}

/** Next/previous comment anchor relative to the selection, wrapping around */
export function findAdjacentComment(
  anchors: CommentAnchor[],
  state: EditorState,
  direction: 1 | -1
): CommentAnchor | null {
  if (anchors.length === 0) return null;
  const { from, to } = state.selection;
  const sorted = [...anchors].sort((a, b) => a.from - b.from);
  if (direction === 1) {
    return sorted.find((a) => a.from > from || (a.from === from && a.to > to)) ?? sorted[0];
  }
  const before = sorted.filter((a) => a.from < from);
  return before[before.length - 1] ?? sorted[sorted.length - 1];
}

/** Unicode-aware word character (letters, numbers, underscore, connector punctuation) */
function isWordChar(ch: string | undefined): boolean {
  return !!ch && /[\p{L}\p{N}\p{Pc}]/u.test(ch);
}

/**
 * Expand a collapsed caret to the word under the cursor, or return the current
 * non-empty selection. Used so New Comment works after selecting/double-clicking
 * a word (including accented text) even if the live selection briefly collapses.
 *
 * When a saved `fallback` range is provided (from lastSelectionRef), prefer it
 * over the live selection — focusing the ribbon/composer often moves or corrupts
 * the live PM selection, which made the comment highlight jump elsewhere.
 */
export function resolveCommentRange(
  state: EditorState,
  fallback?: { from: number; to: number } | null
): { from: number; to: number } | null {
  const size = state.doc.content.size;

  if (fallback && fallback.from < fallback.to) {
    const a = Math.max(0, Math.min(fallback.from, size));
    const b = Math.max(0, Math.min(fallback.to, size));
    if (a < b) return { from: a, to: b };
  }

  const { from, to, empty, $from } = state.selection;
  if (!empty && from < to) return { from, to };

  if (!$from.parent.isTextblock) return null;
  const text = $from.parent.textContent;
  const offset = $from.parentOffset;
  let start = offset;
  while (start > 0 && isWordChar(text[start - 1])) start--;
  let end = offset;
  while (end < text.length && isWordChar(text[end])) end++;
  if (start >= end) return null;
  return { from: $from.start() + start, to: $from.start() + end };
}

/** Anchor a comment to a document range. Returns false for an empty range. */
export function addCommentMark(
  view: EditorView,
  commentId: number,
  range: { from: number; to: number } = view.state.selection
): boolean {
  const markType = view.state.schema.marks.comment;
  const size = view.state.doc.content.size;
  const from = Math.min(range.from, size);
  const to = Math.min(range.to, size);
  if (from >= to || !markType) return false;
  const tr = view.state.tr
    .addMark(from, to, markType.create({ commentId }))
    .setMeta(SUGGESTION_SKIP_META, true);
  view.dispatch(tr);
  return true;
}

/** Remove a comment's anchor highlight from the document */
export function removeCommentMark(view: EditorView, commentId: number): boolean {
  const markType = view.state.schema.marks.comment;
  if (!markType) return false;
  const tr = view.state.tr
    .removeMark(0, view.state.doc.content.size, markType.create({ commentId }))
    .setMeta(SUGGESTION_SKIP_META, true);
  if (!tr.docChanged) return false;
  view.dispatch(tr);
  return true;
}

/** Comment ID under the cursor, if any */
export function getCommentIdAtSelection(state: EditorState): number | null {
  const { $from, empty } = state.selection;
  // A range selection starts at the text after $from; a cursor touches both sides
  const nodes = empty ? [$from.nodeAfter, $from.nodeBefore] : [$from.nodeAfter];
  for (const node of nodes) {
    const mark = node?.marks.find((m) => m.type.name === 'comment');
    if (mark) return mark.attrs.commentId as number;
  }
  return null;
}

/** Select a document range and scroll it into view */
export function selectRange(view: EditorView, from: number, to: number): void {
  const size = view.state.doc.content.size;
  const tr = view.state.tr.setSelection(
    TextSelection.create(view.state.doc, Math.min(from, size), Math.min(to, size))
  );
  view.dispatch(tr.scrollIntoView());
}

// ============================================================================
// WORD TAGS (CONTENT CONTROLS)
// ============================================================================

export type ContentControlLock = 'sdtLocked' | 'contentLocked' | 'sdtContentLocked' | 'unlocked';

export interface ContentControlAttrs {
  tag: string;
  alias: string;
  placeholder?: string;
  /** OOXML w:lock; omit / unlocked = no lock */
  lock?: ContentControlLock | null;
}

export interface ContentControlInfo {
  /** Position of the sdt node */
  pos: number;
  tag: string | null;
  alias: string | null;
  lock: ContentControlLock | null;
  sdtType: string;
  text: string;
}

function isSdtDeletionLocked(lock: unknown): boolean {
  return lock === 'sdtLocked' || lock === 'sdtContentLocked';
}

function isSdtContentLocked(lock: unknown): boolean {
  return lock === 'contentLocked' || lock === 'sdtContentLocked';
}

/** List inline content controls in document order */
export function getContentControls(doc: PMNode): ContentControlInfo[] {
  const controls: ContentControlInfo[] = [];
  doc.descendants((node, pos) => {
    if (node.type.name === 'sdt') {
      controls.push({
        pos,
        tag: (node.attrs.tag as string | null) ?? null,
        alias: (node.attrs.alias as string | null) ?? null,
        lock: (node.attrs.lock as ContentControlLock | null) ?? null,
        sdtType: node.attrs.sdtType as string,
        text: node.textContent,
      });
      return false;
    }
    return true;
  });
  return controls;
}

/**
 * If `pos` is inside an inline SDT, return the content range (excluding the
 * wrapper node itself) so double-click can select the whole tag body.
 */
export function findSdtContentRangeAt(
  doc: PMNode,
  pos: number
): { from: number; to: number } | null {
  if (pos < 0 || pos > doc.content.size) return null;
  const $pos = doc.resolve(Math.min(pos, doc.content.size));
  for (let d = $pos.depth; d > 0; d--) {
    if ($pos.node(d).type.name === 'sdt') {
      const start = $pos.before(d);
      const node = $pos.node(d);
      const from = start + 1;
      const to = start + node.nodeSize - 1;
      if (from <= to) return { from, to };
      return { from, to: from };
    }
  }
  return null;
}

/**
 * Wrap the selection in a content control (or insert one with placeholder
 * text when nothing is selected). Always assigns a w:tag so the saved DOCX
 * matches Word's content-control Properties (Tag field).
 */
export function insertContentControl(view: EditorView, attrs: ContentControlAttrs): boolean {
  const { state } = view;
  const sdtType = state.schema.nodes.sdt;
  if (!sdtType) return false;
  const { from, to, empty, $from, $to } = state.selection;
  if (!$from.sameParent($to) || !$from.parent.inlineContent) return false;

  const occupied = getContentControls(state.doc).map((c) => c.tag);
  const { tag, alias } = normalizeContentControlTagAttrs(attrs, occupied);
  const placeholder = attrs.placeholder || alias || tag || 'Click to enter text';
  const content = empty ? state.schema.text(placeholder) : state.doc.slice(from, to).content;

  const lock = attrs.lock && attrs.lock !== 'unlocked' ? attrs.lock : null;
  const node = sdtType.create(
    {
      sdtType: 'richText',
      tag,
      alias: alias || null,
      lock,
      placeholder,
      showingPlaceholder: false,
    },
    content
  );
  if (!sdtType.validContent(node.content)) return false;

  const tr = state.tr.replaceWith(from, to, node);
  tr.setSelection(TextSelection.create(tr.doc, from + 1, from + node.nodeSize - 1));
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** Change a content control's tag / title / lock (keeps a non-empty w:tag). */
export function updateContentControl(
  view: EditorView,
  pos: number,
  attrs: { tag?: string; alias?: string; lock?: ContentControlLock | null }
): boolean {
  const node = view.state.doc.nodeAt(pos);
  if (!node || node.type.name !== 'sdt') return false;
  const nextLock =
    attrs.lock !== undefined
      ? attrs.lock && attrs.lock !== 'unlocked'
        ? attrs.lock
        : null
      : node.attrs.lock;

  const occupied = getContentControls(view.state.doc)
    .filter((c) => c.pos !== pos)
    .map((c) => c.tag);
  const { tag, alias } = normalizeContentControlTagAttrs(
    {
      tag: attrs.tag !== undefined ? attrs.tag : (node.attrs.tag as string | null),
      alias: attrs.alias !== undefined ? attrs.alias : (node.attrs.alias as string | null),
    },
    occupied
  );

  const tr = view.state.tr.setNodeMarkup(pos, undefined, {
    ...node.attrs,
    tag,
    alias: alias || null,
    lock: nextLock,
  });
  view.dispatch(tr);
  return true;
}

/** Remove a content control but keep its text. No-op when the control is deletion-locked. */
export function removeContentControl(view: EditorView, pos: number): boolean {
  const node = view.state.doc.nodeAt(pos);
  if (!node || node.type.name !== 'sdt') return false;
  if (isSdtDeletionLocked(node.attrs.lock)) return false;
  const tr = view.state.tr.replaceWith(pos, pos + node.nodeSize, node.content);
  view.dispatch(tr);
  return true;
}

/**
 * Enforce OOXML content-control locks while editing:
 * - sdtLocked / sdtContentLocked: control cannot be unwrapped/deleted
 * - contentLocked / sdtContentLocked: contents cannot be edited
 */
export function createSdtLockPlugin(): Plugin {
  return new Plugin({
    filterTransaction(tr, state) {
      if (!tr.docChanged) return true;

      const locked: { pos: number; lock: string; text: string }[] = [];
      state.doc.descendants((node, pos) => {
        const lock = node.attrs.lock as string | null;
        if (node.type.name === 'sdt' && lock && lock !== 'unlocked') {
          locked.push({ pos, lock, text: node.textContent });
        }
        return true;
      });
      if (locked.length === 0) return true;

      for (const item of locked) {
        const newPos = tr.mapping.map(item.pos, -1);
        const node = tr.doc.nodeAt(newPos);
        const stillSdt = !!node && node.type.name === 'sdt';

        if (isSdtDeletionLocked(item.lock) && !stillSdt) {
          return false;
        }
        if (isSdtContentLocked(item.lock) && stillSdt && node!.textContent !== item.text) {
          return false;
        }
      }
      return true;
    },
  });
}
