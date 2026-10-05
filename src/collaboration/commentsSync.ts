/**
 * Sync review comments over the shared Y.Doc using OOXML Comment shape.
 *
 * Comment *ranges* (w:commentRangeStart / End / w:commentReference) live in the
 * ProseMirror fragment and sync via y-prosemirror into document.xml.
 *
 * Comment *bodies* (w:comment in comments.xml + w15:commentEx) are stored here
 * as the same `Comment` model the parser/serializer use, so File → Save writes
 * valid Word OOXML that opens with balloons intact.
 *
 * OOXML Reference:
 * - ECMA-376 §17.13.4 (comments)
 * - MS-DOCX commentsExtended (w15:commentEx)
 */

import type * as Y from 'yjs';
import type { Comment, Paragraph } from '../types/document';
import { commentParagraphsFromText, commentToText } from '../docx/serializer/commentSerializer';

export const COMMENTS_MAP_FIELD = 'comments';

/**
 * JSON-serializable OOXML comment record stored in Yjs.
 * Mirrors `Comment` / w:comment (+ reply/done from commentsExtended).
 */
export interface SharedCommentRecord {
  /** w:id */
  id: number;
  /** w:author */
  author: string;
  /** w:initials */
  initials?: string;
  /** w:date (ISO-8601) */
  date?: string;
  /**
   * w:comment body as OOXML paragraphs (runs/text).
   * Preferred over legacy `text`.
   */
  content?: Paragraph[];
  /**
   * @deprecated Legacy plain-text body from early collab sync.
   * Still read for rooms that seeded before OOXML content landed.
   */
  text?: string;
  /** w15:paraIdParent → parent comment id */
  parentId?: number;
  /** w15:done */
  done?: boolean;
}

export function getCommentsMap(doc: Y.Doc): Y.Map<SharedCommentRecord> {
  return doc.getMap(COMMENTS_MAP_FIELD) as Y.Map<SharedCommentRecord>;
}

/** Deep-clone via JSON so Yjs stores plain data (no shared object refs). */
function cloneJson<T>(value: T): T {
  return JSON.parse(JSON.stringify(value)) as T;
}

export function commentToShared(comment: Comment): SharedCommentRecord {
  const content =
    comment.content?.length > 0 ? cloneJson(comment.content) : commentParagraphsFromText('');
  return {
    id: comment.id,
    author: comment.author,
    initials: comment.initials,
    date: comment.date,
    content,
    // Keep plain text alongside for older peers / debugging
    text: commentToText({ ...comment, content }),
    parentId: comment.parentId,
    done: comment.done,
  };
}

export function sharedToComment(record: SharedCommentRecord): Comment {
  const content =
    Array.isArray(record.content) && record.content.length > 0
      ? cloneJson(record.content)
      : commentParagraphsFromText(record.text ?? '');
  return {
    id: record.id,
    author: record.author || 'Unknown',
    initials: record.initials,
    date: record.date,
    content,
    parentId: record.parentId,
    done: record.done,
  };
}

export function readSharedComments(doc: Y.Doc): Comment[] {
  const out: Comment[] = [];
  getCommentsMap(doc).forEach((value) => {
    if (value && typeof value.id === 'number') {
      out.push(sharedToComment(value));
    }
  });
  return out.sort((a, b) => a.id - b.id);
}

/** Next free comment id across the shared map and local list */
export function allocateCommentId(doc: Y.Doc | null, localComments: Comment[]): number {
  let max = -1;
  if (doc) {
    getCommentsMap(doc).forEach((value) => {
      if (value?.id != null && value.id > max) max = value.id;
    });
  }
  for (const c of localComments) {
    if (c.id > max) max = c.id;
  }
  return max + 1;
}

export function upsertSharedComment(doc: Y.Doc, comment: Comment): void {
  const map = getCommentsMap(doc);
  doc.transact(() => {
    map.set(String(comment.id), commentToShared(comment));
  });
}

export function removeSharedComment(doc: Y.Doc, commentId: number): void {
  const map = getCommentsMap(doc);
  doc.transact(() => {
    map.delete(String(commentId));
    const replyKeys: string[] = [];
    map.forEach((value, key) => {
      if (value?.parentId === commentId) replyKeys.push(key);
    });
    for (const key of replyKeys) map.delete(key);
  });
}

export function clearSharedComments(doc: Y.Doc): void {
  const map = getCommentsMap(doc);
  doc.transact(() => {
    for (const key of [...map.keys()]) map.delete(key);
  });
}

/**
 * If the shared map is empty, copy local comments in (first peer / seed).
 * Preserves full OOXML paragraph content from the loaded DOCX.
 * Returns true when seeding happened.
 */
export function seedSharedCommentsIfEmpty(doc: Y.Doc, comments: Comment[]): boolean {
  const map = getCommentsMap(doc);
  if (map.size > 0 || comments.length === 0) return false;
  doc.transact(() => {
    for (const c of comments) {
      map.set(String(c.id), commentToShared(c));
    }
  });
  return true;
}

/** Observe shared comments; fires immediately with the current snapshot. */
export function observeSharedComments(
  doc: Y.Doc,
  onChange: (comments: Comment[]) => void
): () => void {
  const map = getCommentsMap(doc);
  const emit = () => onChange(readSharedComments(doc));
  map.observe(emit);
  emit();
  return () => map.unobserve(emit);
}
