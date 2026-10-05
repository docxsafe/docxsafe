import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  allocateCommentId,
  commentToShared,
  observeSharedComments,
  readSharedComments,
  removeSharedComment,
  seedSharedCommentsIfEmpty,
  sharedToComment,
  upsertSharedComment,
  type SharedCommentRecord,
} from './commentsSync';
import type { Comment } from '../types/document';
import { commentParagraphsFromText, serializeComments } from '../docx/serializer/commentSerializer';

function makeComment(partial: Partial<Comment> & Pick<Comment, 'id' | 'author'>): Comment {
  return {
    content: commentParagraphsFromText('hello'),
    initials: 'AB',
    date: '2026-10-03T18:00:00Z',
    ...partial,
  };
}

describe('commentsSync (OOXML)', () => {
  test('upsert stores full paragraph content, not just plain text', () => {
    const doc = new Y.Doc();
    const rich = makeComment({
      id: 0,
      author: 'Alice',
      content: commentParagraphsFromText('line one\nline two'),
    });
    upsertSharedComment(doc, rich);
    const list = readSharedComments(doc);
    expect(list).toHaveLength(1);
    expect(list[0].content).toHaveLength(2);
    expect(list[0].content[0].type).toBe('paragraph');
    expect(list[0].initials).toBe('AB');
  });

  test('round-trip still serializes to valid comments.xml', () => {
    const doc = new Y.Doc();
    upsertSharedComment(
      doc,
      makeComment({
        id: 3,
        author: 'Bob',
        content: commentParagraphsFromText('Word balloon'),
      })
    );
    const xml = serializeComments(readSharedComments(doc));
    expect(xml).toContain('<w:comments');
    expect(xml).toContain('w:id="3"');
    expect(xml).toContain('w:author="Bob"');
    expect(xml).toContain('Word balloon');
    expect(xml).toContain('<w:annotationRef/>');
  });

  test('legacy text-only records still load as OOXML paragraphs', () => {
    const legacy: SharedCommentRecord = {
      id: 1,
      author: 'Legacy',
      text: 'old body',
    };
    const comment = sharedToComment(legacy);
    expect(comment.content[0].type).toBe('paragraph');
    expect(serializeComments([comment])).toContain('old body');
  });

  test('commentToShared keeps content + text mirror', () => {
    const shared = commentToShared(
      makeComment({ id: 2, author: 'C', content: commentParagraphsFromText('mirrored') })
    );
    expect(shared.content?.length).toBe(1);
    expect(shared.text).toBe('mirrored');
  });

  test('observe fires for remote-style updates', () => {
    const doc = new Y.Doc();
    const seen: number[] = [];
    const stop = observeSharedComments(doc, (comments) => {
      seen.push(comments.length);
    });
    expect(seen).toEqual([0]);
    upsertSharedComment(doc, makeComment({ id: 1, author: 'Bob' }));
    expect(seen.at(-1)).toBe(1);
    stop();
  });

  test('seedSharedCommentsIfEmpty preserves DOCX paragraph bodies', () => {
    const doc = new Y.Doc();
    const local = [
      makeComment({
        id: 2,
        author: 'Carol',
        content: commentParagraphsFromText('from docx'),
      }),
    ];
    expect(seedSharedCommentsIfEmpty(doc, local)).toBe(true);
    expect(seedSharedCommentsIfEmpty(doc, local)).toBe(false);
    expect(readSharedComments(doc)[0].content[0].type).toBe('paragraph');
  });

  test('allocateCommentId avoids collisions', () => {
    const doc = new Y.Doc();
    upsertSharedComment(doc, makeComment({ id: 5, author: 'A' }));
    expect(allocateCommentId(doc, [makeComment({ id: 3, author: 'B' })])).toBe(6);
  });

  test('removeSharedComment drops replies', () => {
    const doc = new Y.Doc();
    upsertSharedComment(doc, makeComment({ id: 1, author: 'A' }));
    upsertSharedComment(doc, makeComment({ id: 2, author: 'B', parentId: 1 }));
    removeSharedComment(doc, 1);
    expect(readSharedComments(doc)).toHaveLength(0);
  });
});
