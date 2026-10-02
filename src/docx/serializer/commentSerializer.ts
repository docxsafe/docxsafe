/**
 * Comment Serializer - Serialize comments to comments.xml / commentsExtended.xml
 *
 * OOXML Reference:
 * - w:comments / w:comment (ECMA-376 §17.13.4.2)
 * - Replies and resolved state live in commentsExtended.xml (MS-DOCX w15:commentEx),
 *   linked through the w14:paraId of each comment's last paragraph.
 */

import type { Comment, Paragraph } from '../../types/document';
import { serializeParagraph } from './paragraphSerializer';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006';

export const COMMENTS_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml';
export const COMMENTS_EXTENDED_CONTENT_TYPE =
  'application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml';
export const COMMENTS_REL_TYPE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments';
export const COMMENTS_EXTENDED_REL_TYPE =
  'http://schemas.microsoft.com/office/2011/relationships/commentsExtended';

function escapeXml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Deterministic w14:paraId for a comment (must be < 0x80000000)
 */
export function commentParaId(commentId: number): string {
  return (0x10000000 + commentId).toString(16).toUpperCase().padStart(8, '0');
}

/**
 * Build a plain-text comment body as paragraphs (one per line)
 */
export function commentParagraphsFromText(text: string): Paragraph[] {
  return text.split('\n').map((line) => ({
    type: 'paragraph',
    content: line ? [{ type: 'run', content: [{ type: 'text', text: line }] }] : [],
  }));
}

/**
 * Plain text of a comment's paragraphs (for display in the sidebar)
 */
export function commentToText(comment: Comment): string {
  return comment.content
    .map((p) =>
      p.content
        .map((item) =>
          item.type === 'run'
            ? item.content.map((c) => (c.type === 'text' ? c.text : '')).join('')
            : item.type === 'hyperlink'
              ? item.children
                  .map((r) =>
                    r.type === 'run'
                      ? r.content.map((c) => (c.type === 'text' ? c.text : '')).join('')
                      : ''
                  )
                  .join('')
              : ''
        )
        .join('')
    )
    .join('\n');
}

function serializeComment(comment: Comment): string {
  const attrs = [`w:id="${comment.id}"`, `w:author="${escapeXml(comment.author || 'Unknown')}"`];
  if (comment.date) attrs.push(`w:date="${escapeXml(comment.date)}"`);
  if (comment.initials) attrs.push(`w:initials="${escapeXml(comment.initials)}"`);

  const source = comment.content.length > 0 ? comment.content : commentParagraphsFromText('');
  const last = source.length - 1;
  const body = source
    .map((p, i) => {
      // Only the last paragraph carries the paraId commentsExtended.xml links to
      const paragraph: Paragraph = {
        ...p,
        paraId: i === last ? commentParaId(comment.id) : undefined,
        textId: i === last ? '77777777' : undefined,
      };
      let xml = serializeParagraph(paragraph);
      if (i === 0) {
        // Word expects the annotation reference mark in the first paragraph
        xml = xml.replace(
          /^(<w:p(?:\s[^>]*)?>(?:<w:pPr>.*?<\/w:pPr>)?)/,
          '$1<w:r><w:annotationRef/></w:r>'
        );
      }
      return xml;
    })
    .join('');

  return `<w:comment ${attrs.join(' ')}>${body}</w:comment>`;
}

/**
 * Serialize comments to word/comments.xml
 */
export function serializeComments(comments: Comment[]): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w14">` +
    comments.map(serializeComment).join('') +
    `</w:comments>`
  );
}

/**
 * Serialize reply/resolved metadata to word/commentsExtended.xml
 */
export function serializeCommentsExtended(comments: Comment[]): string {
  const entries = comments.map((c) => {
    const attrs = [`w15:paraId="${commentParaId(c.id)}"`];
    if (c.parentId != null) attrs.push(`w15:paraIdParent="${commentParaId(c.parentId)}"`);
    attrs.push(`w15:done="${c.done ? 1 : 0}"`);
    return `<w15:commentEx ${attrs.join(' ')}/>`;
  });
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<w15:commentsEx xmlns:w15="${W15_NS}" xmlns:mc="${MC_NS}" mc:Ignorable="w15">` +
    entries.join('') +
    `</w15:commentsEx>`
  );
}
