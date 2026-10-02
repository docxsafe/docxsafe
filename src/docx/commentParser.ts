/**
 * Comment Parser - Parse comments.xml
 *
 * Parses OOXML comments (w:comment) from comments.xml file.
 *
 * OOXML Reference:
 * - Comments: w:comments
 * - Comment: w:comment (w:id, w:author, w:date, w:initials)
 * - Comment content: child w:p elements
 */

import type { Comment, Paragraph, Theme, RelationshipMap, MediaFile } from '../types/document';
import type { StyleMap } from './styleParser';
import { parseXml, findChild, getChildElements, getAttribute } from './xmlParser';
import { parseParagraph } from './paragraphParser';

/**
 * Parse comments.xml into an array of Comment objects
 */
export function parseComments(
  commentsXml: string | null,
  styles: StyleMap | null,
  theme: Theme | null,
  rels: RelationshipMap,
  media: Map<string, MediaFile>,
  commentsExtendedXml?: string | null
): Comment[] {
  if (!commentsXml) return [];
  const extended = parseCommentsExtended(commentsExtendedXml);

  const root = parseXml(commentsXml);
  if (!root) return [];

  const commentsEl = findChild(root, 'w', 'comments') ?? root;
  const children = getChildElements(commentsEl);
  const comments: Comment[] = [];

  for (const child of children) {
    const localName = child.name?.replace(/^.*:/, '') ?? '';
    if (localName !== 'comment') continue;

    const id = parseInt(getAttribute(child, 'w', 'id') ?? '0', 10);
    const author = getAttribute(child, 'w', 'author') ?? 'Unknown';
    const initials = getAttribute(child, 'w', 'initials') ?? undefined;
    const date = getAttribute(child, 'w', 'date') ?? undefined;

    // Parse comment content (paragraphs)
    const paragraphs: Paragraph[] = [];
    let lastParaId: string | undefined;
    for (const contentChild of getChildElements(child)) {
      const contentName = contentChild.name?.replace(/^.*:/, '') ?? '';
      if (contentName === 'p') {
        lastParaId = getAttribute(contentChild, 'w14', 'paraId') ?? lastParaId;
        const paragraph = parseParagraph(contentChild, styles, theme, null, rels, media);
        paragraphs.push(paragraph);
      }
    }

    comments.push({
      id,
      author,
      initials,
      date,
      content: paragraphs,
      ...(lastParaId ? { paraId: lastParaId } : {}),
    } as Comment);
  }

  // Resolve reply threads and resolved state from commentsExtended.xml
  if (extended.size > 0) {
    const idByParaId = new Map<string, number>();
    for (const c of comments as (Comment & { paraId?: string })[]) {
      if (c.paraId) idByParaId.set(c.paraId, c.id);
    }
    for (const c of comments as (Comment & { paraId?: string })[]) {
      const ext = c.paraId ? extended.get(c.paraId) : undefined;
      if (ext) {
        if (ext.done) c.done = true;
        if (ext.parentParaId && idByParaId.has(ext.parentParaId)) {
          c.parentId = idByParaId.get(ext.parentParaId);
        }
      }
      delete c.paraId;
    }
  }

  return comments;
}

/**
 * Parse commentsExtended.xml into paraId → { parentParaId, done }
 */
function parseCommentsExtended(
  xml: string | null | undefined
): Map<string, { parentParaId?: string; done: boolean }> {
  const result = new Map<string, { parentParaId?: string; done: boolean }>();
  if (!xml) return result;
  const root = parseXml(xml);
  if (!root) return result;
  const container = findChild(root, 'w15', 'commentsEx') ?? root;
  for (const el of getChildElements(container)) {
    if (el.name?.replace(/^.*:/, '') !== 'commentEx') continue;
    const paraId = getAttribute(el, 'w15', 'paraId');
    if (!paraId) continue;
    result.set(paraId, {
      parentParaId: getAttribute(el, 'w15', 'paraIdParent') ?? undefined,
      done: getAttribute(el, 'w15', 'done') === '1',
    });
  }
  return result;
}
