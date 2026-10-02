/**
 * CommentBalloons — Word-style comment cards in the right margin, aligned with
 * the text they comment on.
 *
 * Rendered inside the paged editor's overlay layer so balloons scroll and zoom
 * with the pages. Positions are measured from the painted anchor spans
 * (`data-comment-ids`) and stacked so cards never overlap; the active card is
 * kept level with its anchor and others move out of its way, as in Word.
 *
 * All Markup: full cards. Simple Markup: compact comment icons that expand.
 */

import {
  useCallback,
  useEffect,
  useLayoutEffect,
  useMemo,
  useRef,
  useState,
  type ReactNode,
} from 'react';
import type { Comment } from '../../types/document';
import type { CommentAnchor } from '../../prosemirror/commands/review';
import { commentToText } from '../../docx/serializer/commentSerializer';
import { Avatar, Composer, formatDate, type CommentDraft } from './ReviewSidebar';
import { RibbonIcon } from '../ribbon/RibbonIcons';
import type { MarkupView } from './types';

/** Width of the markup area reserved to the right of the page (layout px) */
export const MARKUP_AREA_WIDTH = 300;
const BALLOON_GAP = 8;
const PAGE_OFFSET = 20;

export interface CommentBalloonsProps {
  comments: Comment[];
  anchors: CommentAnchor[];
  activeCommentId: number | null;
  readOnly: boolean;
  markupView: MarkupView;
  draft: CommentDraft | null;
  onSubmitDraft: (text: string) => void;
  onCancelDraft: () => void;
  onActivate: (commentId: number | null) => void;
  onSelect: (commentId: number) => void;
  onReply: (parentId: number, text: string) => void;
  onResolve: (commentId: number, done: boolean) => void;
  onDelete: (commentId: number) => void;
}

interface Thread {
  /** Thread key: root comment id, or -1 for the draft */
  id: number;
  root?: Comment;
  replies: Comment[];
}

interface Measured {
  anchorTop: Record<number, number>;
  anchorLine: Record<number, { x: number; y: number }>;
  pageRight: number;
}

const stop = (e: { stopPropagation: () => void }) => e.stopPropagation();

/**
 * Find the painted text run containing a ProseMirror position.
 * Must prefer the tightest `.layout-run` — fragment/paragraph wrappers also
 * carry data-pm-start/end and would otherwise win (document order), pinning
 * the comment connector to the wrong block.
 */
function spanAtPos(root: Element, pos: number): HTMLElement | null {
  const runs = root.querySelectorAll<HTMLElement>(
    '.layout-page .layout-run[data-pm-start][data-pm-end]'
  );
  let best: HTMLElement | null = null;
  let bestSpan = Infinity;
  for (const span of runs) {
    const start = Number(span.dataset.pmStart);
    const end = Number(span.dataset.pmEnd);
    if (Number.isNaN(start) || Number.isNaN(end)) continue;
    // Inclusive end: caret at end of a run still maps to that run
    if (start <= pos && pos <= end) {
      const spanSize = end - start;
      if (spanSize < bestSpan) {
        best = span;
        bestSpan = spanSize;
      }
    }
  }
  return best;
}

/** Client rect covering a PM range on the painted page (for draft connectors). */
function rectForPmRange(
  root: Element,
  from: number,
  to: number
): { top: number; left: number; bottom: number; right: number } | null {
  const runs = root.querySelectorAll<HTMLElement>(
    '.layout-page .layout-run[data-pm-start][data-pm-end]'
  );
  let top = Infinity;
  let left = Infinity;
  let bottom = -Infinity;
  let right = -Infinity;
  let found = false;

  for (const span of runs) {
    const pmStart = Number(span.dataset.pmStart);
    const pmEnd = Number(span.dataset.pmEnd);
    if (!(pmEnd > from && pmStart < to)) continue;

    if (span.firstChild?.nodeType === Node.TEXT_NODE) {
      const textNode = span.firstChild as Text;
      const startChar = Math.max(0, from - pmStart);
      const endChar = Math.min(textNode.length, to - pmStart);
      if (startChar < endChar) {
        const range = span.ownerDocument!.createRange();
        range.setStart(textNode, startChar);
        range.setEnd(textNode, endChar);
        for (const r of Array.from(range.getClientRects())) {
          if (r.width <= 0 && r.height <= 0) continue;
          found = true;
          top = Math.min(top, r.top);
          left = Math.min(left, r.left);
          bottom = Math.max(bottom, r.bottom);
          right = Math.max(right, r.right);
        }
        continue;
      }
    }

    const r = span.getBoundingClientRect();
    if (r.width <= 0 && r.height <= 0) continue;
    found = true;
    top = Math.min(top, r.top);
    left = Math.min(left, r.left);
    bottom = Math.max(bottom, r.bottom);
    right = Math.max(right, r.right);
  }

  return found ? { top, left, bottom, right } : null;
}

export function CommentBalloons(props: CommentBalloonsProps) {
  const { comments, anchors, activeCommentId, readOnly, markupView, draft } = props;
  const layerRef = useRef<HTMLDivElement>(null);
  const cardRefs = useRef(new Map<number, HTMLDivElement>());
  const [measured, setMeasured] = useState<Measured | null>(null);
  const [tops, setTops] = useState<Record<number, number>>({});
  const [replyingTo, setReplyingTo] = useState<number | null>(null);
  const [menuFor, setMenuFor] = useState<number | null>(null);

  const compact = markupView === 'simple';

  const threads = useMemo<Thread[]>(() => {
    const order = new Map(anchors.map((a) => [a.commentId, a.from]));
    const roots = comments
      .filter((c) => c.parentId == null && order.has(c.id))
      .sort((a, b) => (order.get(a.id) ?? 0) - (order.get(b.id) ?? 0));
    const list: Thread[] = roots.map((root) => ({
      id: root.id,
      root,
      replies: comments.filter((c) => c.parentId === root.id),
    }));
    if (draft) list.push({ id: -1, replies: [] });
    return list;
  }, [comments, anchors, draft]);

  // --------------------------------------------------------------------------
  // Measure anchor positions from the painted pages
  // --------------------------------------------------------------------------

  const measure = useCallback(() => {
    const layer = layerRef.current;
    const editor = layer?.closest('.paged-editor');
    if (!layer || !editor) return;
    const layerRect = layer.getBoundingClientRect();
    const scale = layer.offsetWidth ? layerRect.width / layer.offsetWidth : 1;
    const page = editor.querySelector('.layout-page');
    if (!page) return;
    // Balloons go after any plugin panel that sits beside the page
    const pluginSpace =
      parseFloat(getComputedStyle(editor).getPropertyValue('--plugin-panel-space')) || 0;
    const pageRight = (page.getBoundingClientRect().right - layerRect.left) / scale + pluginSpace;

    const anchorTop: Record<number, number> = {};
    const anchorLine: Record<number, { x: number; y: number }> = {};
    const recordRect = (
      id: number,
      r: { top: number; left: number; bottom: number; right: number } | null
    ) => {
      if (!r) return;
      anchorTop[id] = (r.top - layerRect.top) / scale;
      // Connector meets the text at the right edge of the last line of the range
      anchorLine[id] = {
        x: (r.right - layerRect.left) / scale,
        y: (r.bottom - layerRect.top) / scale,
      };
    };
    const recordEl = (id: number, el: Element | null) => {
      if (!el) return;
      const box = el.getBoundingClientRect();
      recordRect(id, {
        top: box.top,
        left: box.left,
        bottom: box.bottom,
        right: box.right,
      });
    };
    for (const t of threads) {
      if (t.id === -1) {
        if (draft) {
          // Prefer a precise DOM range over a single span — avoids pinning the
          // draft connector to a fragment wrapper or the wrong paragraph.
          const rangeRect = rectForPmRange(editor, draft.from, draft.to);
          if (rangeRect) recordRect(-1, rangeRect);
          else recordEl(-1, spanAtPos(editor, draft.from));
        }
      } else {
        recordEl(t.id, editor.querySelector(`.layout-page [data-comment-ids~="${t.id}"]`));
      }
    }
    setMeasured((prev) => {
      // Keep the last known position for anchors that are momentarily not painted
      // (mid-repaint), so balloons — and any half-typed reply — stay mounted.
      const live = new Set(threads.map((t) => t.id));
      const keep = <T,>(prevMap: Record<number, T> | undefined, nextMap: Record<number, T>) => {
        const merged: Record<number, T> = {};
        for (const id of live) {
          const value = nextMap[id] ?? prevMap?.[id];
          if (value !== undefined) merged[id] = value;
        }
        return merged;
      };
      const next = {
        anchorTop: keep(prev?.anchorTop, anchorTop),
        anchorLine: keep(prev?.anchorLine, anchorLine),
        pageRight,
      };
      return prev && JSON.stringify(prev) === JSON.stringify(next) ? prev : next;
    });
  }, [threads, draft]);

  // Re-measure when pages repaint, zoom changes, or the window resizes
  useEffect(() => {
    const layer = layerRef.current;
    const editor = layer?.closest('.paged-editor');
    if (!editor) return;
    let frame = 0;
    const schedule = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(measure);
    };
    const observer = new MutationObserver((mutations) => {
      if (mutations.some((m) => !layer?.contains(m.target))) schedule();
    });
    observer.observe(editor, {
      childList: true,
      subtree: true,
      attributes: true,
      attributeFilter: ['style'],
    });
    window.addEventListener('resize', schedule);
    // Side panes opening/closing resize the editor without a window resize
    const resizeObserver = new ResizeObserver(schedule);
    resizeObserver.observe(editor);
    schedule();
    return () => {
      cancelAnimationFrame(frame);
      observer.disconnect();
      resizeObserver.disconnect();
      window.removeEventListener('resize', schedule);
    };
  }, [measure]);

  // --------------------------------------------------------------------------
  // Stack balloons without overlap, keeping the active one level with its anchor
  // --------------------------------------------------------------------------

  const focusId = draft ? -1 : activeCommentId;

  // Runs after every render on purpose: card heights are read from the DOM, and
  // setTops bails out when nothing moved, so this settles in one extra pass.
  // eslint-disable-next-line react-hooks/exhaustive-deps
  useLayoutEffect(() => {
    if (!measured) return;
    const placed = threads
      .filter((t) => measured.anchorTop[t.id] != null)
      .map((t) => ({
        id: t.id,
        want: measured.anchorTop[t.id],
        height: cardRefs.current.get(t.id)?.offsetHeight ?? 40,
      }))
      .sort((a, b) => a.want - b.want);

    const next: Record<number, number> = {};
    const pivot = Math.max(
      0,
      placed.findIndex((p) => p.id === focusId)
    );
    if (placed.length > 0) {
      next[placed[pivot].id] = placed[pivot].want;
      // Downwards from the pivot
      for (let i = pivot + 1; i < placed.length; i++) {
        const prev = placed[i - 1];
        next[placed[i].id] = Math.max(placed[i].want, next[prev.id] + prev.height + BALLOON_GAP);
      }
      // Upwards from the pivot
      for (let i = pivot - 1; i >= 0; i--) {
        const below = placed[i + 1];
        next[placed[i].id] = Math.min(
          placed[i].want,
          next[below.id] - placed[i].height - BALLOON_GAP
        );
      }
    }
    setTops((prev) => (JSON.stringify(prev) === JSON.stringify(next) ? prev : next));
  });

  if (markupView === 'none' || markupView === 'original') return null;

  const left = (measured?.pageRight ?? 0) + PAGE_OFFSET;
  const activeAnchor = focusId != null ? measured?.anchorLine[focusId] : undefined;
  const activeTop = focusId != null ? tops[focusId] : undefined;

  const card = (thread: Thread): ReactNode => {
    const isActive = thread.id === focusId;
    const top = tops[thread.id];
    if (top == null || !measured) return null;

    const style = { top, left, width: MARKUP_AREA_WIDTH - PAGE_OFFSET - 12 };

    if (thread.id === -1) {
      return (
        <div
          key="draft"
          ref={(el) => {
            if (el) cardRefs.current.set(-1, el);
            else cardRefs.current.delete(-1);
          }}
          className="ep-balloon is-active"
          style={style}
          data-testid="comment-draft"
        >
          <Composer
            placeholder="Start a conversation"
            submitLabel="Comment"
            autoFocus
            testId="comment-composer"
            onSubmit={props.onSubmitDraft}
            onCancel={props.onCancelDraft}
          />
        </div>
      );
    }

    const root = thread.root!;
    if (compact && !isActive) {
      return (
        <button
          key={thread.id}
          ref={(el) => {
            if (el) cardRefs.current.set(thread.id, el as unknown as HTMLDivElement);
            else cardRefs.current.delete(thread.id);
          }}
          type="button"
          className="ep-balloon-icon"
          style={{ top, left }}
          title={`${root.author}: ${commentToText(root)}`}
          aria-label={`Comment by ${root.author}`}
          data-testid="comment-indicator"
          onClick={() => props.onSelect(thread.id)}
        >
          <RibbonIcon name="showComments" size={18} />
        </button>
      );
    }

    return (
      <div
        key={thread.id}
        ref={(el) => {
          if (el) cardRefs.current.set(thread.id, el);
          else cardRefs.current.delete(thread.id);
        }}
        className={`ep-balloon${isActive ? ' is-active' : ''}${root.done ? ' is-resolved' : ''}`}
        style={style}
        data-testid="comment-thread"
        data-comment-id={root.id}
        onClick={() => props.onSelect(thread.id)}
      >
        {root.done && <div className="ep-balloon__resolved">Resolved</div>}
        {[root, ...thread.replies].map((c, i) => (
          <div key={c.id} className={`ep-balloon__entry${i > 0 ? ' is-reply' : ''}`}>
            <div className="ep-balloon__head">
              <Avatar name={c.author} />
              <div className="ep-balloon__meta">
                <span className="ep-balloon__author">{c.author}</span>
                <span className="ep-balloon__date">{formatDate(c.date)}</span>
              </div>
              {i === 0 && !readOnly && (
                <div className="ep-balloon__actions" onClick={(e) => e.stopPropagation()}>
                  <button
                    type="button"
                    className="ep-balloon__action"
                    title={root.done ? 'Reopen thread' : 'Resolve thread'}
                    aria-label={root.done ? 'Reopen thread' : 'Resolve thread'}
                    data-testid="comment-resolve"
                    onClick={() => props.onResolve(root.id, !root.done)}
                  >
                    <RibbonIcon name="check" size={16} />
                  </button>
                  <div className="ep-balloon__more">
                    <button
                      type="button"
                      className="ep-balloon__action"
                      title="More thread actions"
                      aria-label="More thread actions"
                      aria-haspopup="menu"
                      aria-expanded={menuFor === root.id}
                      onClick={() => setMenuFor(menuFor === root.id ? null : root.id)}
                    >
                      <svg
                        width="16"
                        height="16"
                        viewBox="0 0 24 24"
                        fill="currentColor"
                        aria-hidden="true"
                      >
                        <circle cx="5" cy="12" r="1.6" />
                        <circle cx="12" cy="12" r="1.6" />
                        <circle cx="19" cy="12" r="1.6" />
                      </svg>
                    </button>
                    {menuFor === root.id && (
                      <div className="ep-ribbon-menu ep-ribbon-menu--right" role="menu">
                        <button
                          type="button"
                          role="menuitem"
                          className="ep-ribbon-menu__item"
                          data-testid="comment-delete"
                          onClick={() => {
                            setMenuFor(null);
                            props.onDelete(root.id);
                          }}
                        >
                          <span className="ep-ribbon-menu__icon">
                            <RibbonIcon name="deleteComment" size={18} />
                          </span>
                          <span className="ep-ribbon-menu__label">Delete thread</span>
                        </button>
                        <button
                          type="button"
                          role="menuitem"
                          className="ep-ribbon-menu__item"
                          onClick={() => {
                            setMenuFor(null);
                            props.onResolve(root.id, !root.done);
                          }}
                        >
                          <span className="ep-ribbon-menu__icon">
                            <RibbonIcon name="check" size={18} />
                          </span>
                          <span className="ep-ribbon-menu__label">
                            {root.done ? 'Reopen thread' : 'Resolve thread'}
                          </span>
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              )}
              {i > 0 && !readOnly && (
                <button
                  type="button"
                  className="ep-balloon__action"
                  title="Delete reply"
                  aria-label="Delete reply"
                  onClick={(e) => {
                    e.stopPropagation();
                    props.onDelete(c.id);
                  }}
                >
                  <RibbonIcon name="close" size={14} />
                </button>
              )}
            </div>
            <p className="ep-balloon__text">{commentToText(c)}</p>
          </div>
        ))}

        {!readOnly && !root.done && (
          <div className="ep-balloon__reply" onClick={(e) => e.stopPropagation()}>
            {replyingTo === root.id ? (
              <Composer
                placeholder="Reply"
                submitLabel="Reply"
                autoFocus
                testId="reply-composer"
                onSubmit={(text) => {
                  props.onReply(root.id, text);
                  setReplyingTo(null);
                }}
                onCancel={() => setReplyingTo(null)}
              />
            ) : (
              <button
                type="button"
                className="ep-balloon__reply-btn"
                data-testid="comment-reply"
                onClick={() => {
                  props.onActivate(root.id);
                  setReplyingTo(root.id);
                }}
              >
                Reply
              </button>
            )}
          </div>
        )}
      </div>
    );
  };

  return (
    <div
      ref={layerRef}
      className="ep-balloons"
      aria-label="Comments"
      role="complementary"
      // The paged editor redirects focus, clicks and keys to its hidden
      // ProseMirror instance; keep balloon text fields out of that.
      onFocus={stop}
      onMouseDown={stop}
      onMouseUp={stop}
      onClick={stop}
      onKeyDown={stop}
    >
      {/* Dashed connector from the active comment's text to its balloon (Word style) */}
      {activeAnchor &&
        activeTop != null &&
        measured &&
        !(compact && focusId !== -1 && !activeCommentId) && (
          <svg className="ep-balloon-connector" aria-hidden="true">
            <polyline
              points={`${activeAnchor.x},${activeAnchor.y} ${measured.pageRight + 6},${activeAnchor.y} ${left},${activeTop + 16}`}
            />
          </svg>
        )}
      {threads.map(card)}
    </div>
  );
}
