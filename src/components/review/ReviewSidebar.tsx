/**
 * ReviewSidebar — comments, suggestions (tracked changes), and Word tags
 *
 * Comments: threaded, with reply / resolve / reopen / delete.
 * Suggestions: accept or reject individually or all at once.
 * Tags: insert, rename, and remove content controls.
 */

import { useEffect, useMemo, useRef, useState, type FormEvent, type ReactNode } from 'react';
import type { Comment } from '../../types/document';
import type {
  CommentAnchor,
  ContentControlAttrs,
  ContentControlInfo,
  RevisionInfo,
} from '../../prosemirror/commands/review';
import { sdtLockFlags, sdtLockFromFlags, type SdtLock } from '../../types/content';
import { commentToText } from '../../docx/serializer/commentSerializer';
import { cn } from '../../lib/utils';

export type ReviewTab = 'comments' | 'suggestions' | 'tags';

export interface CommentDraft {
  from: number;
  to: number;
  quote: string;
}

export interface ReviewSidebarProps {
  tab: ReviewTab;
  onTabChange: (tab: ReviewTab) => void;
  onClose: () => void;
  /** Viewing mode: browse only */
  readOnly: boolean;

  comments: Comment[];
  anchors: CommentAnchor[];
  activeCommentId: number | null;
  draft: CommentDraft | null;
  onSubmitDraft: (text: string) => void;
  onCancelDraft: () => void;
  onReply: (parentId: number, text: string) => void;
  onResolve: (commentId: number, done: boolean) => void;
  onDeleteComment: (commentId: number) => void;
  onSelectComment: (commentId: number) => void;

  revisions: RevisionInfo[];
  onAcceptRevision: (key: string) => void;
  onRejectRevision: (key: string) => void;
  onAcceptAll: () => void;
  onRejectAll: () => void;
  onSelectRevision: (revision: RevisionInfo) => void;

  contentControls: ContentControlInfo[];
  tagFormOpen: boolean;
  onTagFormOpenChange: (open: boolean) => void;
  onInsertTag: (attrs: ContentControlAttrs) => void;
  onUpdateTag: (pos: number, attrs: ContentControlAttrs) => void;
  onRemoveTag: (pos: number) => void;
  onSelectTag: (control: ContentControlInfo) => void;
}

// ============================================================================
// HELPERS
// ============================================================================

export function formatDate(date?: string | null): string {
  if (!date) return '';
  const d = new Date(date);
  if (Number.isNaN(d.getTime())) return '';
  return d.toLocaleString(undefined, {
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

function initialsOf(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  return (parts[0][0] + (parts.length > 1 ? parts[parts.length - 1][0] : '')).toUpperCase();
}

const AVATAR_COLORS = ['#7c3aed', '#0891b2', '#c2410c', '#15803d', '#be185d', '#1d4ed8'];

export function Avatar({ name }: { name: string }) {
  let hash = 0;
  for (const ch of name) hash = (hash * 31 + ch.charCodeAt(0)) >>> 0;
  return (
    <span
      className="inline-flex items-center justify-center w-7 h-7 rounded-full text-white text-xs font-semibold flex-shrink-0"
      style={{ backgroundColor: AVATAR_COLORS[hash % AVATAR_COLORS.length] }}
      aria-hidden="true"
    >
      {initialsOf(name)}
    </span>
  );
}

function TextAction({
  onClick,
  children,
  tone = 'default',
  testId,
}: {
  onClick: () => void;
  children: ReactNode;
  tone?: 'default' | 'danger' | 'primary';
  testId?: string;
}) {
  return (
    <button
      type="button"
      data-testid={testId}
      onClick={(e) => {
        e.stopPropagation();
        onClick();
      }}
      className={cn(
        'text-xs font-medium px-1.5 py-0.5 rounded hover:bg-slate-100',
        tone === 'danger' && 'text-red-600 hover:bg-red-50',
        tone === 'primary' && 'text-blue-700 hover:bg-blue-50',
        tone === 'default' && 'text-slate-600'
      )}
    >
      {children}
    </button>
  );
}

export function Composer({
  placeholder,
  submitLabel,
  onSubmit,
  onCancel,
  autoFocus,
  testId,
}: {
  placeholder: string;
  submitLabel: string;
  onSubmit: (text: string) => void;
  onCancel?: () => void;
  autoFocus?: boolean;
  testId: string;
}) {
  const [text, setText] = useState('');
  const ref = useRef<HTMLTextAreaElement>(null);

  useEffect(() => {
    if (autoFocus) ref.current?.focus();
  }, [autoFocus]);

  const submit = (e?: FormEvent) => {
    e?.preventDefault();
    if (!text.trim()) return;
    onSubmit(text.trim());
    setText('');
  };

  return (
    <form onSubmit={submit} className="flex flex-col gap-2" data-testid={testId}>
      <textarea
        ref={ref}
        value={text}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === 'Enter' && (e.metaKey || e.ctrlKey)) submit();
          if (e.key === 'Escape') onCancel?.();
        }}
        placeholder={placeholder}
        rows={2}
        className="w-full resize-y rounded-md border border-slate-300 px-2 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/40 focus:border-blue-500"
      />
      <div className="flex justify-end gap-1.5">
        {onCancel && (
          <button
            type="button"
            onClick={onCancel}
            className="h-7 px-3 rounded-md text-sm text-slate-600 hover:bg-slate-100"
          >
            Cancel
          </button>
        )}
        <button
          type="submit"
          disabled={!text.trim()}
          className="h-7 px-3 rounded-md text-sm font-medium bg-blue-600 text-white hover:bg-blue-700 disabled:opacity-40"
        >
          {submitLabel}
        </button>
      </div>
    </form>
  );
}

function EmptyState({ children }: { children: ReactNode }) {
  return <p className="text-sm text-slate-500 text-center px-4 py-8">{children}</p>;
}

// ============================================================================
// COMMENTS TAB
// ============================================================================

function CommentThread({
  root,
  replies,
  anchor,
  active,
  readOnly,
  onReply,
  onResolve,
  onDelete,
  onSelect,
}: {
  root: Comment;
  replies: Comment[];
  anchor?: CommentAnchor;
  active: boolean;
  readOnly: boolean;
  onReply: (text: string) => void;
  onResolve: (done: boolean) => void;
  onDelete: (id: number) => void;
  onSelect: () => void;
}) {
  const [replying, setReplying] = useState(false);
  const cardRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (active) cardRef.current?.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
  }, [active]);

  return (
    <div
      ref={cardRef}
      onClick={onSelect}
      data-testid="pane-comment-thread"
      data-comment-id={root.id}
      className={cn(
        'rounded-lg border bg-white p-3 cursor-pointer transition-shadow',
        active ? 'border-amber-400 shadow-md' : 'border-slate-200 hover:shadow-sm',
        root.done && 'opacity-70'
      )}
    >
      {anchor?.text && (
        <p className="text-xs text-slate-500 border-l-2 border-amber-400 pl-2 mb-2 line-clamp-2">
          {anchor.text}
        </p>
      )}

      {[root, ...replies].map((c, i) => (
        <div
          key={c.id}
          className={cn('flex gap-2', i > 0 && 'mt-3 pt-3 border-t border-slate-100')}
        >
          <Avatar name={c.author} />
          <div className="min-w-0 flex-1">
            <div className="flex items-baseline gap-2">
              <span className="text-sm font-semibold text-slate-900 truncate">{c.author}</span>
              <span className="text-xs text-slate-500 flex-shrink-0">{formatDate(c.date)}</span>
            </div>
            <p className="text-sm text-slate-800 whitespace-pre-wrap break-words mt-0.5">
              {commentToText(c)}
            </p>
            {!readOnly && i > 0 && (
              <div className="mt-1 -ml-1.5">
                <TextAction tone="danger" onClick={() => onDelete(c.id)}>
                  Delete
                </TextAction>
              </div>
            )}
          </div>
        </div>
      ))}

      {!readOnly && (
        <div className="mt-2 -ml-1.5 flex flex-wrap gap-1">
          {!root.done && !replying && (
            <TextAction
              tone="primary"
              onClick={() => setReplying(true)}
              testId="pane-comment-reply"
            >
              Reply
            </TextAction>
          )}
          <TextAction onClick={() => onResolve(!root.done)} testId="pane-comment-resolve">
            {root.done ? 'Reopen' : 'Resolve'}
          </TextAction>
          <TextAction tone="danger" onClick={() => onDelete(root.id)} testId="pane-comment-delete">
            Delete
          </TextAction>
        </div>
      )}

      {replying && (
        <div className="mt-2" onClick={(e) => e.stopPropagation()}>
          <Composer
            placeholder="Reply…"
            submitLabel="Reply"
            autoFocus
            testId="pane-reply-composer"
            onSubmit={(text) => {
              onReply(text);
              setReplying(false);
            }}
            onCancel={() => setReplying(false)}
          />
        </div>
      )}
    </div>
  );
}

function CommentsTab(props: ReviewSidebarProps) {
  const { comments, anchors, activeCommentId, draft, readOnly } = props;
  const [showResolved, setShowResolved] = useState(false);

  const threads = useMemo(() => {
    const anchorById = new Map(anchors.map((a) => [a.commentId, a]));
    const order = new Map(anchors.map((a, i) => [a.commentId, i]));
    const roots = comments
      .filter((c) => c.parentId == null && anchorById.has(c.id))
      .sort((a, b) => (order.get(a.id) ?? 0) - (order.get(b.id) ?? 0));
    return roots.map((root) => ({
      root,
      anchor: anchorById.get(root.id),
      replies: comments.filter((c) => c.parentId === root.id),
    }));
  }, [comments, anchors]);

  const open = threads.filter((t) => !t.root.done);
  const resolved = threads.filter((t) => t.root.done);

  const renderThread = (t: (typeof threads)[number]) => (
    <CommentThread
      key={t.root.id}
      root={t.root}
      replies={t.replies}
      anchor={t.anchor}
      active={activeCommentId === t.root.id}
      readOnly={readOnly}
      onReply={(text) => props.onReply(t.root.id, text)}
      onResolve={(done) => props.onResolve(t.root.id, done)}
      onDelete={props.onDeleteComment}
      onSelect={() => props.onSelectComment(t.root.id)}
    />
  );

  return (
    <div className="flex flex-col gap-3">
      {draft && !readOnly && (
        <div className="rounded-lg border border-blue-400 bg-white p-3 shadow-md">
          <p className="text-xs text-slate-500 border-l-2 border-amber-400 pl-2 mb-2 line-clamp-2">
            {draft.quote}
          </p>
          <Composer
            placeholder="Add a comment…"
            submitLabel="Comment"
            autoFocus
            testId="pane-comment-composer"
            onSubmit={props.onSubmitDraft}
            onCancel={props.onCancelDraft}
          />
        </div>
      )}

      {open.length === 0 && !draft && (
        <EmptyState>
          {readOnly
            ? 'No open comments.'
            : 'No comments yet. Select text and click Comment to start a discussion.'}
        </EmptyState>
      )}
      {open.map(renderThread)}

      {resolved.length > 0 && (
        <button
          type="button"
          onClick={() => setShowResolved((v) => !v)}
          className="text-xs font-medium text-slate-500 hover:text-slate-800 text-left px-1"
        >
          {showResolved ? 'Hide' : 'Show'} {resolved.length} resolved
        </button>
      )}
      {showResolved && resolved.map(renderThread)}
    </div>
  );
}

// ============================================================================
// SUGGESTIONS TAB
// ============================================================================

function SuggestionsTab(props: ReviewSidebarProps) {
  const { revisions, readOnly } = props;

  if (revisions.length === 0) {
    return (
      <EmptyState>
        No suggestions. Switch to <strong>Suggesting</strong> mode to propose edits as tracked
        changes.
      </EmptyState>
    );
  }

  return (
    <div className="flex flex-col gap-3">
      {!readOnly && (
        <div className="flex gap-2">
          <button
            type="button"
            data-testid="accept-all"
            onClick={props.onAcceptAll}
            className="flex-1 h-8 rounded-md text-sm font-medium bg-emerald-600 text-white hover:bg-emerald-700"
          >
            Accept all
          </button>
          <button
            type="button"
            data-testid="reject-all"
            onClick={props.onRejectAll}
            className="flex-1 h-8 rounded-md text-sm font-medium border border-slate-300 text-slate-700 hover:bg-slate-50"
          >
            Reject all
          </button>
        </div>
      )}

      {revisions.map((rev) => (
        <div
          key={`${rev.key}@${rev.from}`}
          onClick={() => props.onSelectRevision(rev)}
          data-testid="suggestion-card"
          className="rounded-lg border border-slate-200 bg-white p-3 cursor-pointer hover:shadow-sm"
        >
          <div className="flex gap-2">
            <Avatar name={rev.author} />
            <div className="min-w-0 flex-1">
              <div className="flex items-baseline gap-2">
                <span className="text-sm font-semibold text-slate-900 truncate">{rev.author}</span>
                <span className="text-xs text-slate-500 flex-shrink-0">{formatDate(rev.date)}</span>
              </div>
              <p className="text-sm mt-0.5 break-words">
                <span className="text-slate-500">
                  {rev.type === 'insertion' ? 'Add: ' : 'Delete: '}
                </span>
                <span
                  className={
                    rev.type === 'insertion'
                      ? 'text-emerald-800 underline'
                      : 'text-red-700 line-through'
                  }
                >
                  {rev.text.length > 140 ? `${rev.text.slice(0, 140)}…` : rev.text}
                </span>
              </p>
            </div>
            {!readOnly && (
              <div className="flex gap-0.5 flex-shrink-0">
                <button
                  type="button"
                  title="Accept"
                  aria-label="Accept suggestion"
                  data-testid="accept-suggestion"
                  onClick={(e) => {
                    e.stopPropagation();
                    props.onAcceptRevision(rev.key);
                  }}
                  className="w-7 h-7 rounded-md inline-flex items-center justify-center text-emerald-700 hover:bg-emerald-50"
                >
                  <svg
                    width="18"
                    height="18"
                    viewBox="0 0 24 24"
                    fill="currentColor"
                    aria-hidden="true"
                  >
                    <path d="M9 16.17 4.83 12l-1.42 1.41L9 19 21 7l-1.41-1.41z" />
                  </svg>
                </button>
                <button
                  type="button"
                  title="Reject"
                  aria-label="Reject suggestion"
                  data-testid="reject-suggestion"
                  onClick={(e) => {
                    e.stopPropagation();
                    props.onRejectRevision(rev.key);
                  }}
                  className="w-7 h-7 rounded-md inline-flex items-center justify-center text-red-600 hover:bg-red-50"
                >
                  <svg
                    width="18"
                    height="18"
                    viewBox="0 0 24 24"
                    fill="currentColor"
                    aria-hidden="true"
                  >
                    <path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z" />
                  </svg>
                </button>
              </div>
            )}
          </div>
        </div>
      ))}
    </div>
  );
}

// ============================================================================
// TAGS TAB
// ============================================================================

function TagForm({
  initial,
  submitLabel,
  onSubmit,
  onCancel,
}: {
  initial?: { tag: string; alias: string; lock?: SdtLock | null };
  submitLabel: string;
  onSubmit: (attrs: ContentControlAttrs) => void;
  onCancel: () => void;
}) {
  const initialFlags = sdtLockFlags(initial?.lock);
  const [alias, setAlias] = useState(initial?.alias ?? '');
  const [tag, setTag] = useState(initial?.tag ?? '');
  const [cannotDelete, setCannotDelete] = useState(initialFlags.cannotDelete);
  const [cannotEditContents, setCannotEditContents] = useState(initialFlags.cannotEditContents);
  const aliasRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    aliasRef.current?.focus();
  }, []);

  const inputClass =
    'w-full h-8 rounded-md border border-slate-300 px-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500/40 focus:border-blue-500';

  return (
    <form
      data-testid="tag-form"
      className="flex flex-col gap-2"
      onClick={(e) => e.stopPropagation()}
      onSubmit={(e) => {
        e.preventDefault();
        if (!alias.trim() && !tag.trim()) return;
        onSubmit({
          alias: alias.trim(),
          tag: tag.trim(),
          lock: sdtLockFromFlags(cannotDelete, cannotEditContents) ?? null,
        });
      }}
    >
      <label className="text-xs font-medium text-slate-600">
        Title
        <input
          ref={aliasRef}
          value={alias}
          onChange={(e) => setAlias(e.target.value)}
          placeholder="e.g. Client name"
          className={cn(inputClass, 'mt-1 font-normal')}
          data-testid="tag-title-input"
        />
      </label>
      <label className="text-xs font-medium text-slate-600">
        Tag
        <input
          value={tag}
          onChange={(e) => setTag(e.target.value)}
          onKeyDown={(e) => e.key === 'Escape' && onCancel()}
          placeholder="e.g. client_name"
          className={cn(inputClass, 'mt-1 font-mono font-normal')}
          data-testid="tag-tag-input"
        />
      </label>
      <fieldset
        className="rounded-md border border-slate-200 px-2.5 pb-2 pt-1"
        data-testid="tag-lock-options"
      >
        <legend className="px-1 text-xs font-medium text-slate-600">Locking</legend>
        <label className="mt-1 flex items-center gap-2 text-sm text-slate-700 cursor-pointer">
          <input
            type="checkbox"
            checked={cannotDelete}
            onChange={(e) => setCannotDelete(e.target.checked)}
            data-testid="tag-lock-cannot-delete"
          />
          Content control cannot be deleted
        </label>
        <label className="mt-1.5 flex items-center gap-2 text-sm text-slate-700 cursor-pointer">
          <input
            type="checkbox"
            checked={cannotEditContents}
            onChange={(e) => setCannotEditContents(e.target.checked)}
            data-testid="tag-lock-cannot-edit"
          />
          Contents cannot be edited
        </label>
      </fieldset>
      <div className="flex justify-end gap-1.5">
        <button
          type="button"
          onClick={onCancel}
          className="h-7 px-3 rounded-md text-sm text-slate-600 hover:bg-slate-100"
        >
          Cancel
        </button>
        <button
          type="submit"
          disabled={!alias.trim() && !tag.trim()}
          className="h-7 px-3 rounded-md text-sm font-medium bg-blue-600 text-white hover:bg-blue-700 disabled:opacity-40"
        >
          {submitLabel}
        </button>
      </div>
    </form>
  );
}

function TagsTab(props: ReviewSidebarProps) {
  const { contentControls, readOnly, tagFormOpen } = props;
  const [editingPos, setEditingPos] = useState<number | null>(null);

  return (
    <div className="flex flex-col gap-3">
      {!readOnly &&
        (tagFormOpen ? (
          <div className="rounded-lg border border-blue-400 bg-white p-3 shadow-md">
            <p className="text-xs text-slate-500 mb-2">
              Wraps the selected text in a Word content control. With no selection, a placeholder is
              inserted at the cursor.
            </p>
            <TagForm
              submitLabel="Insert tag"
              onSubmit={(attrs) => {
                props.onInsertTag(attrs);
                props.onTagFormOpenChange(false);
              }}
              onCancel={() => props.onTagFormOpenChange(false)}
            />
          </div>
        ) : (
          <button
            type="button"
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => props.onTagFormOpenChange(true)}
            className="h-8 rounded-md text-sm font-medium border border-dashed border-slate-300 text-slate-700 hover:bg-slate-50"
            data-testid="new-tag"
          >
            + Insert Word tag
          </button>
        ))}

      {contentControls.length === 0 && <EmptyState>No Word tags in this document.</EmptyState>}

      {contentControls.map((cc) => (
        <div
          key={cc.pos}
          onClick={() => props.onSelectTag(cc)}
          data-testid="tag-card"
          className="rounded-lg border border-slate-200 bg-white p-3 cursor-pointer hover:shadow-sm"
        >
          {editingPos === cc.pos ? (
            <TagForm
              initial={{ tag: cc.tag ?? '', alias: cc.alias ?? '', lock: cc.lock }}
              submitLabel="Save"
              onSubmit={(attrs) => {
                props.onUpdateTag(cc.pos, attrs);
                setEditingPos(null);
              }}
              onCancel={() => setEditingPos(null)}
            />
          ) : (
            <>
              <div className="flex items-baseline gap-2 flex-wrap">
                <span className="text-sm font-semibold text-slate-900 truncate">
                  {cc.alias || 'Untitled'}
                </span>
                {cc.tag && (
                  <code className="text-xs text-blue-700 bg-blue-50 rounded px-1 truncate">
                    {cc.tag}
                  </code>
                )}
                {cc.lock && cc.lock !== 'unlocked' && (
                  <span
                    className="text-[10px] uppercase tracking-wide text-amber-800 bg-amber-50 border border-amber-200 rounded px-1"
                    data-testid="tag-lock-badge"
                  >
                    Locked
                  </span>
                )}
              </div>
              <p className="text-sm text-slate-600 mt-1 line-clamp-2 break-words">
                {cc.text || <em className="text-slate-400">empty</em>}
              </p>
              {!readOnly && (
                <div className="mt-1 -ml-1.5 flex gap-1">
                  <TextAction onClick={() => setEditingPos(cc.pos)}>Edit</TextAction>
                  {!(cc.lock === 'sdtLocked' || cc.lock === 'sdtContentLocked') && (
                    <TextAction tone="danger" onClick={() => props.onRemoveTag(cc.pos)}>
                      Remove tag
                    </TextAction>
                  )}
                </div>
              )}
            </>
          )}
        </div>
      ))}
    </div>
  );
}

// ============================================================================
// SIDEBAR
// ============================================================================

export function ReviewSidebar(props: ReviewSidebarProps) {
  const { tab, onTabChange, onClose } = props;
  const openComments = props.anchors.filter((a) =>
    props.comments.some((c) => c.id === a.commentId && c.parentId == null && !c.done)
  ).length;

  const tabs: { id: ReviewTab; label: string; count: number }[] = [
    { id: 'comments', label: 'Comments', count: openComments },
    { id: 'suggestions', label: 'Suggestions', count: props.revisions.length },
    { id: 'tags', label: 'Tags', count: props.contentControls.length },
  ];

  return (
    <aside
      className="flex flex-col w-80 max-w-[85vw] flex-shrink-0 border-l border-slate-200 bg-slate-50"
      aria-label="Review"
      data-testid="review-sidebar"
    >
      <div className="flex items-center border-b border-slate-200 bg-white px-1">
        <div className="flex flex-1 min-w-0" role="tablist">
          {tabs.map((t) => (
            <button
              key={t.id}
              type="button"
              role="tab"
              aria-selected={tab === t.id}
              data-testid={`review-tab-${t.id}`}
              onClick={() => onTabChange(t.id)}
              className={cn(
                'h-10 px-2.5 text-sm font-medium border-b-2 -mb-px whitespace-nowrap',
                tab === t.id
                  ? 'border-blue-600 text-slate-900'
                  : 'border-transparent text-slate-500 hover:text-slate-800'
              )}
            >
              {t.label}
              {t.count > 0 && (
                <span className="ml-1 text-xs text-slate-500 tabular-nums">{t.count}</span>
              )}
            </button>
          ))}
        </div>
        <button
          type="button"
          onClick={onClose}
          aria-label="Close review panel"
          className="w-8 h-8 rounded-md inline-flex items-center justify-center text-slate-500 hover:bg-slate-100"
        >
          <svg width="16" height="16" viewBox="0 0 24 24" fill="currentColor" aria-hidden="true">
            <path d="M19 6.41 17.59 5 12 10.59 6.41 5 5 6.41 10.59 12 5 17.59 6.41 19 12 13.41 17.59 19 19 17.59 13.41 12z" />
          </svg>
        </button>
      </div>

      <div className="flex-1 overflow-y-auto p-3" role="tabpanel">
        {tab === 'comments' && <CommentsTab {...props} />}
        {tab === 'suggestions' && <SuggestionsTab {...props} />}
        {tab === 'tags' && <TagsTab {...props} />}
      </div>
    </aside>
  );
}
