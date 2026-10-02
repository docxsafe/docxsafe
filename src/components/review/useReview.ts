/**
 * useReview — state and actions for editing modes, comments, suggestions,
 * and Word tags. DocxEditor wires the returned plugin into ProseMirror and
 * renders the Ribbon, CommentBalloons and ReviewSidebar from the returned state.
 */

import { useCallback, useMemo, useRef, useState } from 'react';
import type { Node as PMNode } from 'prosemirror-model';
import type { EditorView } from 'prosemirror-view';
import type { Comment, Document } from '../../types/document';
import { createSuggestionModePlugin } from '../../prosemirror/plugins/suggestionMode';
import { TextSelection } from 'prosemirror-state';
import {
  findAdjacentComment,
  findAdjacentRevision,
  getRevisionsAtSelection,
  acceptAllRevisions,
  acceptRevision,
  addCommentMark,
  getCommentAnchors,
  getCommentIdAtSelection,
  createSdtLockPlugin,
  getContentControls,
  getRevisions,
  insertContentControl,
  rejectAllRevisions,
  rejectRevision,
  removeCommentMark,
  removeContentControl,
  resolveCommentRange,
  selectRange,
  updateContentControl,
  type ContentControlAttrs,
  type ContentControlInfo,
  type ContentControlLock,
  type RevisionInfo,
} from '../../prosemirror/commands/review';
import { commentParagraphsFromText } from '../../docx/serializer/commentSerializer';
import type { EditorMode, MarkupView } from './types';
import type { CommentDraft, ReviewTab } from './ReviewSidebar';

export interface UseReviewOptions {
  getView: () => EditorView | null;
  /** Author name for new comments and suggestions */
  author: string;
  /** Controlled mode (optional) */
  mode?: EditorMode;
  defaultMode: EditorMode;
  onModeChange?: (mode: EditorMode) => void;
  /** Open the sidebar initially */
  defaultSidebarOpen: boolean;
}

function commentDate(): string {
  return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
}

export function useReview({
  getView: getViewOption,
  author,
  mode: controlledMode,
  defaultMode,
  onModeChange,
  defaultSidebarOpen,
}: UseReviewOptions) {
  const [uncontrolledMode, setUncontrolledMode] = useState<EditorMode>(defaultMode);
  const mode = controlledMode ?? uncontrolledMode;

  const [comments, setComments] = useState<Comment[]>([]);
  /** Whether the loaded file had comments.xml (so an empty list still gets written) */
  const hadCommentsRef = useRef(false);
  const [pmDoc, setPmDoc] = useState<PMNode | null>(null);
  const [hasSelection, setHasSelection] = useState(false);
  /** Position of the content control (sdt) containing the cursor, if any */
  const [controlAtSelection, setControlAtSelection] = useState<number | null>(null);
  const [tagDialog, setTagDialog] = useState<{
    pos: number;
    tag: string;
    alias: string;
    lock: ContentControlLock | null;
  } | null>(null);
  const [sidebarOpen, setSidebarOpen] = useState(defaultSidebarOpen);
  const [tab, setTab] = useState<ReviewTab>('comments');
  const [activeCommentId, setActiveCommentId] = useState<number | null>(null);
  const [draft, setDraft] = useState<CommentDraft | null>(null);
  const [tagFormOpen, setTagFormOpen] = useState(false);
  const [markupView, setMarkupView] = useState<MarkupView>('all');
  const [showComments, setShowComments] = useState(true);

  // Stable view getter so callbacks don't change on every render
  const getViewRef = useRef(getViewOption);
  getViewRef.current = getViewOption;
  const getView = useCallback(() => getViewRef.current(), []);

  // Refs read by the ProseMirror plugin (created once)
  const modeRef = useRef(mode);
  modeRef.current = mode;
  const authorRef = useRef(author);
  authorRef.current = author;

  const suggestionPlugin = useMemo(
    () =>
      createSuggestionModePlugin({
        isEnabled: () => modeRef.current === 'suggesting',
        getAuthor: () => authorRef.current,
      }),
    []
  );
  const sdtLockPlugin = useMemo(() => createSdtLockPlugin(), []);

  const setMode = useCallback(
    (next: EditorMode) => {
      if (controlledMode === undefined) setUncontrolledMode(next);
      onModeChange?.(next);
      setDraft(null);
      setTagFormOpen(false);
    },
    [controlledMode, onModeChange]
  );

  // --------------------------------------------------------------------------
  // Document sync
  // --------------------------------------------------------------------------

  /** Load comments from a freshly parsed document */
  const loadComments = useCallback((doc: Document | null) => {
    const loaded = doc?.package.document.comments ?? [];
    hadCommentsRef.current = !!doc?.package.document.comments;
    setComments(loaded);
    setActiveCommentId(null);
    setDraft(null);
  }, []);

  /** Called on every editor transaction/selection change */
  const syncFromView = useCallback((view: EditorView | null) => {
    if (!view) return;
    setPmDoc(view.state.doc);
    setHasSelection(!view.state.selection.empty);
    const { $from } = view.state.selection;
    let sdtPos: number | null = null;
    for (let d = $from.depth; d > 0; d--) {
      if ($from.node(d).type.name === 'sdt') {
        sdtPos = $from.before(d);
        break;
      }
    }
    setControlAtSelection(sdtPos);
    // Like Word, the comment under the cursor is active; moving away deactivates it
    setActiveCommentId(getCommentIdAtSelection(view.state));
  }, []);

  const anchors = useMemo(() => (pmDoc ? getCommentAnchors(pmDoc) : []), [pmDoc]);
  const revisions = useMemo(() => (pmDoc ? getRevisions(pmDoc) : []), [pmDoc]);
  const contentControls = useMemo(() => (pmDoc ? getContentControls(pmDoc) : []), [pmDoc]);

  /** Word-count statistics (deleted suggestions excluded, as in Word) */
  const stats = useMemo(() => {
    if (!pmDoc) return { words: 0, characters: 0, charactersNoSpaces: 0, paragraphs: 0 };
    const paragraphTexts: string[] = [];
    pmDoc.descendants((node) => {
      if (node.type.name !== 'paragraph') return true;
      let text = '';
      node.descendants((child) => {
        if (child.isText && !child.marks.some((m) => m.type.name === 'deletion')) {
          text += child.text;
        }
        return true;
      });
      paragraphTexts.push(text);
      return false;
    });
    const all = paragraphTexts.join(' ');
    return {
      words: (all.match(/\S+/g) ?? []).length,
      characters: paragraphTexts.reduce((n, t) => n + t.length, 0),
      charactersNoSpaces: all.replace(/\s/g, '').length,
      paragraphs: paragraphTexts.filter((t) => t.trim()).length,
    };
  }, [pmDoc]);

  /** Comments still anchored in the document (plus their replies) */
  const liveComments = useMemo(() => {
    const anchored = new Set(anchors.map((a) => a.commentId));
    return comments.filter((c) => anchored.has(c.parentId ?? c.id));
  }, [comments, anchors]);

  /** Merge current comments into a document before saving */
  const withComments = useCallback(
    (doc: Document): Document => {
      if (liveComments.length === 0 && !hadCommentsRef.current) return doc;
      return {
        ...doc,
        package: {
          ...doc.package,
          document: { ...doc.package.document, comments: liveComments },
        },
      };
    },
    [liveComments]
  );

  // --------------------------------------------------------------------------
  // Comments
  // --------------------------------------------------------------------------

  /**
   * Start a new comment on the current selection.
   * If the live selection collapsed (e.g. after ribbon interaction), pass a
   * saved range via `fallback`. When both are empty, expand to the word under
   * the cursor so selecting/double-clicking a word still works.
   */
  const startComment = useCallback(
    (fallback?: { from: number; to: number } | null) => {
      const view = getView();
      if (!view || mode === 'viewing') return;
      const range = resolveCommentRange(view.state, fallback);
      if (!range) return;

      // Keep PM selection aligned with the comment range (enables canComment + balloons)
      const { from: selFrom, to: selTo } = view.state.selection;
      if (selFrom !== range.from || selTo !== range.to) {
        try {
          view.dispatch(
            view.state.tr.setSelection(TextSelection.create(view.state.doc, range.from, range.to))
          );
        } catch {
          // Positions may be stale after a doc change; still open the draft
        }
      }

      setHasSelection(true);
      setDraft({
        from: range.from,
        to: range.to,
        quote: view.state.doc.textBetween(range.from, range.to, ' ', ' '),
      });
      // Comments are shown as balloons; make sure they're visible (as Word does)
      setShowComments(true);
      setMarkupView((v) => (v === 'none' || v === 'original' ? 'all' : v));
    },
    [getView, mode]
  );

  const nextCommentId = useCallback(
    () => comments.reduce((max, c) => Math.max(max, c.id), -1) + 1,
    [comments]
  );

  const submitDraft = useCallback(
    (text: string) => {
      const view = getView();
      if (!view || !draft) return;
      const id = nextCommentId();
      if (!addCommentMark(view, id, draft)) return;
      setComments((prev) => [
        ...prev,
        { id, author, date: commentDate(), content: commentParagraphsFromText(text) },
      ]);
      setActiveCommentId(id);
      setDraft(null);
    },
    [getView, draft, nextCommentId, author]
  );

  const reply = useCallback(
    (parentId: number, text: string) => {
      const id = nextCommentId();
      setComments((prev) => [
        ...prev,
        { id, parentId, author, date: commentDate(), content: commentParagraphsFromText(text) },
      ]);
    },
    [nextCommentId, author]
  );

  const resolve = useCallback((commentId: number, done: boolean) => {
    setComments((prev) => prev.map((c) => (c.id === commentId ? { ...c, done } : c)));
  }, []);

  const deleteComment = useCallback(
    (commentId: number) => {
      const target = comments.find((c) => c.id === commentId);
      if (!target) return;
      if (target.parentId == null) {
        // Only remove the anchor: the thread drops out of liveComments (and the
        // saved file), and Undo restoring the anchor brings it back intact.
        const view = getView();
        if (view) removeCommentMark(view, commentId);
        if (activeCommentId === commentId) setActiveCommentId(null);
      } else {
        setComments((prev) => prev.filter((c) => c.id !== commentId));
      }
    },
    [comments, getView, activeCommentId]
  );

  const goToComment = useCallback(
    (direction: 1 | -1) => {
      const view = getView();
      if (!view) return;
      const anchor = findAdjacentComment(anchors, view.state, direction);
      if (anchor) {
        setActiveCommentId(anchor.commentId);
        selectRange(view, anchor.from, anchor.to);
      }
    },
    [getView, anchors]
  );

  /** Delete the comment at the cursor (or the active one) */
  const deleteActiveComment = useCallback(() => {
    const view = getView();
    const id = (view && getCommentIdAtSelection(view.state)) ?? activeCommentId;
    if (id != null) deleteComment(id);
  }, [getView, activeCommentId, deleteComment]);

  const deleteAllComments = useCallback(() => {
    const view = getView();
    if (!view) return;
    for (const anchor of anchors) removeCommentMark(view, anchor.commentId);
    setActiveCommentId(null);
  }, [getView, anchors]);

  const selectComment = useCallback(
    (commentId: number) => {
      setActiveCommentId(commentId);
      const view = getView();
      const anchor = anchors.find((a) => a.commentId === commentId);
      if (view && anchor) selectRange(view, anchor.from, anchor.to);
    },
    [getView, anchors]
  );

  // --------------------------------------------------------------------------
  // Suggestions
  // --------------------------------------------------------------------------

  const withView = useCallback(
    (fn: (view: EditorView) => void) => () => {
      const view = getView();
      if (view) fn(view);
    },
    [getView]
  );

  const acceptOne = useCallback(
    (key: string) => withView((v) => acceptRevision(v, key))(),
    [withView]
  );
  const rejectOne = useCallback(
    (key: string) => withView((v) => rejectRevision(v, key))(),
    [withView]
  );
  const acceptAll = useMemo(() => withView((v) => acceptAllRevisions(v)), [withView]);
  const rejectAll = useMemo(() => withView((v) => rejectAllRevisions(v)), [withView]);
  const goToRevision = useCallback(
    (direction: 1 | -1) =>
      withView((v) => {
        const rev = findAdjacentRevision(v.state, direction);
        if (rev) selectRange(v, rev.from, rev.to);
      })(),
    [withView]
  );

  /** Accept/reject the change(s) at the selection, then move to the next one (Word default) */
  const resolveAtSelection = useCallback(
    (accept: boolean, moveNext: boolean) =>
      withView((v) => {
        const keys = [...new Set(getRevisionsAtSelection(v.state).map((r) => r.key))];
        if (keys.length === 0) {
          const next = findAdjacentRevision(v.state, 1);
          if (next) selectRange(v, next.from, next.to);
          return;
        }
        for (const key of keys) {
          if (accept) acceptRevision(v, key);
          else rejectRevision(v, key);
        }
        if (moveNext) {
          const next = findAdjacentRevision(v.state, 1);
          if (next) selectRange(v, next.from, next.to);
        }
      })(),
    [withView]
  );

  const selectRevision = useCallback(
    (rev: RevisionInfo) => withView((v) => selectRange(v, rev.from, rev.to))(),
    [withView]
  );

  // --------------------------------------------------------------------------
  // Word tags
  // --------------------------------------------------------------------------

  const openTagForm = useCallback(() => {
    if (mode === 'viewing') return;
    setSidebarOpen(true);
    setTab('tags');
    setTagFormOpen(true);
  }, [mode]);

  const insertTag = useCallback(
    (attrs: ContentControlAttrs) =>
      withView((v) => {
        v.focus();
        insertContentControl(v, attrs);
      })(),
    [withView]
  );
  /** Developer > Rich Text Content Control: wrap the selection or insert a placeholder */
  const insertControlNow = useCallback(
    () =>
      withView((v) => {
        v.focus();
        insertContentControl(v, {
          tag: '',
          alias: '',
          placeholder: 'Click or tap here to enter text.',
        });
      })(),
    [withView]
  );

  const openTagProperties = useCallback(() => {
    const view = getView();
    if (!view || controlAtSelection == null) return;
    const node = view.state.doc.nodeAt(controlAtSelection);
    if (!node) return;
    setTagDialog({
      pos: controlAtSelection,
      tag: (node.attrs.tag as string | null) ?? '',
      alias: (node.attrs.alias as string | null) ?? '',
      lock: (node.attrs.lock as ContentControlLock | null) ?? null,
    });
  }, [getView, controlAtSelection]);

  const updateTag = useCallback(
    (pos: number, attrs: ContentControlAttrs) =>
      withView((v) => updateContentControl(v, pos, attrs))(),
    [withView]
  );
  const removeTag = useCallback(
    (pos: number) => withView((v) => removeContentControl(v, pos))(),
    [withView]
  );
  const selectTag = useCallback(
    (cc: ContentControlInfo) => {
      const view = getView();
      const node = view?.state.doc.nodeAt(cc.pos);
      if (view && node) selectRange(view, cc.pos + 1, cc.pos + node.nodeSize - 1);
    },
    [getView]
  );

  // --------------------------------------------------------------------------
  // Highlight CSS for active / resolved comments
  // --------------------------------------------------------------------------

  const commentHighlightCss = useMemo(() => {
    const rules: string[] = [];
    for (const c of liveComments) {
      if (c.parentId == null && c.done) {
        rules.push(
          `.docx-editor [data-comment-ids~="${c.id}"]{background-color:transparent;border-bottom-color:transparent;}`
        );
      }
    }
    if (activeCommentId != null && (sidebarOpen || showComments)) {
      rules.push(
        `.docx-editor [data-comment-ids~="${activeCommentId}"]{background-color:rgba(255,196,0,0.55);}`
      );
    }
    return rules.join('\n');
  }, [liveComments, activeCommentId, sidebarOpen, showComments]);

  return {
    mode,
    setMode,
    suggestionPlugin,
    sdtLockPlugin,
    loadComments,
    syncFromView,
    withComments,
    hasSelection,
    markupView,
    setMarkupView,
    showComments,
    setShowComments,
    stats,
    anchors,
    liveComments,
    activeCommentId,
    setActiveCommentId,
    draft,
    submitDraft,
    cancelDraft: () => setDraft(null),
    reply,
    resolve,
    deleteComment,
    deleteActiveComment,
    deleteAllComments,
    selectComment,
    goToComment,
    goToRevision,
    acceptAtSelection: (moveNext = true) => resolveAtSelection(true, moveNext),
    rejectAtSelection: (moveNext = true) => resolveAtSelection(false, moveNext),
    acceptAll,
    rejectAll,
    revisions,
    commentHighlightCss,
    controlAtSelection,
    insertControlNow,
    openTagProperties,
    tagDialog,
    closeTagDialog: () => setTagDialog(null),
    saveTagDialog: (attrs: ContentControlAttrs) => {
      if (tagDialog) updateTag(tagDialog.pos, attrs);
      setTagDialog(null);
    },
    openPane: (paneTab: ReviewTab) => {
      setSidebarOpen(true);
      setTab(paneTab);
    },
    openCommentCount: liveComments.filter((c) => c.parentId == null && !c.done).length,
    startComment,
    openTagForm,
    sidebar: {
      open: sidebarOpen,
      setOpen: setSidebarOpen,
      props: {
        tab,
        onTabChange: setTab,
        onClose: () => setSidebarOpen(false),
        readOnly: mode === 'viewing',
        comments: liveComments,
        anchors,
        activeCommentId,
        draft,
        onSubmitDraft: submitDraft,
        onCancelDraft: () => setDraft(null),
        onReply: reply,
        onResolve: resolve,
        onDeleteComment: deleteComment,
        onSelectComment: selectComment,
        revisions,
        onAcceptRevision: acceptOne,
        onRejectRevision: rejectOne,
        onAcceptAll: acceptAll,
        onRejectAll: rejectAll,
        onSelectRevision: selectRevision,
        contentControls,
        tagFormOpen,
        onTagFormOpenChange: setTagFormOpen,
        onInsertTag: insertTag,
        onUpdateTag: updateTag,
        onRemoveTag: removeTag,
        onSelectTag: selectTag,
      },
    },
  };
}
