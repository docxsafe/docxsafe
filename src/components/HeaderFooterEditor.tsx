/**
 * HeaderFooterEditor — Word-style in-document header/footer editing
 *
 * Edits the header/footer band in place with a real ProseMirror editor so the
 * main ribbon (bold, color, alignment, page number, …) applies here — not a
 * separate form/dialog. Click the dimmed body or “Close Header and Footer” to
 * exit (auto-saves).
 */

import React, { useRef, useEffect, useCallback, useState, useLayoutEffect } from 'react';
import type { CSSProperties } from 'react';
import { createPortal } from 'react-dom';
import { EditorState, TextSelection } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { Fragment, Slice } from 'prosemirror-model';
import { history, undo, redo } from 'prosemirror-history';
import { keymap } from 'prosemirror-keymap';
import { baseKeymap, toggleMark } from 'prosemirror-commands';

import { schema, singletonManager } from '../prosemirror/schema';
import { headerFooterToProseDoc } from '../prosemirror/conversion/toProseDoc';
import { proseDocToBlocks } from '../prosemirror/conversion/fromProseDoc';
import type { HeaderFooter, Paragraph, Table, StyleDefinitions } from '../types/document';

import 'prosemirror-view/style/prosemirror.css';

export type HeaderFooterPageNumberDesign = 'pageNumber' | 'pageXofY';

export interface HeaderFooterEditorProps {
  /** The header or footer being edited */
  headerFooter: HeaderFooter;
  /** Whether editing header or footer */
  position: 'header' | 'footer';
  /** Document styles for style resolution */
  styles?: StyleDefinitions | null;
  /** Root element that contains painted pages (used to find the zone + portal) */
  pagesRoot: HTMLElement | null;
  /** Callback when editing is complete — receives updated content blocks */
  onSave: (content: Array<Paragraph | Table>) => void;
  /** Callback when editing is cancelled without saving (Escape) */
  onClose: () => void;
  /** Fired when the nested EditorView is ready (and cleared with null on unmount) */
  onViewReady?: (view: EditorView | null) => void;
  /** Selection/formatting sync for the main ribbon */
  onSelectionChange?: (view: EditorView) => void;
}

type ZoneBox = { top: number; left: number; width: number; height: number };

/** Position the editor in viewport coordinates (fixed), over the painted HF zone */
function measureZone(pagesRoot: HTMLElement, position: 'header' | 'footer'): ZoneBox | null {
  const sel = position === 'header' ? '.layout-page-header' : '.layout-page-footer';
  const zone = pagesRoot.querySelector(sel) as HTMLElement | null;
  if (!zone) return null;
  const zoneRect = zone.getBoundingClientRect();
  return {
    top: zoneRect.top,
    left: zoneRect.left,
    width: Math.max(zoneRect.width, 120),
    height: Math.max(zoneRect.height, 48),
  };
}

/** Build PAGE / Page X of Y inline nodes for insertion at the cursor */
export function createPageNumberSlice(design: HeaderFooterPageNumberDesign): Slice {
  const field = schema.nodes.field;
  const page = field.create({
    fieldType: 'PAGE',
    instruction: ' PAGE ',
    displayText: '1',
    fieldKind: 'simple',
  });
  if (design === 'pageNumber') {
    return new Slice(Fragment.from(page), 0, 0);
  }
  const numPages = field.create({
    fieldType: 'NUMPAGES',
    instruction: ' NUMPAGES ',
    displayText: '1',
    fieldKind: 'simple',
  });
  return new Slice(
    Fragment.from([schema.text('Page '), page, schema.text(' of '), numPages]),
    0,
    0
  );
}

/** Insert a page-number field at the cursor in an HF editor view */
export function insertPageNumberInView(
  view: EditorView,
  design: HeaderFooterPageNumberDesign = 'pageNumber'
): void {
  const slice = createPageNumberSlice(design);
  const { from, to } = view.state.selection;
  // “Current Position” inserts at the caret; if a range is selected, insert after it
  // so we don't wipe the header/footer text the user just typed.
  let tr = view.state.tr;
  if (from !== to) {
    tr = tr.setSelection(TextSelection.create(tr.doc, to));
  }
  tr = tr.replaceSelection(slice).scrollIntoView();
  view.dispatch(tr);
  view.focus();
}

export function HeaderFooterEditor({
  headerFooter,
  position,
  styles,
  pagesRoot,
  onSave,
  onClose,
  onViewReady,
  onSelectionChange,
}: HeaderFooterEditorProps): React.ReactElement | null {
  const editorContainerRef = useRef<HTMLDivElement>(null);
  const viewRef = useRef<EditorView | null>(null);
  const isDirtyRef = useRef(false);
  const [box, setBox] = useState<ZoneBox | null>(null);
  const onSaveRef = useRef(onSave);
  const onCloseRef = useRef(onClose);
  const onViewReadyRef = useRef(onViewReady);
  const onSelectionChangeRef = useRef(onSelectionChange);
  onSaveRef.current = onSave;
  onCloseRef.current = onClose;
  onViewReadyRef.current = onViewReady;
  onSelectionChangeRef.current = onSelectionChange;

  // Measure / re-measure the painted header/footer zone
  useLayoutEffect(() => {
    if (!pagesRoot) return;
    const update = () => setBox(measureZone(pagesRoot, position));
    update();
    const ro = new ResizeObserver(update);
    ro.observe(pagesRoot);
    pagesRoot.addEventListener('scroll', update, { passive: true });
    window.addEventListener('resize', update);
    window.addEventListener('scroll', update, true);
    return () => {
      ro.disconnect();
      pagesRoot.removeEventListener('scroll', update);
      window.removeEventListener('resize', update);
      window.removeEventListener('scroll', update, true);
    };
  }, [pagesRoot, position]);

  // Mark pages as editing so CSS can dim the body and highlight the zone
  useEffect(() => {
    if (!pagesRoot) return;
    pagesRoot.classList.add('is-hf-editing', `is-hf-editing-${position}`);
    return () => {
      pagesRoot.classList.remove('is-hf-editing', `is-hf-editing-${position}`);
    };
  }, [pagesRoot, position]);

  const saveFromView = useCallback(() => {
    if (!viewRef.current) return;
    const blocks = proseDocToBlocks(viewRef.current.state.doc);
    onSaveRef.current(blocks);
  }, []);

  const closeAndSave = useCallback(() => {
    if (isDirtyRef.current) saveFromView();
    else onCloseRef.current();
  }, [saveFromView]);

  // Create ProseMirror editor once the zone is measured
  useEffect(() => {
    if (!editorContainerRef.current || !box) return;

    const pmDoc = headerFooterToProseDoc(headerFooter.content, {
      styles: styles || undefined,
    });

    const managerPlugins = singletonManager.getPlugins();
    const hasHistory = managerPlugins.some(
      (p) => (p as unknown as { key: string }).key === 'history$'
    );
    const plugins = [
      ...managerPlugins,
      ...(hasHistory ? [] : [history()]),
      keymap({
        ...baseKeymap,
        // Explicit marks so ribbon shortcuts work even if plugin order differs
        'Mod-b': toggleMark(schema.marks.bold),
        'Mod-i': toggleMark(schema.marks.italic),
        'Mod-u': toggleMark(schema.marks.underline),
        'Mod-z': undo,
        'Mod-y': redo,
        'Shift-Mod-z': redo,
        Escape: () => {
          onCloseRef.current();
          return true;
        },
      }),
    ];

    const state = EditorState.create({
      doc: pmDoc,
      schema,
      plugins,
    });

    const view = new EditorView(editorContainerRef.current, {
      state,
      dispatchTransaction(tr) {
        const newState = view.state.apply(tr);
        view.updateState(newState);
        if (tr.docChanged) isDirtyRef.current = true;
        if (tr.selectionSet || tr.docChanged) {
          onSelectionChangeRef.current?.(view);
        }
      },
      attributes: {
        class: 'hf-editor-pm-content',
        'aria-label': position === 'header' ? 'Header' : 'Footer',
      },
    });

    viewRef.current = view;
    onViewReadyRef.current?.(view);
    // Place caret at start and sync ribbon formatting
    try {
      const sel = TextSelection.atStart(view.state.doc);
      view.dispatch(view.state.tr.setSelection(sel));
    } catch {
      // empty doc
    }
    view.focus();
    onSelectionChangeRef.current?.(view);

    return () => {
      onViewReadyRef.current?.(null);
      view.destroy();
      viewRef.current = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- mount once per edit session
  }, [box != null, position]);

  // Clicking the dimmed document body closes HF edit (ribbon stays usable)
  useEffect(() => {
    if (!pagesRoot) return;
    const onMouseDown = (e: MouseEvent) => {
      const target = e.target as HTMLElement | null;
      if (!target) return;
      // Only treat clicks on the main body as “exit header/footer”
      if (target.closest('.layout-page-content')) {
        e.preventDefault();
        e.stopPropagation();
        closeAndSave();
      }
    };
    pagesRoot.addEventListener('mousedown', onMouseDown, true);
    return () => pagesRoot.removeEventListener('mousedown', onMouseDown, true);
  }, [pagesRoot, closeAndSave]);

  if (!pagesRoot || !box) return null;

  const label = position === 'header' ? 'Header' : 'Footer';

  const shellStyle: CSSProperties = {
    position: 'fixed',
    top: box.top,
    left: box.left,
    width: box.width,
    minHeight: box.height,
    zIndex: 40,
    display: 'flex',
    flexDirection: 'column',
    background: '#fff',
    boxSizing: 'border-box',
  };

  return createPortal(
    <div
      className={`hf-inline-editor hf-inline-editor--${position}`}
      style={shellStyle}
      data-testid="header-footer-editor"
      onMouseDown={(e) => e.stopPropagation()}
      onClick={(e) => e.stopPropagation()}
    >
      {/* Word-style corner tag — not a dialog chrome */}
      <div className="hf-inline-editor__label" aria-hidden="true">
        {label}
      </div>

      <div ref={editorContainerRef} className="hf-editor-pm" data-testid="hf-editor-pm" />

      <div className="hf-inline-editor__toolbar">
        <span className="hf-inline-editor__hint">
          Use the ribbon to format · Insert → Page Number
        </span>
        <button
          type="button"
          className="hf-inline-editor__close"
          data-testid="hf-editor-close"
          onClick={closeAndSave}
        >
          Close Header and Footer
        </button>
        {/* Keep Save test id for existing e2e — same as close (Word auto-saves) */}
        <button
          type="button"
          className="hf-inline-editor__save-hidden"
          data-testid="hf-editor-save"
          tabIndex={-1}
          aria-hidden="true"
          onClick={saveFromView}
        >
          Save
        </button>
      </div>
    </div>,
    document.body
  );
}
