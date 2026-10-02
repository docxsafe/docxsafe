/**
 * Selection Overlay Component
 *
 * Renders the selection overlay for the paged editor, including:
 * - Caret cursor (blinking vertical line for collapsed selection)
 * - Selection highlights (blue rectangles for range selection)
 * - Word-style "New Comment" icon at the right edge of a text selection
 *
 * The overlay is positioned absolutely over the pages container and
 * renders selection rectangles in container-relative coordinates.
 */

import React, { useEffect, useState, useRef, useMemo } from 'react';
import type { SelectionRect, CaretPosition } from '../layout-bridge/selectionRects';

// =============================================================================
// TYPES
// =============================================================================

/** Remote collaborator caret rendered on the paged overlay */
export interface RemoteCaretMarker {
  clientId: number;
  initials: string;
  name: string;
  color: string;
  x: number;
  y: number;
  height: number;
}

export interface SelectionOverlayProps {
  /** Selection rectangles for range selection. */
  selectionRects: SelectionRect[];
  /** Caret position for collapsed selection. */
  caretPosition: CaretPosition | null;
  /** Whether the editor is focused. */
  isFocused: boolean;
  /** Hide caret/selection when in read-only mode. */
  readOnly?: boolean;
  /** Gap between pages (for coordinate adjustment). */
  pageGap?: number;
  /** Custom caret color. */
  caretColor?: string;
  /** Custom selection background color. */
  selectionColor?: string;
  /** Caret width in pixels. */
  caretWidth?: number;
  /** Blink interval in milliseconds (0 to disable). */
  blinkInterval?: number;
  /**
   * When set, show a Word-style comment icon at the right edge of the
   * highlighted selection. Hidden while a draft comment is already open.
   */
  onNewComment?: () => void;
  /** Hide the floating comment icon (e.g. draft already open / viewing mode). */
  hideCommentButton?: boolean;
  /** Remote collaborator carets with initials (collab mode). */
  remoteCarets?: RemoteCaretMarker[];
}

// =============================================================================
// CONSTANTS
// =============================================================================

const DEFAULT_CARET_COLOR = '#000';
const DEFAULT_SELECTION_COLOR = 'rgba(66, 133, 244, 0.3)'; // Google Docs style blue
const DEFAULT_CARET_WIDTH = 2;
const DEFAULT_BLINK_INTERVAL = 530; // Standard cursor blink rate

// =============================================================================
// STYLES
// =============================================================================

const overlayStyles: React.CSSProperties = {
  position: 'absolute',
  top: 0,
  left: 0,
  right: 0,
  bottom: 0,
  pointerEvents: 'none',
  zIndex: 10,
  overflow: 'hidden',
};

const caretStyles = (
  caret: CaretPosition,
  color: string,
  width: number,
  visible: boolean
): React.CSSProperties => ({
  position: 'absolute',
  left: caret.x,
  top: caret.y,
  width: width,
  height: caret.height,
  backgroundColor: color,
  opacity: visible ? 1 : 0,
  transition: 'opacity 0.05s ease-out',
  pointerEvents: 'none',
});

const selectionRectStyles = (rect: SelectionRect, color: string): React.CSSProperties => ({
  position: 'absolute',
  left: rect.x,
  top: rect.y,
  width: rect.width,
  height: rect.height,
  backgroundColor: color,
  pointerEvents: 'none',
});

const remoteCaretWrapStyles = (caret: RemoteCaretMarker): React.CSSProperties => ({
  position: 'absolute',
  left: caret.x,
  top: caret.y,
  height: caret.height,
  width: 0,
  pointerEvents: 'none',
  zIndex: 12,
});

const remoteCaretBarStyles = (color: string, height: number): React.CSSProperties => ({
  position: 'absolute',
  left: 0,
  top: 0,
  width: 2,
  height,
  backgroundColor: color,
  pointerEvents: 'none',
});

const remoteCaretLabelStyles = (color: string): React.CSSProperties => ({
  position: 'absolute',
  left: 0,
  top: -16,
  minWidth: 18,
  height: 16,
  padding: '0 4px',
  borderRadius: '3px 3px 3px 0',
  backgroundColor: color,
  color: '#fff',
  fontSize: 10,
  fontWeight: 700,
  fontFamily: "'Segoe UI', system-ui, sans-serif",
  lineHeight: '16px',
  textAlign: 'center',
  whiteSpace: 'nowrap',
  userSelect: 'none',
  pointerEvents: 'none',
});

// =============================================================================
// COMPONENT
// =============================================================================

const RemoteCaret: React.FC<{ caret: RemoteCaretMarker }> = ({ caret }) => (
  <div
    style={remoteCaretWrapStyles(caret)}
    data-testid="remote-caret"
    data-client-id={String(caret.clientId)}
    title={caret.name}
  >
    <div style={remoteCaretBarStyles(caret.color, caret.height)} />
    <div style={remoteCaretLabelStyles(caret.color)} data-testid="remote-caret-initials">
      {caret.initials}
    </div>
  </div>
);

/**
 * Caret component with blinking animation.
 */
const Caret: React.FC<{
  position: CaretPosition;
  color: string;
  width: number;
  blinkInterval: number;
  isFocused: boolean;
}> = ({ position, color, width, blinkInterval, isFocused }) => {
  const [visible, setVisible] = useState(isFocused);
  const blinkTimerRef = useRef<number | null>(null);

  useEffect(() => {
    // Clear any existing timer
    if (blinkTimerRef.current) {
      window.clearInterval(blinkTimerRef.current);
      blinkTimerRef.current = null;
    }

    // Only blink when focused and interval is set
    if (isFocused && blinkInterval > 0) {
      setVisible(true);
      blinkTimerRef.current = window.setInterval(() => {
        setVisible((v) => !v);
      }, blinkInterval);
    } else {
      // Hide caret when not focused
      setVisible(false);
    }

    return () => {
      if (blinkTimerRef.current) {
        window.clearInterval(blinkTimerRef.current);
      }
    };
  }, [isFocused, blinkInterval]);

  // Reset blink cycle when position changes (show immediately after typing/navigation)
  useEffect(() => {
    if (!isFocused) return;

    setVisible(true);

    // Restart blink timer from this moment
    if (blinkTimerRef.current) {
      window.clearInterval(blinkTimerRef.current);
    }
    if (blinkInterval > 0) {
      blinkTimerRef.current = window.setInterval(() => {
        setVisible((v) => !v);
      }, blinkInterval);
    }

    return () => {
      if (blinkTimerRef.current) {
        window.clearInterval(blinkTimerRef.current);
      }
    };
  }, [position.x, position.y, isFocused, blinkInterval]);

  return <div style={caretStyles(position, color, width, visible)} data-testid="caret" />;
};

/**
 * Selection rectangle component.
 */
const SelectionRectangle: React.FC<{
  rect: SelectionRect;
  color: string;
  index: number;
}> = ({ rect, color, index }) => {
  return (
    <div
      style={selectionRectStyles(rect, color)}
      data-testid={`selection-rect-${index}`}
      data-page-index={rect.pageIndex}
    />
  );
};

const COMMENT_BTN_SIZE = 28;
/** Inset from the page's right edge — keeps the icon on the document, clear of text. */
const COMMENT_BTN_PAGE_INSET = 8;

/** Track an in-progress pointer gesture so the icon never appears mid drag-select. */
let selectionPointerDown = false;
if (typeof window !== 'undefined') {
  window.addEventListener(
    'pointerdown',
    (e) => {
      // Clicks on the comment icon itself are not a selection gesture
      const t = e.target as HTMLElement | null;
      if (t?.closest?.('[data-testid="selection-comment-button"]')) return;
      selectionPointerDown = true;
    },
    true
  );
  window.addEventListener('pointerup', () => {
    selectionPointerDown = false;
  }, true);
  window.addEventListener('pointercancel', () => {
    selectionPointerDown = false;
  }, true);
}

/**
 * Vertically align with the end of the selection; horizontally pin to the
 * right edge of the page so the icon never sits on the highlighted text
 * (which blocked drag-select).
 */
function computeCommentAnchor(
  rects: SelectionRect[],
  overlayEl: HTMLElement | null
): { left: number; top: number } | null {
  if (rects.length === 0 || !overlayEl) return null;

  const last = rects.reduce((a, b) =>
    b.y > a.y || (b.y === a.y && b.x + b.width > a.x + a.width) ? b : a
  );

  const pagesRoot = overlayEl.parentElement?.querySelector(
    '.paged-editor__pages'
  ) as HTMLElement | null;
  const pageEls = pagesRoot?.querySelectorAll('.layout-page');
  const pageEl =
    (pageEls && pageEls[last.pageIndex]) ||
    pagesRoot?.querySelector(`.layout-page[data-page-number="${last.pageIndex + 1}"]`) ||
    null;

  let left: number;
  if (pageEl) {
    const pageRect = pageEl.getBoundingClientRect();
    const overlayRect = overlayEl.getBoundingClientRect();
    left = pageRect.right - overlayRect.left - COMMENT_BTN_SIZE - COMMENT_BTN_PAGE_INSET;
  } else {
    // Fallback: far right of the selection line (still clear of the text)
    left = Math.max(...rects.map((r) => r.x + r.width)) + 24;
  }

  return {
    left,
    top: last.y + last.height / 2 - COMMENT_BTN_SIZE / 2,
  };
}

/**
 * Word-style comment icon at the right end of the page, level with the selection.
 * Hidden while the pointer is down so it cannot interrupt drag-select.
 */
const SelectionCommentButton: React.FC<{
  rects: SelectionRect[];
  onNewComment: () => void;
  overlayEl: HTMLElement | null;
}> = ({ rects, onNewComment, overlayEl }) => {
  const [ready, setReady] = useState(false);

  // Never show while the user is still drag-selecting — the old near-text
  // placement stole pointer events mid-gesture. Wait for pointerup; for
  // keyboard / programmatic selections, reveal on the next frame.
  useEffect(() => {
    if (rects.length === 0) {
      setReady(false);
      return;
    }

    let cancelled = false;
    const reveal = () => {
      if (!cancelled && !selectionPointerDown) setReady(true);
    };

    setReady(false);

    if (selectionPointerDown) {
      const onUp = () => reveal();
      window.addEventListener('pointerup', onUp, true);
      window.addEventListener('pointercancel', onUp, true);
      return () => {
        cancelled = true;
        window.removeEventListener('pointerup', onUp, true);
        window.removeEventListener('pointercancel', onUp, true);
      };
    }

    const raf = window.requestAnimationFrame(() => reveal());
    return () => {
      cancelled = true;
      window.cancelAnimationFrame(raf);
    };
  }, [rects]);

  const anchor = useMemo(
    () => computeCommentAnchor(rects, overlayEl),
    [rects, overlayEl]
  );

  if (!ready || !anchor) return null;

  return (
    <button
      type="button"
      className="ep-selection-comment-btn"
      data-testid="selection-comment-button"
      title="New Comment"
      aria-label="New Comment"
      style={{ left: anchor.left, top: anchor.top }}
      onMouseDown={(e) => {
        // Keep the text selection when clicking the icon
        e.preventDefault();
        e.stopPropagation();
      }}
      onClick={(e) => {
        e.preventDefault();
        e.stopPropagation();
        onNewComment();
      }}
    >
      <svg width="16" height="16" viewBox="0 0 24 24" fill="none" aria-hidden="true">
        <path
          d="M4 5.5A1.5 1.5 0 0 1 5.5 4h13A1.5 1.5 0 0 1 20 5.5v9a1.5 1.5 0 0 1-1.5 1.5H10l-4 4v-4h-.5A1.5 1.5 0 0 1 4 14.5z"
          stroke="currentColor"
          strokeWidth="1.75"
          strokeLinejoin="round"
        />
        <path
          d="M12 8v5M9.5 10.5h5"
          stroke="currentColor"
          strokeWidth="1.75"
          strokeLinecap="round"
        />
      </svg>
    </button>
  );
};

/**
 * Selection overlay component.
 *
 * Renders selection highlights and caret cursor over the paginated document.
 * Should be positioned as a child of the pages container with relative positioning.
 */
export const SelectionOverlay: React.FC<SelectionOverlayProps> = ({
  selectionRects,
  caretPosition,
  isFocused,
  readOnly = false,
  caretColor = DEFAULT_CARET_COLOR,
  selectionColor = DEFAULT_SELECTION_COLOR,
  caretWidth = DEFAULT_CARET_WIDTH,
  blinkInterval = DEFAULT_BLINK_INTERVAL,
  onNewComment,
  hideCommentButton = false,
  remoteCarets = [],
}) => {
  // Callback ref so comment-button positioning re-runs once the overlay mounts
  const [overlayEl, setOverlayEl] = useState<HTMLElement | null>(null);

  if (readOnly) {
    return null;
  }
  // Determine if we have a range selection or collapsed selection
  const hasRangeSelection = selectionRects.length > 0;
  const hasCollapsedSelection = caretPosition !== null && !hasRangeSelection;
  const showCommentButton =
    hasRangeSelection && !!onNewComment && !hideCommentButton;

  return (
    <div ref={setOverlayEl} style={overlayStyles} data-testid="selection-overlay">
      {/* Render selection rectangles for range selection */}
      {hasRangeSelection &&
        selectionRects.map((rect, index) => (
          <SelectionRectangle
            key={`sel-${rect.pageIndex}-${rect.x}-${rect.y}-${index}`}
            rect={rect}
            color={selectionColor}
            index={index}
          />
        ))}

      {/* Word-style New Comment icon at the right end of the page */}
      {showCommentButton && (
        <SelectionCommentButton
          rects={selectionRects}
          onNewComment={onNewComment}
          overlayEl={overlayEl}
        />
      )}

      {/* Render caret for collapsed selection */}
      {hasCollapsedSelection && caretPosition && (
        <Caret
          position={caretPosition}
          color={caretColor}
          width={caretWidth}
          blinkInterval={blinkInterval}
          isFocused={isFocused}
        />
      )}

      {/* Remote collaborator carets + initials */}
      {remoteCarets.map((caret) => (
        <RemoteCaret key={caret.clientId} caret={caret} />
      ))}
    </div>
  );
};

// =============================================================================
// HELPER HOOKS
// =============================================================================

/**
 * Hook to manage selection overlay state.
 *
 * @param pmSelection - ProseMirror selection {from, to}.
 * @param layout - Document layout.
 * @param blocks - Flow blocks.
 * @param measures - Measurements.
 * @returns Selection overlay props.
 */
export function useSelectionOverlay(
  pmSelection: { from: number; to: number } | null,
  layout: import('../layout-engine/types').Layout | null,
  blocks: import('../layout-engine/types').FlowBlock[],
  measures: import('../layout-engine/types').Measure[]
): {
  selectionRects: SelectionRect[];
  caretPosition: CaretPosition | null;
} {
  const [selectionRects, setSelectionRects] = useState<SelectionRect[]>([]);
  const [caretPosition, setCaretPosition] = useState<CaretPosition | null>(null);

  useEffect(() => {
    if (!layout || !pmSelection) {
      setSelectionRects([]);
      setCaretPosition(null);
      return;
    }

    // Import dynamically to avoid circular dependencies
    import('../layout-bridge/selectionRects').then(({ selectionToRects, getCaretPosition }) => {
      const { from, to } = pmSelection;

      if (from === to) {
        // Collapsed selection - show caret
        const caret = getCaretPosition(layout, blocks, measures, from);
        setCaretPosition(caret);
        setSelectionRects([]);
      } else {
        // Range selection - show highlight
        const rects = selectionToRects(layout, blocks, measures, from, to);
        setSelectionRects(rects);
        setCaretPosition(null);
      }
    });
  }, [pmSelection, layout, blocks, measures]);

  return { selectionRects, caretPosition };
}

export default SelectionOverlay;
