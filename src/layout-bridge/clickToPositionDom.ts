/**
 * DOM-based Click-to-Position Mapping
 *
 * Uses the browser's actual rendered DOM to find ProseMirror positions.
 * This is more accurate than geometry-based calculation because it uses
 * the browser's own text rendering with document.elementsFromPoint().
 *
 * DOM elements are tagged with data-pm-start and data-pm-end attributes,
 * enabling binary search to find exact character positions.
 */

/**
 * Resolve the text node inside a layout run span.
 * Hyperlinks nest text under an `<a>`, so `firstChild` is not always a Text node.
 */
export function getSpanTextNode(spanEl: HTMLElement): Text | null {
  const direct = spanEl.firstChild;
  if (direct && direct.nodeType === Node.TEXT_NODE) {
    return direct as Text;
  }
  const nested = spanEl.querySelector('a')?.firstChild;
  if (nested && nested.nodeType === Node.TEXT_NODE) {
    return nested as Text;
  }
  // Fallback: first descendant text node
  const walker = spanEl.ownerDocument?.createTreeWalker(spanEl, NodeFilter.SHOW_TEXT);
  const first = walker?.nextNode();
  return first && first.nodeType === Node.TEXT_NODE ? (first as Text) : null;
}

/**
 * Caret position inside a paragraph mapped from layout `data-pm-start/end`.
 * Layout stores node boundaries (pos … pos+nodeSize); the editable caret lives
 * at pos+1. Using pmEnd put the cursor after the paragraph so empty-line clicks
 * appeared broken.
 */
export function paragraphContentPos(pmStart: number, pmEnd: number): number {
  if (!Number.isFinite(pmStart) || !Number.isFinite(pmEnd)) return 0;
  if (pmEnd > pmStart + 1) return pmStart + 1;
  return Math.max(0, pmStart);
}

function paragraphContentPosFromEl(paragraph: HTMLElement | null): number | null {
  if (!paragraph) return null;
  const pmStart = Number(paragraph.dataset.pmStart);
  const pmEnd = Number(paragraph.dataset.pmEnd);
  if (Number.isNaN(pmStart) || Number.isNaN(pmEnd)) return null;
  return paragraphContentPos(pmStart, pmEnd);
}

/**
 * Find ProseMirror position from a click using DOM-based detection.
 *
 * @param container - The pages container element
 * @param clientX - Client X coordinate from mouse event
 * @param clientY - Client Y coordinate from mouse event
 * @param zoom - Current zoom level (default 1)
 * @returns ProseMirror position, or null if not found
 */
export function clickToPositionDom(
  container: HTMLElement,
  clientX: number,
  clientY: number,
  zoom: number = 1
): number | null {
  // Get all elements at the click point
  const elements = document.elementsFromPoint(clientX, clientY);

  // Find the page element
  const pageEl = elements.find((el) => el.classList.contains('layout-page')) as HTMLElement | null;
  if (!pageEl) return null;

  // Find span with PM position data
  const spanEl = elements.find(
    (el) =>
      el.tagName === 'SPAN' &&
      (el as HTMLElement).dataset.pmStart !== undefined &&
      (el as HTMLElement).dataset.pmEnd !== undefined
  ) as HTMLElement | null;

  if (spanEl) {
    // Empty-line placeholder spans are tagged at the content caret position
    if (spanEl.classList.contains('layout-empty-run')) {
      const fromSpan = Number(spanEl.dataset.pmStart);
      if (!Number.isNaN(fromSpan)) return fromSpan;
      return paragraphContentPosFromEl(spanEl.closest('.layout-paragraph') as HTMLElement | null);
    }
    return findPositionInSpan(spanEl, clientX, clientY);
  }

  // Empty paragraphs (incl. table cells): caret goes inside the empty block
  const emptyRun = elements.find((el) =>
    el.classList.contains('layout-empty-run')
  ) as HTMLElement | null;
  if (emptyRun) {
    const fromSpan = Number(emptyRun.dataset.pmStart);
    if (!Number.isNaN(fromSpan)) return fromSpan;
    return paragraphContentPosFromEl(emptyRun.closest('.layout-paragraph') as HTMLElement | null);
  }

  // Click landed on a paragraph/line chrome (padding) — prefer that block
  const lineEl = elements.find((el) => el.classList.contains('layout-line')) as HTMLElement | null;
  if (lineEl) {
    const emptyInLine = lineEl.querySelector('.layout-empty-run') as HTMLElement | null;
    if (emptyInLine) {
      const fromSpan = Number(emptyInLine.dataset.pmStart);
      if (!Number.isNaN(fromSpan)) return fromSpan;
      return paragraphContentPosFromEl(lineEl.closest('.layout-paragraph') as HTMLElement | null);
    }
  }
  const paragraphEl = elements.find((el) =>
    el.classList.contains('layout-paragraph')
  ) as HTMLElement | null;
  if (paragraphEl?.querySelector('.layout-empty-run')) {
    return paragraphContentPosFromEl(paragraphEl);
  }

  // Whitespace / paragraph padding / empty page: Word-like end-of-line /
  // end-of-content mapping. Do not snap to paragraph start when the click
  // missed a text span (that made empty clicks jump to the first character).
  return findNearestSpan(container, pageEl, clientX, clientY, zoom);
}

/**
 * Find exact position within a text span using binary search on character boundaries.
 */
function findPositionInSpan(spanEl: HTMLElement, clientX: number, _clientY: number): number | null {
  const pmStart = Number(spanEl.dataset.pmStart);
  const pmEnd = Number(spanEl.dataset.pmEnd);

  // Special handling for tab spans - they have a visual width but only contain NBSP
  // Clicking anywhere on a tab should position cursor at start or end based on click position
  if (spanEl.classList.contains('layout-run-tab')) {
    const rect = spanEl.getBoundingClientRect();
    const midpoint = (rect.left + rect.right) / 2;
    // Click in left half -> start of tab, right half -> end of tab
    return clientX < midpoint ? pmStart : pmEnd;
  }

  const text = getSpanTextNode(spanEl);
  if (!text) {
    // No text content - return start position
    return pmStart;
  }

  const textLength = text.length;

  if (textLength === 0) {
    return pmStart;
  }

  const ownerDoc = spanEl.ownerDocument;
  if (!ownerDoc) return pmStart;

  // Binary search for the character boundary
  let left = 0;
  let right = textLength;

  while (left < right) {
    const mid = Math.floor((left + right) / 2);
    const range = ownerDoc.createRange();
    range.setStart(text, mid);
    range.setEnd(text, mid);

    const rect = range.getBoundingClientRect();
    const charX = rect.left;

    if (clientX < charX) {
      right = mid;
    } else {
      left = mid + 1;
    }
  }

  // Refine: check if we're closer to left-1 or left
  if (left > 0 && left <= textLength) {
    const range = ownerDoc.createRange();

    // Get position of character at left-1
    range.setStart(text, left - 1);
    range.setEnd(text, left - 1);
    const leftRect = range.getBoundingClientRect();

    // Get position of character at left
    range.setStart(text, Math.min(left, textLength));
    range.setEnd(text, Math.min(left, textLength));
    const rightRect = range.getBoundingClientRect();

    // Use the closer boundary
    const distLeft = Math.abs(clientX - leftRect.left);
    const distRight = Math.abs(clientX - rightRect.left);

    if (distLeft < distRight) {
      return pmStart + (left - 1);
    }
  }

  return pmStart + Math.min(left, pmEnd - pmStart);
}

/** End PM position of the last mapped span inside an element */
function endOfLastSpan(root: ParentNode): number | null {
  const spans = root.querySelectorAll('span[data-pm-start][data-pm-end]');
  if (spans.length === 0) return null;
  let maxEnd = -1;
  for (const span of Array.from(spans)) {
    const end = Number((span as HTMLElement).dataset.pmEnd);
    if (!Number.isNaN(end) && end > maxEnd) maxEnd = end;
  }
  return maxEnd >= 0 ? maxEnd : null;
}

/**
 * Find the nearest text span when click is not directly on text.
 * Word-like: right-of-line → end of line; below content → end of last content.
 */
function findNearestSpan(
  _container: HTMLElement,
  pageEl: HTMLElement,
  clientX: number,
  clientY: number,
  _zoom: number
): number | null {
  const spans = pageEl.querySelectorAll('span[data-pm-start][data-pm-end]');
  if (spans.length === 0) {
    const paragraphs = pageEl.querySelectorAll('.layout-paragraph');
    if (paragraphs.length > 0) {
      const lastP = paragraphs[paragraphs.length - 1] as HTMLElement;
      return paragraphContentPosFromEl(lastP);
    }
    return null;
  }

  const lines = Array.from(pageEl.querySelectorAll('.layout-line')) as HTMLElement[];
  if (lines.length === 0) {
    return endOfLastSpan(pageEl);
  }

  // Click clearly below the last line → end of last content
  const lastLine = lines[lines.length - 1];
  const lastLineRect = lastLine.getBoundingClientRect();
  if (clientY > lastLineRect.bottom + 4) {
    return endOfLastSpan(lastLine) ?? endOfLastSpan(pageEl);
  }

  // Prefer the line that contains the click Y; else nearest by center
  let closestLine: HTMLElement | null = null;
  let closestLineDistance = Infinity;

  for (const lineEl of lines) {
    const rect = lineEl.getBoundingClientRect();
    if (clientY >= rect.top && clientY <= rect.bottom) {
      closestLine = lineEl;
      closestLineDistance = 0;
      break;
    }
    const centerY = (rect.top + rect.bottom) / 2;
    const distance = Math.abs(clientY - centerY);
    if (distance < closestLineDistance) {
      closestLineDistance = distance;
      closestLine = lineEl;
    }
  }

  if (!closestLine) return endOfLastSpan(pageEl);

  // Empty line (placeholder NBSP only)
  const emptyInLine = closestLine.querySelector('.layout-empty-run') as HTMLElement | null;
  if (emptyInLine) {
    const fromSpan = Number(emptyInLine.dataset.pmStart);
    if (!Number.isNaN(fromSpan)) return fromSpan;
    return paragraphContentPosFromEl(
      closestLine.closest('.layout-paragraph') as HTMLElement | null
    );
  }

  const lineSpans = Array.from(
    closestLine.querySelectorAll('span[data-pm-start][data-pm-end]')
  ) as HTMLElement[];
  if (lineSpans.length === 0) {
    return paragraphContentPosFromEl(
      closestLine.closest('.layout-paragraph') as HTMLElement | null
    );
  }

  lineSpans.sort((a, b) => a.getBoundingClientRect().left - b.getBoundingClientRect().left);
  const firstSpan = lineSpans[0];
  const lastSpan = lineSpans[lineSpans.length - 1];
  const firstRect = firstSpan.getBoundingClientRect();
  const lastRect = lastSpan.getBoundingClientRect();

  // Left of line → start; right of line (whitespace) → end of last character
  if (clientX < firstRect.left) {
    return Number(firstSpan.dataset.pmStart);
  }
  if (clientX > lastRect.right) {
    return Number(lastSpan.dataset.pmEnd);
  }

  for (const spanEl of lineSpans) {
    const rect = spanEl.getBoundingClientRect();
    if (clientX >= rect.left && clientX <= rect.right) {
      return findPositionInSpan(spanEl, clientX, clientY);
    }
  }

  // Gap between spans: snap to nearer edge
  let closestSpan: HTMLElement | null = null;
  let closestSpanDistance = Infinity;
  for (const spanEl of lineSpans) {
    const rect = spanEl.getBoundingClientRect();
    const distance = clientX < rect.left ? rect.left - clientX : clientX - rect.right;
    if (distance < closestSpanDistance) {
      closestSpanDistance = distance;
      closestSpan = spanEl;
    }
  }

  if (!closestSpan) return Number(lastSpan.dataset.pmEnd);
  const rect = closestSpan.getBoundingClientRect();
  return clientX < rect.left
    ? Number(closestSpan.dataset.pmStart)
    : Number(closestSpan.dataset.pmEnd);
}

/**
 * Get selection rectangles for a PM range using DOM-based detection.
 *
 * @param container - The pages container element
 * @param from - Start PM position
 * @param to - End PM position
 * @param overlayRect - Bounding rect of the selection overlay
 * @returns Array of selection rectangles in overlay coordinates
 */
export interface DomSelectionRect {
  x: number;
  y: number;
  width: number;
  height: number;
  pageIndex: number;
}

export function getSelectionRectsFromDom(
  container: HTMLElement,
  from: number,
  to: number,
  overlayRect: DOMRect
): DomSelectionRect[] {
  const rects: DomSelectionRect[] = [];

  // Find all spans that intersect with the selection
  const spans = container.querySelectorAll('span[data-pm-start][data-pm-end]');

  for (const span of Array.from(spans)) {
    const spanEl = span as HTMLElement;
    const pmStart = Number(spanEl.dataset.pmStart);
    const pmEnd = Number(spanEl.dataset.pmEnd);

    // Check if span overlaps with selection
    if (pmEnd <= from || pmStart >= to) continue;

    const text = getSpanTextNode(spanEl);
    if (!text) continue;

    const ownerDoc = spanEl.ownerDocument;
    if (!ownerDoc) continue;

    // Calculate character range within this span
    const startChar = Math.max(0, from - pmStart);
    const endChar = Math.min(text.length, to - pmStart);

    if (startChar >= endChar) continue;

    // Create range for the selected text
    const range = ownerDoc.createRange();
    range.setStart(text, startChar);
    range.setEnd(text, endChar);

    // Get all client rects (handles line wraps)
    const clientRects = range.getClientRects();

    // Find page index
    const pageEl = spanEl.closest('.layout-page') as HTMLElement | null;
    const pageIndex = pageEl ? Number(pageEl.dataset.pageNumber || 1) - 1 : 0;

    for (const clientRect of Array.from(clientRects)) {
      rects.push({
        x: clientRect.left - overlayRect.left,
        y: clientRect.top - overlayRect.top,
        width: clientRect.width,
        height: clientRect.height,
        pageIndex,
      });
    }
  }

  return rects;
}

/**
 * Get caret position from DOM for a PM position.
 *
 * @param container - The pages container element
 * @param pmPos - ProseMirror position
 * @param overlayRect - Bounding rect of the selection overlay
 * @returns Caret position in overlay coordinates, or null
 */
export interface DomCaretPosition {
  x: number;
  y: number;
  height: number;
  pageIndex: number;
}

export function getCaretPositionFromDom(
  container: HTMLElement,
  pmPos: number,
  overlayRect: DOMRect
): DomCaretPosition | null {
  // Find span containing this position
  const spans = container.querySelectorAll('span[data-pm-start][data-pm-end]');

  for (const span of Array.from(spans)) {
    const spanEl = span as HTMLElement;
    const pmStart = Number(spanEl.dataset.pmStart);
    const pmEnd = Number(spanEl.dataset.pmEnd);

    // Special handling for tab spans - use exclusive end to avoid boundary conflicts
    // Tab at [5,6) means position 6 belongs to the next run, not the tab
    if (spanEl.classList.contains('layout-run-tab')) {
      if (pmPos >= pmStart && pmPos < pmEnd) {
        const spanRect = spanEl.getBoundingClientRect();
        const pageEl = spanEl.closest('.layout-page') as HTMLElement | null;
        const pageIndex = pageEl ? Number(pageEl.dataset.pageNumber || 1) - 1 : 0;
        const lineEl = spanEl.closest('.layout-line');
        const lineHeight = lineEl ? (lineEl as HTMLElement).offsetHeight : 16;

        // Position caret at start of tab (only position within tab)
        return {
          x: spanRect.left - overlayRect.left,
          y: spanRect.top - overlayRect.top,
          height: lineHeight,
          pageIndex,
        };
      }
      continue; // Skip to next span
    }

    // For text runs, use inclusive range
    if (pmPos >= pmStart && pmPos <= pmEnd) {
      const text = getSpanTextNode(spanEl);
      if (!text) {
        // No text - use span bounds
        const spanRect = spanEl.getBoundingClientRect();
        const pageEl = spanEl.closest('.layout-page') as HTMLElement | null;
        const pageIndex = pageEl ? Number(pageEl.dataset.pageNumber || 1) - 1 : 0;
        const lineEl = spanEl.closest('.layout-line');
        const lineHeight = lineEl ? (lineEl as HTMLElement).offsetHeight : 16;

        return {
          x: spanRect.left - overlayRect.left,
          y: spanRect.top - overlayRect.top,
          height: lineHeight,
          pageIndex,
        };
      }

      const charIndex = Math.min(pmPos - pmStart, text.length);

      const ownerDoc = spanEl.ownerDocument;
      if (!ownerDoc) continue;

      const range = ownerDoc.createRange();
      range.setStart(text, charIndex);
      range.setEnd(text, charIndex);

      let rangeRect = range.getBoundingClientRect();
      // Collapsed ranges can report an empty rect in some browsers — fall back
      // to the character box or the span bounds so remote initials stay visible.
      if (rangeRect.width === 0 && rangeRect.height === 0) {
        if (charIndex < text.length) {
          range.setStart(text, charIndex);
          range.setEnd(text, charIndex + 1);
          rangeRect = range.getBoundingClientRect();
        } else if (charIndex > 0) {
          range.setStart(text, charIndex - 1);
          range.setEnd(text, charIndex);
          const prev = range.getBoundingClientRect();
          rangeRect = new DOMRect(prev.right, prev.top, 0, prev.height);
        } else {
          rangeRect = spanEl.getBoundingClientRect();
        }
      }

      const pageEl = spanEl.closest('.layout-page') as HTMLElement | null;
      const pageIndex = pageEl ? Number(pageEl.dataset.pageNumber || 1) - 1 : 0;
      const lineEl = spanEl.closest('.layout-line');
      const lineHeight = lineEl ? (lineEl as HTMLElement).offsetHeight : 16;

      return {
        x: rangeRect.left - overlayRect.left,
        y: rangeRect.top - overlayRect.top,
        height: lineHeight,
        pageIndex,
      };
    }
  }

  // Check empty paragraphs
  const paragraphs = container.querySelectorAll('.layout-paragraph');
  for (const p of Array.from(paragraphs)) {
    const pEl = p as HTMLElement;
    const pStart = Number(pEl.dataset.pmStart);
    const pEnd = Number(pEl.dataset.pmEnd);

    if (pmPos >= pStart && pmPos <= pEnd) {
      const emptyRun = pEl.querySelector('.layout-empty-run');
      const targetEl = emptyRun || pEl;
      const rect = targetEl.getBoundingClientRect();

      const pageEl = pEl.closest('.layout-page') as HTMLElement | null;
      const pageIndex = pageEl ? Number(pageEl.dataset.pageNumber || 1) - 1 : 0;
      const lineEl = targetEl.closest('.layout-line') || targetEl;
      const lineHeight = (lineEl as HTMLElement).offsetHeight || 16;

      return {
        x: rect.left - overlayRect.left,
        y: rect.top - overlayRect.top,
        height: lineHeight,
        pageIndex,
      };
    }
  }

  return null;
}
