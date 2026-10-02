/**
 * Read remote peer caret positions from Yjs awareness and map them for
 * the paged editor's visual overlay (HiddenProseMirror is off-screen).
 */

import * as Y from 'yjs';
import type { EditorView } from 'prosemirror-view';
import {
  absolutePositionToRelativePosition,
  relativePositionToAbsolutePosition,
  ySyncPluginKey,
} from 'y-prosemirror';

import type { CollaborationSession } from './types';

/** Visual remote caret for the selection overlay */
export interface RemoteCursorOverlay {
  clientId: number;
  /** Display initials (e.g. "AB") */
  initials: string;
  /** Full display name (tooltip) */
  name: string;
  /** CSS color */
  color: string;
  /** Caret head PM position */
  head: number;
  /** Selection anchor PM position */
  anchor: number;
}

/** Initials from a display name — "Alice Bob" → "AB", "Alice" → "AL" */
export function initialsForName(name: string): string {
  const parts = name.trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return '?';
  if (parts.length === 1) {
    const w = parts[0];
    return w.length === 1 ? w.toUpperCase() : w.slice(0, 2).toUpperCase();
  }
  return (parts[0].charAt(0) + parts[parts.length - 1].charAt(0)).toUpperCase();
}

/**
 * Build a Word-style caret DOM node with initials (for y-prosemirror decorations
 * in editors that render PM directly).
 */
export function cursorBuilderWithInitials(
  user: { name?: string; color?: string },
  _clientId?: number
): HTMLElement {
  const color = user.color && /^#[0-9a-fA-F]{6}$/.test(user.color) ? user.color : '#185abd';
  const name = user.name?.trim() || 'Guest';
  const initials = initialsForName(name);

  const cursor = document.createElement('span');
  cursor.classList.add('ProseMirror-yjs-cursor');
  cursor.setAttribute('style', `border-color: ${color}`);
  cursor.setAttribute('data-testid', 'remote-cursor');
  cursor.title = name;

  const label = document.createElement('div');
  label.classList.add('ProseMirror-yjs-cursor__label');
  label.setAttribute('style', `background-color: ${color}`);
  label.textContent = initials;

  cursor.appendChild(document.createTextNode('\u2060'));
  cursor.appendChild(label);
  cursor.appendChild(document.createTextNode('\u2060'));
  return cursor;
}

/**
 * Resolve remote awareness cursors to absolute ProseMirror positions.
 */
export function readRemoteCursors(
  session: CollaborationSession,
  view: EditorView
): RemoteCursorOverlay[] {
  const ystate = ySyncPluginKey.getState(view.state) as
    | {
        doc: Y.Doc;
        type: Y.XmlFragment;
        binding?: { mapping: Map<unknown, unknown> };
      }
    | undefined;

  if (!ystate?.binding) return [];

  const maxSize = Math.max(view.state.doc.content.size - 1, 0);
  const out: RemoteCursorOverlay[] = [];

  session.awareness.getStates().forEach((aw, clientId) => {
    if (clientId === session.doc.clientID) return;
    const cursor = (aw as { cursor?: { anchor: unknown; head: unknown } }).cursor;
    if (!cursor?.anchor || !cursor?.head) return;

    const user = (aw as { user?: { name?: string; color?: string } }).user ?? {};
    const name = user.name?.trim() || `User ${clientId}`;
    const color =
      user.color && /^#[0-9a-fA-F]{6}$/.test(user.color) ? user.color : '#185abd';

    let head = relativePositionToAbsolutePosition(
      ystate.doc,
      ystate.type,
      Y.createRelativePositionFromJSON(cursor.head),
      ystate.binding!.mapping as never
    );
    let anchor = relativePositionToAbsolutePosition(
      ystate.doc,
      ystate.type,
      Y.createRelativePositionFromJSON(cursor.anchor),
      ystate.binding!.mapping as never
    );

    if (head == null || anchor == null) return;
    head = Math.min(head, maxSize);
    anchor = Math.min(anchor, maxSize);

    out.push({
      clientId,
      initials: initialsForName(name),
      name,
      color,
      head,
      anchor,
    });
  });

  return out;
}

/**
 * Publish the local selection into awareness so peers can render our caret.
 * Call from the paged editor on selection changes — HiddenProseMirror focus
 * can be flaky for y-prosemirror's built-in hasFocus() gate.
 */
export function publishLocalCursor(session: CollaborationSession, view: EditorView): void {
  const ystate = ySyncPluginKey.getState(view.state) as
    | {
        type: Y.XmlFragment;
        binding?: { mapping: Map<unknown, unknown> };
      }
    | undefined;
  if (!ystate?.binding) return;

  const { anchor, head } = view.state.selection;
  const relAnchor = absolutePositionToRelativePosition(
    anchor,
    ystate.type,
    ystate.binding.mapping as never
  );
  const relHead = absolutePositionToRelativePosition(
    head,
    ystate.type,
    ystate.binding.mapping as never
  );

  session.awareness.setLocalStateField('cursor', {
    anchor: relAnchor,
    head: relHead,
  });
}
