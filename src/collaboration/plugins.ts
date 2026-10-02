/**
 * ProseMirror plugins for Yjs collaboration (sync, remote cursors, shared undo).
 *
 * Important: when these plugins are active, disable `prosemirror-history`
 * (StarterKit `history`) — y-prosemirror provides `yUndoPlugin` instead.
 */

import { keymap } from 'prosemirror-keymap';
import type { Plugin } from 'prosemirror-state';
import type { Schema, Node as PMNode } from 'prosemirror-model';
import {
  ySyncPlugin,
  yCursorPlugin,
  yUndoPlugin,
  undo,
  redo,
  undoCommand,
  redoCommand,
  defaultSelectionBuilder,
  initProseMirrorDoc,
} from 'y-prosemirror';

import { seedFragmentFromProseDoc } from './seed';
import { cursorBuilderWithInitials } from './remoteCursors';
import type { CollaborationSession } from './types';

export interface CollaborationPluginOptions {
  /**
   * Mapping from `initProseMirrorDoc` — pass when bootstrapping from a Y fragment.
   */
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  mapping?: any;
  /**
   * When true (default), bind Mod-z / Mod-y / Mod-Shift-z to y-prosemirror undo/redo.
   */
  bindUndoKeys?: boolean;
  /**
   * When true, copy `localDoc` into an empty Y fragment during bootstrap.
   * Prefer leaving this false and seeding from DocxEditor after `whenRoomReady`
   * so a late joiner does not overwrite a peer's document.
   */
  seedIfEmpty?: boolean;
}

/**
 * Build the ProseMirror plugin list for a collaboration session.
 */
export function createCollaborationPlugins(
  session: CollaborationSession,
  options: CollaborationPluginOptions = {}
): Plugin[] {
  const { bindUndoKeys = true, mapping } = options;

  const plugins: Plugin[] = [
    ySyncPlugin(session.fragment, mapping ? { mapping } : undefined),
    yCursorPlugin(session.awareness, {
      cursorBuilder: cursorBuilderWithInitials,
      selectionBuilder: defaultSelectionBuilder,
      awarenessStateFilter: (_clientId, _userId, awarenessState) =>
        awarenessState != null && typeof awarenessState === 'object' && 'user' in awarenessState,
    }),
    yUndoPlugin(),
  ];

  if (bindUndoKeys) {
    plugins.push(
      keymap({
        'Mod-z': undoCommand,
        'Mod-y': redoCommand,
        'Mod-Shift-z': redoCommand,
      })
    );
  }

  return plugins;
}

/**
 * Bind to the Y fragment and return the PM doc + collab plugins for
 * `EditorState.create`.
 *
 * Seeding is opt-in (`seedIfEmpty`). Prefer seeding from DocxEditor after
 * `whenRoomReady` so a late joiner does not overwrite a peer's document.
 */
export function bootstrapCollaborationEditor(
  session: CollaborationSession,
  localDoc: PMNode,
  schema: Schema,
  options: CollaborationPluginOptions = {}
): { doc: PMNode; plugins: Plugin[]; seeded: boolean } {
  const seeded =
    options.seedIfEmpty === true ? seedFragmentFromProseDoc(session, localDoc) : false;
  const { doc, mapping } = initProseMirrorDoc(session.fragment, schema);
  const plugins = createCollaborationPlugins(session, { ...options, mapping });
  return { doc, plugins, seeded };
}

export { undo as collabUndo, redo as collabRedo };
