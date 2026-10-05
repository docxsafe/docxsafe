/**
 * Collaboration module — real-time Word-style co-editing with Yjs + WebSocket.
 *
 * Two (or more) users join the same `roomId` via the collaboration server
 * (`npm run collaboration:server`). Works across Chrome profiles and browsers.
 * Remote carets and selections appear in the editor; undo/redo is shared via
 * y-prosemirror.
 *
 * For the host-agnostic library API see `docxsafe-editor/collaborative-docs`.
 *
 * @packageDocumentation
 *
 * @example
 * ```tsx
 * import {
 *   useCollaboration,
 *   CollaborationPresence,
 * } from 'docxsafe-editor/collaboration';
 * import 'docxsafe-editor/collaboration/styles.css';
 * import { DocxEditor } from 'docxsafe-editor';
 *
 * function App() {
 *   const collab = useCollaboration({
 *     roomId: 'contract-42',
 *     user: { name: 'Alice', color: '#185abd' },
 *     // password: 'shared-secret',
 *   });
 *
 *   return (
 *     <>
 *       {collab.session && (
 *         <CollaborationPresence
 *           user={collab.session.user}
 *           peers={collab.peers}
 *           connected={collab.connected}
 *           roomId={collab.session.roomId}
 *         />
 *       )}
 *       <DocxEditor
 *         collaboration={
 *           collab.session
 *             ? { session: collab.session, user: collab.session.user }
 *             : undefined
 *         }
 *       />
 *     </>
 *   );
 * }
 * ```
 *
 * Headless (no React hook):
 * ```ts
 * const session = createCollaborationSession({ roomId, user });
 * const plugins = createCollaborationPlugins(session);
 * // pass plugins to your EditorState; call session.destroy() on teardown
 * ```
 */

export type {
  CollaborationUser,
  CollaborationOptions,
  CollaborationSession,
  CollaborationAwarenessUser,
} from './types';

export { createCollaborationSession } from './session';
export {
  createCollaborationPlugins,
  bootstrapCollaborationEditor,
  collabUndo,
  collabRedo,
} from './plugins';
export type { CollaborationPluginOptions } from './plugins';
export {
  seedFragmentFromProseDoc,
  replaceFragmentWithProseDoc,
  proseDocFromSession,
  isFragmentSeeded,
} from './seed';
export { whenRoomReady } from './ready';
export type { WhenRoomReadyOptions } from './ready';
export { useCollaboration } from './useCollaboration';
export type { UseCollaborationResult } from './useCollaboration';
export { CollaborationPresence } from './CollaborationPresence';
export type { CollaborationPresenceProps } from './CollaborationPresence';
export { colorForUser, randomUserId } from './colors';
export {
  initialsForName,
  cursorBuilderWithInitials,
  readRemoteCursors,
  publishLocalCursor,
} from './remoteCursors';
export type { RemoteCursorOverlay } from './remoteCursors';
export {
  COMMENTS_MAP_FIELD,
  getCommentsMap,
  readSharedComments,
  upsertSharedComment,
  removeSharedComment,
  seedSharedCommentsIfEmpty,
  observeSharedComments,
  allocateCommentId,
} from './commentsSync';
export type { SharedCommentRecord } from './commentsSync';

/** Config accepted by DocxEditor's `collaboration` prop */
export interface DocxCollaborationConfig {
  /** Active session from createCollaborationSession / useCollaboration */
  session: import('./types').CollaborationSession;
  /** Local user (defaults to session.user) */
  user?: import('./types').CollaborationUser;
}
