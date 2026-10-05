/**
 * Collaborative Docs Library — Yjs + WebSocket client (browser-safe).
 *
 * Server APIs: `docxsafe-editor/collaborative-docs/server`
 *
 * @packageDocumentation
 */

export type {
  DocumentPermission,
  CollaboratingUser,
  DocumentAccessContext,
  AuthAdapter,
  AuthorizationAdapter,
  PresenceUser,
  CollaborationConfig,
  OpenDocumentOptions,
  CollaborationEvent,
  CollaborationEventHandler,
} from './types';

export {
  CollaborationError,
  AuthenticationError,
  AuthorizationError,
  ConfigurationError,
  ConnectionError,
  StorageError,
  SyncError,
  ExportError,
} from './errors';

export { buildRoomId, parseRoomId } from './roomId';
export type { DocumentIdentity } from './roomId';

export { CollaborativeDocument, createCollaborativeDocs } from './CollaborativeDocument';
export { DocumentSession } from './DocumentSession';
export type { DocumentSessionOptions } from './DocumentSession';

export type { StorageProvider } from './storage/StorageProvider';
