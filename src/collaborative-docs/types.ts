/**
 * Collaborative Docs Library — public types (host-agnostic).
 */

import type { StorageProvider } from './storage/StorageProvider';

export type DocumentPermission = 'read' | 'write';

export interface CollaboratingUser {
  userId: string;
  displayName: string;
  avatarUrl?: string;
  color?: string;
}

export interface DocumentAccessContext {
  userId: string;
  tenantId: string;
  matterId: string;
  documentId: string;
  roomId: string;
}

export interface AuthAdapter {
  getToken(): Promise<string>;
  getCurrentUser(): Promise<CollaboratingUser>;
}

export interface AuthorizationAdapter {
  canRead(context: DocumentAccessContext): Promise<boolean>;
  canWrite(context: DocumentAccessContext): Promise<boolean>;
}

export interface PresenceUser {
  userId: string;
  displayName: string;
  avatarUrl?: string;
  color?: string;
  cursor?: unknown;
}

export interface CollaborationConfig {
  websocket: {
    /** Base URL, e.g. `ws://localhost:1234` or `wss://collab.example.com` */
    url: string;
    reconnect?: boolean;
    reconnectAttempts?: number;
  };

  auth: AuthAdapter;

  authorization: AuthorizationAdapter;

  storage?: {
    provider: 's3' | 'local' | 'custom';
    bucket?: string;
    prefix?: string;
    /** Directory for local provider (Node server) */
    localDir?: string;
    custom?: StorageProvider;
  };

  persistence?: {
    snapshotIntervalMs?: number;
    updateBatchIntervalMs?: number;
    snapshotOnRoomClose?: boolean;
  };

  presence?: {
    enabled?: boolean;
  };

  /**
   * Max Yjs update size in bytes (large paste guard).
   * Default 5 MiB.
   */
  maxUpdateBytes?: number;
}

export interface OpenDocumentOptions {
  tenantId: string;
  matterId: string;
  documentId: string;
}

export type CollaborationEvent =
  | 'connecting'
  | 'connected'
  | 'disconnected'
  | 'reconnecting'
  | 'reconnected'
  | 'roomJoined'
  | 'roomLeft'
  | 'documentSynced'
  | 'documentChanged'
  | 'userJoined'
  | 'userLeft'
  | 'presenceChanged'
  | 'permissionDenied'
  | 'persistenceStarted'
  | 'persistenceCompleted'
  | 'persistenceFailed'
  | 'exportStarted'
  | 'exportCompleted'
  | 'error';

export type CollaborationEventHandler = (payload?: unknown) => void;
