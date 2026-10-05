/**
 * Collaboration types — Word-style multi-user editing via Yjs + WebSocket
 */

import type * as Y from 'yjs';
import type { Awareness } from 'y-protocols/awareness';
import type { WebsocketProvider } from 'y-websocket';

/** Identity shown on remote carets / presence chips */
export interface CollaborationUser {
  /** Display name (e.g. "Alice") */
  name: string;
  /** CSS color for caret + selection (e.g. "#185abd") */
  color: string;
  /** Stable id; generated if omitted */
  id?: string;
}

/** Options for starting a WebSocket collaboration room */
export interface CollaborationOptions {
  /**
   * Shared room id. Prefer deterministic ids from
   * `buildRoomId({ tenantId, matterId, documentId })`.
   */
  roomId: string;
  /** Local user shown to peers */
  user: CollaborationUser;
  /**
   * Collaboration server base URL (no trailing room path).
   * Default: same-origin `/collab-ws` (Vite proxies to port 1234).
   */
  websocketUrl?: string;
  /**
   * Short-lived auth token for the server.
   * Default: `demo:<userId>:<name>` for local demos.
   */
  token?: string;
  /** Y.XmlFragment field name (default: `prosemirror`) */
  fragmentField?: string;
  /**
   * Disable BroadcastChannel (default true) so sync always uses the server —
   * required for different Chrome profiles / browsers.
   */
  disableBc?: boolean;
  /**
   * @deprecated WebRTC signaling is no longer used. Kept for API compatibility.
   */
  signaling?: string[];
  /**
   * @deprecated WebRTC password is no longer used.
   */
  password?: string;
  /**
   * @deprecated WebRTC maxConns is no longer used.
   */
  maxConns?: number;
}

/** Awareness payload we publish for each peer */
export interface CollaborationAwarenessUser {
  name: string;
  color: string;
  id: string;
}

/** Live collaboration session (Y.Doc + WebSocket provider + fragment) */
export interface CollaborationSession {
  /** Shared Yjs document */
  doc: Y.Doc;
  /** WebSocket provider */
  provider: WebsocketProvider;
  /** Awareness channel (cursors / presence) */
  awareness: Awareness;
  /** ProseMirror-bound XML fragment */
  fragment: Y.XmlFragment;
  /** Local user */
  user: CollaborationUser;
  /** Room id */
  roomId: string;
  /** Whether the provider reports a connection to the server */
  readonly connected: boolean;
  /** Update local awareness user fields */
  setUser(user: Partial<CollaborationUser>): void;
  /** Connected peers (excluding self) */
  getPeers(): CollaborationAwarenessUser[];
  /** Subscribe to peer list changes; returns unsubscribe */
  onPeersChange(cb: (peers: CollaborationAwarenessUser[]) => void): () => void;
  /** Subscribe to connection status; returns unsubscribe */
  onStatusChange(cb: (connected: boolean) => void): () => void;
  /** Tear down provider + doc */
  destroy(): void;
}
