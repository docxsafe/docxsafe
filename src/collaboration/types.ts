/**
 * Collaboration types — Word-style multi-user editing via Yjs + y-webrtc
 */

import type * as Y from 'yjs';
import type { Awareness } from 'y-protocols/awareness';
import type { WebrtcProvider } from 'y-webrtc';

/** Identity shown on remote carets / presence chips */
export interface CollaborationUser {
  /** Display name (e.g. "Alice") */
  name: string;
  /** CSS color for caret + selection (e.g. "#185abd") */
  color: string;
  /** Stable id; generated if omitted */
  id?: string;
}

/** Options for starting a P2P collaboration room */
export interface CollaborationOptions {
  /**
   * Shared room id. Two browsers that join the same room sync the document
   * over WebRTC (plus BroadcastChannel for same-browser tabs).
   */
  roomId: string;
  /** Local user shown to peers */
  user: CollaborationUser;
  /**
   * Optional room password — encrypts signaling traffic so public signaling
   * servers cannot read WebRTC offer/answer payloads.
   */
  password?: string;
  /** Override default public y-webrtc signaling servers */
  signaling?: string[];
  /** Y.XmlFragment field name (default: `prosemirror`) */
  fragmentField?: string;
  /** Max WebRTC peer connections (default: provider default ~20–35) */
  maxConns?: number;
}

/** Awareness payload we publish for each peer */
export interface CollaborationAwarenessUser {
  name: string;
  color: string;
  id: string;
}

/** Live collaboration session (Y.Doc + WebRTC provider + fragment) */
export interface CollaborationSession {
  /** Shared Yjs document */
  doc: Y.Doc;
  /** WebRTC provider (signaling + peer connections) */
  provider: WebrtcProvider;
  /** Awareness channel (cursors / presence) */
  awareness: Awareness;
  /** ProseMirror-bound XML fragment */
  fragment: Y.XmlFragment;
  /** Local user */
  user: CollaborationUser;
  /** Room id */
  roomId: string;
  /** Whether the provider reports a connection to signaling */
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
