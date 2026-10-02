/**
 * Collaboration session — Y.Doc + y-webrtc room for two (or more) editors.
 */

import * as Y from 'yjs';
import { WebrtcProvider } from 'y-webrtc';
import type { Awareness } from 'y-protocols/awareness';

import { colorForUser, randomUserId } from './colors';
import type {
  CollaborationAwarenessUser,
  CollaborationOptions,
  CollaborationSession,
  CollaborationUser,
} from './types';

function normalizeUser(user: CollaborationUser): CollaborationUser {
  const id = user.id?.trim() || randomUserId();
  const name = user.name?.trim() || 'Guest';
  const color = user.color?.trim() || colorForUser(id + name);
  return { id, name, color };
}

function readPeer(state: Record<string, unknown> | undefined): CollaborationAwarenessUser | null {
  const user = state?.user as CollaborationAwarenessUser | undefined;
  if (!user || typeof user.name !== 'string' || typeof user.color !== 'string') return null;
  return {
    id: typeof user.id === 'string' ? user.id : String(state?.clientId ?? user.name),
    name: user.name,
    color: user.color,
  };
}

/**
 * Create a P2P collaboration room. Peers that join the same `roomId`
 * (and password, if set) share one ProseMirror document over WebRTC.
 *
 * @example
 * ```ts
 * const session = createCollaborationSession({
 *   roomId: 'contract-review-42',
 *   user: { name: 'Alice', color: '#185abd' },
 *   password: 'optional-secret',
 * });
 * ```
 */
export function createCollaborationSession(options: CollaborationOptions): CollaborationSession {
  if (!options.roomId?.trim()) {
    throw new Error('createCollaborationSession: roomId is required');
  }

  const user = normalizeUser(options.user);
  const userId = user.id as string;
  const roomId = options.roomId.trim();
  const fragmentField = options.fragmentField ?? 'prosemirror';

  const doc = new Y.Doc();
  const fragment = doc.getXmlFragment(fragmentField);

  // Prefer same-origin `/signaling` (Vite proxies to local y-webrtc-signaling),
  // then direct localhost, then the public fallback. Same-browser tabs also sync
  // via BroadcastChannel without signaling.
  const signaling =
    options.signaling ??
    (typeof window !== 'undefined'
      ? [
          `${window.location.protocol === 'https:' ? 'wss' : 'ws'}://${window.location.host}/signaling`,
          'ws://localhost:4444',
          'wss://y-webrtc-eu.fly.dev',
        ]
      : ['wss://y-webrtc-eu.fly.dev']);

  const provider = new WebrtcProvider(roomId, doc, {
    password: options.password,
    signaling,
    maxConns: options.maxConns,
  });

  const awareness: Awareness = provider.awareness;
  awareness.setLocalStateField('user', {
    id: userId,
    name: user.name,
    color: user.color,
  } satisfies CollaborationAwarenessUser);

  let connected = false;
  const statusListeners = new Set<(c: boolean) => void>();
  const peersListeners = new Set<(p: CollaborationAwarenessUser[]) => void>();

  const emitPeers = () => {
    const peers = listPeers(awareness, doc.clientID);
    for (const cb of peersListeners) cb(peers);
  };

  // y-webrtc status event shape differs across versions — accept either
  const onStatus = (event: { connected?: boolean; status?: string }) => {
    connected =
      typeof event.connected === 'boolean'
        ? event.connected
        : event.status === 'connected';
    for (const cb of statusListeners) cb(connected);
  };

  provider.on('status', onStatus as (arg0: { connected: boolean }) => void);
  awareness.on('change', emitPeers);

  const session: CollaborationSession = {
    doc,
    provider,
    awareness,
    fragment,
    user: { ...user, id: userId },
    roomId,
    get connected() {
      return connected;
    },
    setUser(patch) {
      const next = normalizeUser({ ...user, ...patch, id: patch.id ?? userId });
      user.name = next.name;
      user.color = next.color;
      user.id = next.id;
      session.user = { ...next, id: next.id as string };
      awareness.setLocalStateField('user', {
        id: next.id as string,
        name: next.name,
        color: next.color,
      } satisfies CollaborationAwarenessUser);
      emitPeers();
    },
    getPeers() {
      return listPeers(awareness, doc.clientID);
    },
    onPeersChange(cb) {
      peersListeners.add(cb);
      cb(listPeers(awareness, doc.clientID));
      return () => {
        peersListeners.delete(cb);
      };
    },
    onStatusChange(cb) {
      statusListeners.add(cb);
      cb(connected);
      return () => {
        statusListeners.delete(cb);
      };
    },
    destroy() {
      provider.off('status', onStatus as (arg0: { connected: boolean }) => void);
      awareness.off('change', emitPeers);
      statusListeners.clear();
      peersListeners.clear();
      try {
        provider.destroy();
      } catch {
        // provider may already be destroyed
      }
      doc.destroy();
    },
  };

  return session;
}

function listPeers(awareness: Awareness, selfId: number): CollaborationAwarenessUser[] {
  const peers: CollaborationAwarenessUser[] = [];
  awareness.getStates().forEach((state, clientId) => {
    if (clientId === selfId) return;
    const peer = readPeer(state as Record<string, unknown>);
    if (peer) peers.push(peer);
  });
  return peers;
}
