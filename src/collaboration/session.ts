/**
 * Collaboration session — Y.Doc + WebSocket room (y-websocket).
 *
 * Cross-profile / cross-browser sync goes through the collaboration server
 * (`npm run collaboration:server`). Same-tab BroadcastChannel is disabled by
 * default so the server is always the source of truth.
 */

import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
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

function defaultWebsocketUrl(): string {
  if (typeof window === 'undefined') return 'ws://localhost:1234';
  const proto = window.location.protocol === 'https:' ? 'wss' : 'ws';
  // Vite proxies `/collab-ws` → collaboration server (port 1234)
  return `${proto}://${window.location.host}/collab-ws`;
}

/**
 * Create a WebSocket collaboration room. Peers that join the same `roomId`
 * (with a valid token) share one ProseMirror document via the collab server.
 *
 * @example
 * ```ts
 * const session = createCollaborationSession({
 *   roomId: 'tenant:demo:matter:m1:document:d1',
 *   user: { name: 'Alice', color: '#185abd' },
 *   websocketUrl: 'ws://localhost:1234',
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
  const websocketUrl = (options.websocketUrl ?? defaultWebsocketUrl()).replace(/\/$/, '');

  const doc = new Y.Doc();
  const fragment = doc.getXmlFragment(fragmentField);

  // Prefer a stable user id in the token so renaming does not force reconnect.
  const token = options.token ?? `demo:${userId}:${encodeURIComponent(user.name)}`;

  const provider = new WebsocketProvider(websocketUrl, roomId, doc, {
    params: { token },
    // Force server sync so different Chrome profiles work
    disableBc: options.disableBc ?? true,
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

  const onStatus = (event: { status: string }) => {
    connected = event.status === 'connected';
    for (const cb of statusListeners) cb(connected);
  };

  provider.on('status', onStatus);
  awareness.on('change', emitPeers);

  const session: CollaborationSession = {
    doc,
    provider,
    awareness,
    fragment,
    user: { ...user, id: userId },
    roomId,
    get connected() {
      return connected || provider.wsconnected;
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
      cb(connected || provider.wsconnected);
      return () => {
        statusListeners.delete(cb);
      };
    },
    destroy() {
      provider.off('status', onStatus);
      awareness.off('change', emitPeers);
      statusListeners.clear();
      peersListeners.clear();
      try {
        provider.disconnect();
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
