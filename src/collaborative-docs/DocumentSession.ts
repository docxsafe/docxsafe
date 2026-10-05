/**
 * Active document collaboration session (client-side).
 */

import * as Y from 'yjs';
import { WebsocketProvider } from 'y-websocket';
import type { Awareness } from 'y-protocols/awareness';
import type {
  CollaborationEvent,
  CollaborationEventHandler,
  CollaboratingUser,
  OpenDocumentOptions,
  PresenceUser,
} from './types';
import { buildRoomId } from './roomId';
import { colorForUser } from '../collaboration/colors';

export interface DocumentSessionOptions {
  websocketUrl: string;
  token: string;
  user: CollaboratingUser;
  identity: OpenDocumentOptions;
  fragmentField?: string;
  /** Disable BroadcastChannel so sync goes through the server (cross-profile). */
  disableBc?: boolean;
}

export class DocumentSession {
  readonly doc: Y.Doc;
  readonly provider: WebsocketProvider;
  readonly awareness: Awareness;
  readonly fragment: Y.XmlFragment;
  readonly roomId: string;
  readonly user: CollaboratingUser;
  private listeners = new Map<CollaborationEvent, Set<CollaborationEventHandler>>();
  private closed = false;

  constructor(options: DocumentSessionOptions) {
    this.roomId = buildRoomId(options.identity);
    this.user = options.user;
    this.doc = new Y.Doc();
    this.fragment = this.doc.getXmlFragment(options.fragmentField ?? 'prosemirror');

    const color =
      options.user.color || colorForUser(options.user.userId + options.user.displayName);

    this.provider = new WebsocketProvider(options.websocketUrl, this.roomId, this.doc, {
      params: { token: options.token },
      // Cross-profile / cross-browser must use the server, not BroadcastChannel
      disableBc: options.disableBc ?? true,
    });
    this.awareness = this.provider.awareness;

    this.awareness.setLocalStateField('user', {
      id: options.user.userId,
      name: options.user.displayName,
      color,
      avatarUrl: options.user.avatarUrl,
    });

    this.provider.on('status', (event: { status: string }) => {
      if (event.status === 'connecting') this.emit('connecting');
      if (event.status === 'connected') this.emit('connected');
      if (event.status === 'disconnected') this.emit('disconnected');
    });

    this.provider.on('sync', (synced: boolean) => {
      if (synced) {
        this.emit('documentSynced');
        this.emit('roomJoined', { roomId: this.roomId });
      }
    });

    this.awareness.on('change', () => {
      this.emit('presenceChanged', { users: this.getPresenceUsers() });
    });

    this.doc.on('update', (_u, origin) => {
      if (origin !== this.provider) {
        this.emit('documentChanged');
      }
    });
  }

  get synced(): boolean {
    return this.provider.synced;
  }

  get connected(): boolean {
    return this.provider.wsconnected;
  }

  on(event: CollaborationEvent, handler: CollaborationEventHandler): () => void {
    let set = this.listeners.get(event);
    if (!set) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(handler);
    return () => set!.delete(handler);
  }

  private emit(event: CollaborationEvent, payload?: unknown): void {
    const set = this.listeners.get(event);
    if (!set) return;
    for (const h of set) {
      try {
        h(payload);
      } catch {
        // host handlers must not break the session
      }
    }
  }

  getPresenceUsers(): PresenceUser[] {
    const users: PresenceUser[] = [];
    this.awareness.getStates().forEach((state, clientId) => {
      if (clientId === this.doc.clientID) return;
      const user = state?.user as
        { id?: string; name?: string; color?: string; avatarUrl?: string } | undefined;
      if (!user?.name) return;
      users.push({
        userId: user.id ?? String(clientId),
        displayName: user.name,
        color: user.color,
        avatarUrl: user.avatarUrl,
        cursor: state?.cursor,
      });
    });
    return users;
  }

  async close(): Promise<void> {
    if (this.closed) return;
    this.closed = true;
    this.emit('roomLeft', { roomId: this.roomId });
    try {
      this.provider.disconnect();
      this.provider.destroy();
    } catch {
      // ignore
    }
    this.doc.destroy();
    this.listeners.clear();
  }
}
