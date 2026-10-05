/**
 * In-memory Y.Doc rooms with persistence + recovery.
 */

import * as Y from 'yjs';
import * as awarenessProtocol from 'y-protocols/awareness';
import type { StorageProvider } from '../storage/StorageProvider';
import type { AuthorizationAdapter, DocumentAccessContext } from '../types';
import { parseRoomId } from '../roomId';
import { AuthorizationError } from '../errors';

export interface Room {
  roomId: string;
  doc: Y.Doc;
  awareness: awarenessProtocol.Awareness;
  connections: Set<unknown>;
  canWrite: Map<unknown, boolean>;
  pendingUpdates: Uint8Array[];
  lastSnapshotAt: number;
  persistTimer: ReturnType<typeof setInterval> | null;
}

export interface RoomManagerOptions {
  storage: StorageProvider;
  authorization: AuthorizationAdapter;
  snapshotIntervalMs?: number;
  updateBatchIntervalMs?: number;
  snapshotOnRoomClose?: boolean;
  maxUpdateBytes?: number;
  log?: (fields: Record<string, unknown>) => void;
}

export class RoomManager {
  private rooms = new Map<string, Room>();
  private readonly opts: Required<
    Pick<
      RoomManagerOptions,
      'snapshotIntervalMs' | 'updateBatchIntervalMs' | 'snapshotOnRoomClose' | 'maxUpdateBytes'
    >
  > &
    RoomManagerOptions;

  constructor(options: RoomManagerOptions) {
    this.opts = {
      snapshotIntervalMs: options.snapshotIntervalMs ?? 30_000,
      updateBatchIntervalMs: options.updateBatchIntervalMs ?? 2_000,
      snapshotOnRoomClose: options.snapshotOnRoomClose ?? true,
      maxUpdateBytes: options.maxUpdateBytes ?? 5 * 1024 * 1024,
      ...options,
    };
  }

  get activeRoomCount(): number {
    return this.rooms.size;
  }

  getRoom(roomId: string): Room | undefined {
    return this.rooms.get(roomId);
  }

  async authorizeJoin(
    roomId: string,
    userId: string,
    mode: 'read' | 'write'
  ): Promise<{ canWrite: boolean; context: DocumentAccessContext }> {
    const parsed = parseRoomId(roomId);
    const context: DocumentAccessContext = parsed
      ? {
          userId,
          tenantId: parsed.tenantId,
          matterId: parsed.matterId,
          documentId: parsed.documentId,
          roomId,
        }
      : {
          // Allow plain room ids in demo mode (not tenant-scoped)
          userId,
          tenantId: 'demo',
          matterId: 'demo',
          documentId: roomId,
          roomId,
        };

    const canRead = await this.opts.authorization.canRead(context);
    if (!canRead) {
      throw new AuthorizationError('Read access denied');
    }
    const canWrite = await this.opts.authorization.canWrite(context);
    if (mode === 'write' && !canWrite) {
      // Still allow join as read-only
      this.opts.log?.({
        event: 'join_readonly',
        roomId,
        userId,
      });
    }
    return { canWrite, context };
  }

  async getOrCreateRoom(roomId: string): Promise<Room> {
    const existing = this.rooms.get(roomId);
    if (existing) return existing;

    const doc = new Y.Doc();
    const awareness = new awarenessProtocol.Awareness(doc);
    const storageKey = roomId;

    // Recovery: snapshot + update batches
    const snapshot = await this.opts.storage.loadLatestSnapshot(storageKey);
    if (snapshot) {
      Y.applyUpdate(doc, snapshot);
    }
    const batches = await this.opts.storage.loadUpdatesAfterSnapshot(storageKey);
    for (const batch of batches) {
      Y.applyUpdate(doc, batch);
    }

    const room: Room = {
      roomId,
      doc,
      awareness,
      connections: new Set(),
      canWrite: new Map(),
      pendingUpdates: [],
      lastSnapshotAt: Date.now(),
      persistTimer: null,
    };

    doc.on('update', (update: Uint8Array, origin: unknown) => {
      if (update.byteLength > this.opts.maxUpdateBytes) {
        this.opts.log?.({
          event: 'update_rejected_size',
          roomId,
          updateSize: update.byteLength,
        });
        return;
      }
      // origin === conn means from a client; still persist
      void origin;
      room.pendingUpdates.push(update);
    });

    room.persistTimer = setInterval(() => {
      void this.flushUpdates(room);
      if (Date.now() - room.lastSnapshotAt >= this.opts.snapshotIntervalMs) {
        void this.snapshot(room);
      }
    }, this.opts.updateBatchIntervalMs);

    this.rooms.set(roomId, room);
    this.opts.log?.({ event: 'room_created', roomId });
    return room;
  }

  async flushUpdates(room: Room): Promise<void> {
    if (room.pendingUpdates.length === 0) return;
    const updates = room.pendingUpdates;
    room.pendingUpdates = [];
    try {
      this.opts.log?.({
        event: 'persistenceStarted',
        roomId: room.roomId,
        updateCount: updates.length,
      });
      // Merge into one batch for storage
      const merged = Y.mergeUpdates(updates);
      await this.opts.storage.saveUpdateBatch(room.roomId, merged);
      this.opts.log?.({ event: 'persistenceCompleted', roomId: room.roomId });
    } catch (err) {
      // Put back on failure
      room.pendingUpdates.unshift(...updates);
      this.opts.log?.({
        event: 'persistenceFailed',
        roomId: room.roomId,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  async snapshot(room: Room): Promise<void> {
    try {
      await this.flushUpdates(room);
      const snap = Y.encodeStateAsUpdate(room.doc);
      await this.opts.storage.saveSnapshot(room.roomId, snap);
      room.lastSnapshotAt = Date.now();
      this.opts.log?.({
        event: 'snapshot',
        roomId: room.roomId,
        bytes: snap.byteLength,
      });
    } catch (err) {
      this.opts.log?.({
        event: 'persistenceFailed',
        roomId: room.roomId,
        error: err instanceof Error ? err.message : String(err),
      });
    }
  }

  async releaseConnection(room: Room, conn: unknown): Promise<void> {
    room.connections.delete(conn);
    room.canWrite.delete(conn);
    awarenessProtocol.removeAwarenessStates(
      room.awareness,
      [room.doc.clientID], // no-op placeholder; per-conn awareness handled by caller
      'disconnect'
    );
    if (room.connections.size === 0) {
      await this.closeRoom(room.roomId);
    }
  }

  async closeRoom(roomId: string): Promise<void> {
    const room = this.rooms.get(roomId);
    if (!room) return;
    if (room.persistTimer) clearInterval(room.persistTimer);
    if (this.opts.snapshotOnRoomClose) {
      await this.snapshot(room);
    } else {
      await this.flushUpdates(room);
    }
    room.awareness.destroy();
    room.doc.destroy();
    this.rooms.delete(roomId);
    this.opts.log?.({ event: 'room_closed', roomId });
  }

  async destroy(): Promise<void> {
    const ids = [...this.rooms.keys()];
    for (const id of ids) {
      await this.closeRoom(id);
    }
  }
}
