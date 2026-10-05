/**
 * Standalone Yjs WebSocket collaboration server.
 *
 * Protocol-compatible with `y-websocket` WebsocketProvider.
 * Auth: short-lived token via `?token=` (or `Authorization` query for demos).
 * Authorization runs before room join through the host adapter.
 */

import http from 'node:http';
import { WebSocketServer, type WebSocket } from 'ws';
import * as syncProtocol from 'y-protocols/sync';
import * as awarenessProtocol from 'y-protocols/awareness';
import * as encoding from 'lib0/encoding';
import * as decoding from 'lib0/decoding';
import { RoomManager, type Room } from './RoomManager';
import type { StorageProvider } from '../storage/StorageProvider';
import type { AuthorizationAdapter, CollaboratingUser } from '../types';
import { AuthenticationError, AuthorizationError } from '../errors';

const MSG_SYNC = 0;
const MSG_AWARENESS = 1;
// const MSG_AUTH = 2;
const MSG_QUERY_AWARENESS = 3;

export interface ServerAuthResult {
  user: CollaboratingUser;
}

export interface CollaborationServerOptions {
  port?: number;
  host?: string;
  storage: StorageProvider;
  authorization: AuthorizationAdapter;
  /**
   * Validate the connection token and return the user.
   * Demo mode can accept `demo:<userId>:<displayName>`.
   */
  authenticate: (token: string) => Promise<ServerAuthResult>;
  snapshotIntervalMs?: number;
  updateBatchIntervalMs?: number;
  snapshotOnRoomClose?: boolean;
  maxUpdateBytes?: number;
  pingTimeoutMs?: number;
  log?: (fields: Record<string, unknown>) => void;
}

export interface CollaborationServer {
  readonly port: number;
  readonly roomManager: RoomManager;
  close(): Promise<void>;
}

type AuthedSocket = WebSocket & {
  __room?: Room;
  __userId?: string;
  __canWrite?: boolean;
  __awarenessClientIds?: Set<number>;
  __alive?: boolean;
};

function send(conn: WebSocket, message: Uint8Array): void {
  if (conn.readyState === conn.OPEN) {
    conn.send(message);
  }
}

function getTokenFromRequest(url: URL, protocols: string[]): string | null {
  const q = url.searchParams.get('token') || url.searchParams.get('auth');
  if (q) return q;
  // Optional: subprotocol `auth.<token>`
  for (const p of protocols) {
    if (p.startsWith('auth.')) return decodeURIComponent(p.slice(5));
  }
  return null;
}

function roomIdFromPath(pathname: string): string {
  // WebsocketProvider connects to `${serverUrl}/${roomname}`
  const parts = pathname.split('/').filter(Boolean);
  return decodeURIComponent(parts[parts.length - 1] || '');
}

export async function createCollaborationServer(
  options: CollaborationServerOptions
): Promise<CollaborationServer> {
  const port = options.port ?? 1234;
  const host = options.host ?? '0.0.0.0';
  const pingTimeoutMs = options.pingTimeoutMs ?? 30_000;
  const log =
    options.log ??
    ((fields) => {
      // Never log document contents
      console.log(JSON.stringify({ ts: new Date().toISOString(), ...fields }));
    });

  const roomManager = new RoomManager({
    storage: options.storage,
    authorization: options.authorization,
    snapshotIntervalMs: options.snapshotIntervalMs,
    updateBatchIntervalMs: options.updateBatchIntervalMs,
    snapshotOnRoomClose: options.snapshotOnRoomClose,
    maxUpdateBytes: options.maxUpdateBytes,
    log,
  });

  const server = http.createServer((_req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.end('docxsafe collaborative-docs server ok\n');
  });

  const wss = new WebSocketServer({ noServer: true });

  server.on('upgrade', (request, socket, head) => {
    const hostHeader = request.headers.host ?? 'localhost';
    const url = new URL(request.url ?? '/', `http://${hostHeader}`);
    const protocolsHeader = request.headers['sec-websocket-protocol'];
    const protocols = protocolsHeader ? protocolsHeader.split(',').map((s) => s.trim()) : [];

    wss.handleUpgrade(request, socket, head, (ws) => {
      void handleConnection(ws as AuthedSocket, url, protocols).catch((err) => {
        log({
          event: 'connection_error',
          error: err instanceof Error ? err.message : String(err),
        });
        try {
          ws.close(
            err instanceof AuthorizationError
              ? 4403
              : err instanceof AuthenticationError
                ? 4401
                : 1011,
            err instanceof Error ? err.message.slice(0, 120) : 'error'
          );
        } catch {
          // ignore
        }
      });
    });
  });

  async function handleConnection(
    conn: AuthedSocket,
    url: URL,
    protocols: string[]
  ): Promise<void> {
    const token = getTokenFromRequest(url, protocols);
    if (!token) {
      throw new AuthenticationError('Missing auth token');
    }

    const { user } = await options.authenticate(token);
    const roomId = roomIdFromPath(url.pathname);
    if (!roomId) {
      throw new AuthenticationError('Missing room id in WebSocket path');
    }

    const { canWrite, context } = await roomManager.authorizeJoin(roomId, user.userId, 'write');
    const room = await roomManager.getOrCreateRoom(roomId);

    conn.__room = room;
    conn.__userId = user.userId;
    conn.__canWrite = canWrite;
    conn.__awarenessClientIds = new Set();
    conn.__alive = true;
    room.connections.add(conn);
    room.canWrite.set(conn, canWrite);

    log({
      event: 'roomJoined',
      roomId,
      userId: user.userId,
      tenantId: context.tenantId,
      matterId: context.matterId,
      documentId: context.documentId,
      canWrite,
      connections: room.connections.size,
    });

    // Send sync step 1
    {
      const encoder = encoding.createEncoder();
      encoding.writeVarUint(encoder, MSG_SYNC);
      syncProtocol.writeSyncStep1(encoder, room.doc);
      send(conn, encoding.toUint8Array(encoder));
    }

    // Awareness: ask + send existing
    {
      const encoder = encoding.createEncoder();
      encoding.writeVarUint(encoder, MSG_QUERY_AWARENESS);
      send(conn, encoding.toUint8Array(encoder));
    }
    {
      const encoder = encoding.createEncoder();
      encoding.writeVarUint(encoder, MSG_AWARENESS);
      encoding.writeVarUint8Array(
        encoder,
        awarenessProtocol.encodeAwarenessUpdate(
          room.awareness,
          Array.from(room.awareness.getStates().keys())
        )
      );
      send(conn, encoding.toUint8Array(encoder));
    }

    const onDocUpdate = (update: Uint8Array, origin: unknown) => {
      if (origin === conn) return;
      if (update.byteLength > (options.maxUpdateBytes ?? 5 * 1024 * 1024)) {
        log({ event: 'update_skipped_size', roomId, updateSize: update.byteLength });
        return;
      }
      const encoder = encoding.createEncoder();
      encoding.writeVarUint(encoder, MSG_SYNC);
      syncProtocol.writeUpdate(encoder, update);
      send(conn, encoding.toUint8Array(encoder));
    };
    room.doc.on('update', onDocUpdate);

    const onAwarenessChange = (
      { added, updated, removed }: { added: number[]; updated: number[]; removed: number[] },
      origin: unknown
    ) => {
      if (origin === conn) return;
      const changed = added.concat(updated, removed);
      if (changed.length === 0) return;
      const encoder = encoding.createEncoder();
      encoding.writeVarUint(encoder, MSG_AWARENESS);
      encoding.writeVarUint8Array(
        encoder,
        awarenessProtocol.encodeAwarenessUpdate(room.awareness, changed)
      );
      send(conn, encoding.toUint8Array(encoder));
    };
    room.awareness.on('update', onAwarenessChange);

    conn.on('message', (data: WebSocket.RawData) => {
      try {
        const buf =
          data instanceof ArrayBuffer ? new Uint8Array(data) : new Uint8Array(data as Buffer);
        const decoder = decoding.createDecoder(buf);
        const messageType = decoding.readVarUint(decoder);

        switch (messageType) {
          case MSG_SYNC: {
            if (!conn.__canWrite) {
              // Read-only: allow sync responses that don't apply client updates.
              // syncProtocol.readSyncMessage applies updates from the client —
              // reject write updates for read-only users.
              const encoder = encoding.createEncoder();
              encoding.writeVarUint(encoder, MSG_SYNC);
              // Peek: if it's an update message, drop it
              const syncMessageType = decoding.peekVarUint(decoder);
              if (syncMessageType === syncProtocol.messageYjsUpdate) {
                log({
                  event: 'permissionDenied',
                  roomId,
                  userId: user.userId,
                  reason: 'read_only_write',
                });
                return;
              }
              syncProtocol.readSyncMessage(decoder, encoder, room.doc, conn);
              if (encoding.length(encoder) > 1) {
                send(conn, encoding.toUint8Array(encoder));
              }
              break;
            }
            {
              const encoder = encoding.createEncoder();
              encoding.writeVarUint(encoder, MSG_SYNC);
              syncProtocol.readSyncMessage(decoder, encoder, room.doc, conn);
              if (encoding.length(encoder) > 1) {
                send(conn, encoding.toUint8Array(encoder));
              }
            }
            break;
          }
          case MSG_AWARENESS: {
            const update = decoding.readVarUint8Array(decoder);
            // Track client ids from this connection for cleanup
            try {
              const dec = decoding.createDecoder(update);
              const len = decoding.readVarUint(dec);
              for (let i = 0; i < len; i++) {
                const clientId = decoding.readVarUint(dec);
                conn.__awarenessClientIds?.add(clientId);
                decoding.readVarUint(dec); // clock
                decoding.readVarUint8Array(dec); // state
              }
            } catch {
              // ignore parse errors for tracking
            }
            awarenessProtocol.applyAwarenessUpdate(room.awareness, update, conn);
            break;
          }
          case MSG_QUERY_AWARENESS: {
            const encoder = encoding.createEncoder();
            encoding.writeVarUint(encoder, MSG_AWARENESS);
            encoding.writeVarUint8Array(
              encoder,
              awarenessProtocol.encodeAwarenessUpdate(
                room.awareness,
                Array.from(room.awareness.getStates().keys())
              )
            );
            send(conn, encoding.toUint8Array(encoder));
            break;
          }
          default:
            break;
        }
      } catch (err) {
        log({
          event: 'message_error',
          roomId,
          userId: user.userId,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    });

    const cleanup = async () => {
      room.doc.off('update', onDocUpdate);
      room.awareness.off('update', onAwarenessChange);
      const ids = [...(conn.__awarenessClientIds ?? [])];
      if (ids.length > 0) {
        awarenessProtocol.removeAwarenessStates(room.awareness, ids, 'disconnect');
      }
      room.connections.delete(conn);
      room.canWrite.delete(conn);
      log({
        event: 'roomLeft',
        roomId,
        userId: user.userId,
        connections: room.connections.size,
      });
      if (room.connections.size === 0) {
        await roomManager.closeRoom(roomId);
      }
    };

    conn.on('close', () => {
      void cleanup();
    });

    const pingInterval = setInterval(() => {
      if (conn.__alive === false) {
        clearInterval(pingInterval);
        conn.terminate();
        return;
      }
      conn.__alive = false;
      try {
        conn.ping();
      } catch {
        clearInterval(pingInterval);
      }
    }, pingTimeoutMs);

    conn.on('pong', () => {
      conn.__alive = true;
    });

    conn.on('close', () => clearInterval(pingInterval));
  }

  await new Promise<void>((resolve) => {
    server.listen(port, host, () => resolve());
  });

  log({ event: 'server_listen', host, port });

  return {
    port,
    roomManager,
    async close() {
      await roomManager.destroy();
      await new Promise<void>((resolve, reject) => {
        wss.close((err) => (err ? reject(err) : resolve()));
      });
      await new Promise<void>((resolve, reject) => {
        server.close((err) => (err ? reject(err) : resolve()));
      });
    },
  };
}

/** Demo token format: `demo:<userId>:<displayName>` */
export function parseDemoToken(token: string): ServerAuthResult | null {
  if (!token.startsWith('demo:')) return null;
  const rest = token.slice(5);
  const idx = rest.indexOf(':');
  if (idx <= 0) return null;
  const userId = rest.slice(0, idx);
  const displayName = decodeURIComponent(rest.slice(idx + 1)) || userId;
  return { user: { userId, displayName } };
}
