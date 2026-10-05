/**
 * Wait until a collaboration room is ready to seed local content.
 *
 * Peers may still be empty for a moment after join (WebSocket sync in flight).
 * Seeding too early can overwrite a peer's document with an empty local doc.
 */

import type { CollaborationSession } from './types';

export interface WhenRoomReadyOptions {
  /**
   * If no sync event arrives, seed after this many ms.
   * Default 800 — enough for a local WebSocket server round-trip.
   */
  timeoutMs?: number;
}

/**
 * Invoke `cb` once the provider reports synced, or after `timeoutMs`.
 * Returns a cancel function.
 */
export function whenRoomReady(
  session: CollaborationSession,
  cb: () => void,
  options: WhenRoomReadyOptions = {}
): () => void {
  const timeoutMs = options.timeoutMs ?? 800;
  let settled = false;
  let timeoutId: ReturnType<typeof setTimeout> | undefined;
  const provider = session.provider;

  const onSync = (synced: boolean) => {
    if (synced) settle();
  };

  const cleanup = () => {
    provider.off('sync', onSync);
    if (timeoutId !== undefined) clearTimeout(timeoutId);
  };

  const settle = () => {
    if (settled) return;
    settled = true;
    cleanup();
    cb();
  };

  // Already synced — brief defer so in-flight updates apply first
  if (provider.synced) {
    timeoutId = setTimeout(settle, 50);
    return () => {
      settled = true;
      cleanup();
    };
  }

  provider.on('sync', onSync);
  // Fallback if sync is slow / solo
  timeoutId = setTimeout(settle, timeoutMs);

  return () => {
    settled = true;
    cleanup();
  };
}
