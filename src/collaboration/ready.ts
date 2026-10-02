/**
 * Wait until a collaboration room is ready to seed local content.
 *
 * Peers may still be empty for a moment after join (BroadcastChannel / WebRTC
 * sync in flight). Seeding too early can overwrite a peer's document with an
 * empty local doc.
 */

import type { CollaborationSession } from './types';

export interface WhenRoomReadyOptions {
  /**
   * If no sync event arrives (solo peer), seed after this many ms.
   * Default 800 — enough for same-browser BroadcastChannel.
   */
  timeoutMs?: number;
}

/**
 * Invoke `cb` once the provider reports synced, or after `timeoutMs` if alone.
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

  const onSynced = (event: { synced: boolean } | [{ synced: boolean }]) => {
    const payload = Array.isArray(event) ? event[0] : event;
    if (payload?.synced) settle();
  };

  const cleanup = () => {
    provider.off('synced', onSynced as (arg0: { synced: boolean }) => void);
    if (timeoutId !== undefined) clearTimeout(timeoutId);
  };

  const settle = () => {
    if (settled) return;
    settled = true;
    cleanup();
    cb();
  };

  // Already synced with peers — brief defer so in-flight updates apply first
  const alreadySynced = Boolean(
    (provider as unknown as { synced?: boolean }).synced
  );
  if (alreadySynced) {
    timeoutId = setTimeout(settle, 50);
    return () => {
      settled = true;
      cleanup();
    };
  }

  provider.on('synced', onSynced as (arg0: { synced: boolean }) => void);
  // Solo peer: synced may never fire — seed after timeout
  timeoutId = setTimeout(settle, timeoutMs);

  return () => {
    settled = true;
    cleanup();
  };
}
