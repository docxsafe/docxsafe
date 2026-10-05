/**
 * React hook — create / tear down a collaboration session and its PM plugins.
 */

import { useEffect, useMemo, useState } from 'react';
import type { Plugin } from 'prosemirror-state';

import { createCollaborationSession } from './session';
import { createCollaborationPlugins } from './plugins';
import type {
  CollaborationAwarenessUser,
  CollaborationOptions,
  CollaborationSession,
} from './types';

export interface UseCollaborationResult {
  /** Active session, or null when options are null / destroyed */
  session: CollaborationSession | null;
  /** Plugins to pass to DocxEditor / PagedEditor as `externalPlugins` */
  plugins: Plugin[];
  /** Remote peers currently in the room */
  peers: CollaborationAwarenessUser[];
  /** Signaling connection status */
  connected: boolean;
}

/**
 * Join a WebSocket collaboration room for the lifetime of the component.
 *
 * Pass `null` options to disable collaboration (returns empty plugins).
 *
 * Room identity is `roomId` + `websocketUrl` — changing the display name
 * updates awareness via `setUser` without tearing down the session.
 *
 * @example
 * ```tsx
 * const collab = useCollaboration({
 *   roomId: 'acme-msa',
 *   user: { name: 'Alice', color: '#185abd' },
 *   websocketUrl: 'ws://localhost:1234',
 * });
 *
 * <DocxEditor
 *   collaboration={collab.session ? { session: collab.session } : undefined}
 * />
 * ```
 */
export function useCollaboration(
  options: CollaborationOptions | null | undefined
): UseCollaborationResult {
  const [session, setSession] = useState<CollaborationSession | null>(null);
  const [peers, setPeers] = useState<CollaborationAwarenessUser[]>([]);
  const [connected, setConnected] = useState(false);

  // Recreate only when the room / server changes.
  // Do NOT include token or display name — typing a name must not tear down the
  // WebSocket session (that kills live editing mid-keystroke).
  const roomKey = options
    ? `${options.roomId.trim()}::${options.websocketUrl ?? ''}::${options.user.id ?? ''}`
    : '';

  useEffect(() => {
    if (!options || !roomKey) {
      setSession(null);
      setPeers([]);
      setConnected(false);
      return;
    }

    const next = createCollaborationSession(options);
    setSession(next);
    setPeers(next.getPeers());
    setConnected(next.connected);

    const offPeers = next.onPeersChange(setPeers);
    const offStatus = next.onStatusChange(setConnected);

    return () => {
      offPeers();
      offStatus();
      next.destroy();
      setSession(null);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps -- roomKey captures room identity
  }, [roomKey]);

  // Name / color changes update awareness without reconnecting
  useEffect(() => {
    if (!session || !options?.user) return;
    session.setUser({
      name: options.user.name,
      color: options.user.color,
      id: options.user.id,
    });
  }, [session, options?.user.name, options?.user.color, options?.user.id]);

  const plugins = useMemo(() => (session ? createCollaborationPlugins(session) : []), [session]);

  return { session, plugins, peers, connected };
}
