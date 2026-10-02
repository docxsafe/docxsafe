/**
 * Compact Word-style presence strip: local user + remote peers in the room.
 */

import React from 'react';
import type { CollaborationAwarenessUser, CollaborationUser } from './types';

export interface CollaborationPresenceProps {
  /** Local user */
  user: CollaborationUser;
  /** Remote peers */
  peers: CollaborationAwarenessUser[];
  /** Signaling connected */
  connected?: boolean;
  /** Room label (optional) */
  roomId?: string;
  className?: string;
}

function Avatar({ name, color }: { name: string; color: string }) {
  const initial = (name.trim().charAt(0) || '?').toUpperCase();
  return (
    <span
      className="ep-collab-avatar"
      style={{ background: color }}
      title={name}
      aria-hidden="true"
    >
      {initial}
    </span>
  );
}

export function CollaborationPresence({
  user,
  peers,
  connected = false,
  roomId,
  className = '',
}: CollaborationPresenceProps): React.ReactElement {
  return (
    <div
      className={`ep-collab-presence ${className}`.trim()}
      data-testid="collab-presence"
      role="status"
      aria-live="polite"
    >
      <span
        className={`ep-collab-presence__status${connected ? ' is-connected' : ''}`}
        title={connected ? 'Connected to room' : 'Connecting…'}
      />
      <div className="ep-collab-presence__avatars">
        <Avatar name={user.name} color={user.color} />
        {peers.map((p) => (
          <Avatar key={p.id} name={p.name} color={p.color} />
        ))}
      </div>
      <span className="ep-collab-presence__label">
        {peers.length === 0
          ? 'Waiting for peer…'
          : peers.length === 1
            ? `Editing with ${peers[0].name}`
            : `${peers.length + 1} people editing`}
        {roomId ? ` · ${roomId}` : ''}
      </span>
    </div>
  );
}
