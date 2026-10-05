/**
 * Deterministic room IDs. Room IDs are NOT security credentials —
 * the server must authenticate and authorize before join.
 */

export interface DocumentIdentity {
  tenantId: string;
  matterId: string;
  documentId: string;
}

/** `tenant:{tenantId}:matter:{matterId}:document:{documentId}` */
export function buildRoomId(id: DocumentIdentity): string {
  const tenantId = id.tenantId.trim();
  const matterId = id.matterId.trim();
  const documentId = id.documentId.trim();
  if (!tenantId || !matterId || !documentId) {
    throw new Error('buildRoomId: tenantId, matterId, and documentId are required');
  }
  return `tenant:${tenantId}:matter:${matterId}:document:${documentId}`;
}

/** Parse a room id produced by {@link buildRoomId}, or null if invalid */
export function parseRoomId(roomId: string): DocumentIdentity | null {
  const m = /^tenant:([^:]+):matter:([^:]+):document:(.+)$/.exec(roomId.trim());
  if (!m) return null;
  return { tenantId: m[1], matterId: m[2], documentId: m[3] };
}
