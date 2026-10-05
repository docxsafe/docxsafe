/**
 * Configuration-driven entry point for host applications.
 *
 * Host supplies auth, authorization, and (optionally) storage config.
 * The library owns Yjs + WebSocket collaboration.
 */

import { DocumentSession } from './DocumentSession';
import type { CollaborationConfig, OpenDocumentOptions } from './types';
import { ConfigurationError } from './errors';

export class CollaborativeDocument {
  constructor(private readonly config: CollaborationConfig) {
    if (!config.websocket?.url?.trim()) {
      throw new ConfigurationError('websocket.url is required');
    }
    if (!config.auth?.getToken || !config.auth?.getCurrentUser) {
      throw new ConfigurationError('auth.getToken and auth.getCurrentUser are required');
    }
    if (!config.authorization?.canRead || !config.authorization?.canWrite) {
      throw new ConfigurationError('authorization.canRead and canWrite are required');
    }
  }

  /**
   * Open a collaborative document room.
   * Authorization is enforced on the server; the client sends a short-lived token.
   */
  async openDocument(identity: OpenDocumentOptions): Promise<DocumentSession> {
    const token = await this.config.auth.getToken();
    const user = await this.config.auth.getCurrentUser();

    // Client-side preflight (server still enforces)
    const roomId = `tenant:${identity.tenantId}:matter:${identity.matterId}:document:${identity.documentId}`;
    const canRead = await this.config.authorization.canRead({
      userId: user.userId,
      tenantId: identity.tenantId,
      matterId: identity.matterId,
      documentId: identity.documentId,
      roomId,
    });
    if (!canRead) {
      const { AuthorizationError } = await import('./errors');
      throw new AuthorizationError('Read access denied');
    }

    return new DocumentSession({
      websocketUrl: this.config.websocket.url.replace(/\/$/, ''),
      token,
      user,
      identity,
      disableBc: true,
    });
  }

  /** Alias matching the build-spec connect() shape */
  async connect(identity: OpenDocumentOptions): Promise<DocumentSession> {
    return this.openDocument(identity);
  }
}

export function createCollaborativeDocs(config: CollaborationConfig): CollaborativeDocument {
  return new CollaborativeDocument(config);
}
