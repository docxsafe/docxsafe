/**
 * Example: how DocXSafe (or any host) integrates the collaborative-docs library.
 *
 * The library must NOT import DocXSafe-specific modules — the host injects
 * auth, authorization, and storage through configuration.
 */

import { createCollaborativeDocs, type CollaborationConfig } from '../../src/collaborative-docs';

// --- Host-supplied adapters (examples only) ---------------------------------

async function firebaseAuthToken(): Promise<string> {
  // return await firebase.auth().currentUser.getIdToken();
  return 'demo:alice:Alice';
}

async function currentUser() {
  return { userId: 'alice', displayName: 'Alice', color: '#185abd' };
}

const docxsafePermissions = {
  async canRead(_userId: string, _matterId: string, _documentId: string) {
    return true;
  },
  async canWrite(_userId: string, _matterId: string, _documentId: string) {
    return true;
  },
};

// --- Configuration-only integration ----------------------------------------

const config: CollaborationConfig = {
  websocket: {
    url: process.env.COLLABORATION_WS_URL || 'ws://localhost:1234',
  },
  auth: {
    getToken: () => firebaseAuthToken(),
    getCurrentUser: () => currentUser(),
  },
  authorization: {
    canRead: async ({ userId, matterId, documentId }) =>
      docxsafePermissions.canRead(userId, matterId, documentId),
    canWrite: async ({ userId, matterId, documentId }) =>
      docxsafePermissions.canWrite(userId, matterId, documentId),
  },
  persistence: {
    snapshotIntervalMs: 30_000,
    updateBatchIntervalMs: 2_000,
    snapshotOnRoomClose: true,
  },
};

export async function openMatterDocument(matterId: string, documentId: string) {
  const collaboration = createCollaborativeDocs(config);

  const session = await collaboration.openDocument({
    tenantId: 'firm-from-host',
    matterId,
    documentId,
  });

  // Pass session.doc / session.fragment / session.awareness into DocxEditor
  // via the existing `collaboration` prop bridge (createCollaborationSession
  // already uses the same WebSocket transport).
  return session;
}
