/**
 * Durable persistence for Yjs snapshots / update batches / DOCX exports.
 * Credentials never leave the collaboration server.
 */

export interface StorageProvider {
  saveSnapshot(documentId: string, snapshot: Uint8Array): Promise<void>;
  loadLatestSnapshot(documentId: string): Promise<Uint8Array | null>;
  saveUpdateBatch(documentId: string, updates: Uint8Array): Promise<void>;
  loadUpdatesAfterSnapshot(documentId: string): Promise<Uint8Array[]>;
  saveExport(documentId: string, filename: string, content: Uint8Array): Promise<string>;
}
