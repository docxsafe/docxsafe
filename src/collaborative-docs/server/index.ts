/**
 * Collaboration server package (Node.js only).
 */

export { createCollaborationServer, parseDemoToken } from './createCollaborationServer';
export type {
  CollaborationServer,
  CollaborationServerOptions,
  ServerAuthResult,
} from './createCollaborationServer';
export { RoomManager } from './RoomManager';
export type { Room, RoomManagerOptions } from './RoomManager';
export { LocalStorageProvider } from '../storage/LocalStorageProvider';
export { S3StorageProvider } from '../storage/S3StorageProvider';
export type { S3StorageProviderOptions } from '../storage/S3StorageProvider';
export type { StorageProvider } from '../storage/StorageProvider';
