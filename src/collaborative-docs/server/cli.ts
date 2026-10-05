#!/usr/bin/env node
/**
 * CLI: `npm run collaboration:server`
 *
 * Local demo server with LocalStorageProvider + permissive demo auth.
 */

import { mkdirSync } from 'node:fs';
import { resolve } from 'node:path';
import { createCollaborationServer, parseDemoToken } from './createCollaborationServer';
import { LocalStorageProvider } from '../storage/LocalStorageProvider';
import { AuthenticationError } from '../errors';

const port = Number(process.env.COLLAB_PORT || process.env.PORT || 1234);
const dataDir = resolve(process.env.COLLAB_DATA_DIR || '.collab-data');
mkdirSync(dataDir, { recursive: true });

const storage = new LocalStorageProvider(dataDir);

const server = await createCollaborationServer({
  port,
  storage,
  authenticate: async (token) => {
    const demo = parseDemoToken(token);
    if (demo) return demo;
    // Also accept opaque demo tokens used by the editor bridge: `user:<id>:<name>`
    if (token.startsWith('user:')) {
      const rest = token.slice(5);
      const idx = rest.indexOf(':');
      if (idx > 0) {
        return {
          user: {
            userId: rest.slice(0, idx),
            displayName: decodeURIComponent(rest.slice(idx + 1)),
          },
        };
      }
    }
    throw new AuthenticationError('Invalid token — use demo:<userId>:<name>');
  },
  authorization: {
    async canRead() {
      return true;
    },
    async canWrite() {
      return true;
    },
  },
  snapshotIntervalMs: Number(process.env.COLLAB_SNAPSHOT_MS || 30_000),
  updateBatchIntervalMs: Number(process.env.COLLAB_UPDATE_BATCH_MS || 2_000),
});

console.log(`Collaboration server listening on ws://localhost:${server.port}`);
console.log(`Persistence directory: ${dataDir}`);
console.log('Auth: demo:<userId>:<displayName>  (e.g. demo:alice:Alice)');

const shutdown = async () => {
  console.log('Shutting down…');
  await server.close();
  process.exit(0);
};
process.on('SIGINT', () => void shutdown());
process.on('SIGTERM', () => void shutdown());
