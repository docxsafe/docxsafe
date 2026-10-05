/**
 * Filesystem persistence for local development (no AWS required).
 */

import { mkdir, readFile, writeFile, readdir, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import type { StorageProvider } from './StorageProvider';
import { StorageError } from '../errors';

export class LocalStorageProvider implements StorageProvider {
  constructor(private readonly rootDir: string) {}

  private docDir(documentId: string): string {
    // documentId may contain path-like room ids — flatten safely
    const safe = documentId.replace(/[^a-zA-Z0-9._:-]+/g, '_');
    return join(this.rootDir, safe);
  }

  private async ensureDoc(documentId: string): Promise<string> {
    const dir = this.docDir(documentId);
    await mkdir(join(dir, 'yjs', 'snapshots'), { recursive: true });
    await mkdir(join(dir, 'yjs', 'updates'), { recursive: true });
    await mkdir(join(dir, 'exports'), { recursive: true });
    return dir;
  }

  async saveSnapshot(documentId: string, snapshot: Uint8Array): Promise<void> {
    try {
      const dir = await this.ensureDoc(documentId);
      const ts = Date.now();
      const stamped = join(dir, 'yjs', 'snapshots', `${ts}.bin`);
      const latest = join(dir, 'yjs', 'snapshots', 'latest.bin');
      await writeFile(stamped, snapshot);
      await writeFile(latest, snapshot);
      // Compaction: drop update batches after snapshot
      const updatesDir = join(dir, 'yjs', 'updates');
      const files = await readdir(updatesDir).catch(() => [] as string[]);
      for (const f of files) {
        if (f.endsWith('.bin')) await unlink(join(updatesDir, f)).catch(() => undefined);
      }
    } catch (err) {
      throw new StorageError(
        `saveSnapshot failed: ${err instanceof Error ? err.message : String(err)}`
      );
    }
  }

  async loadLatestSnapshot(documentId: string): Promise<Uint8Array | null> {
    try {
      const dir = await this.ensureDoc(documentId);
      const latest = join(dir, 'yjs', 'snapshots', 'latest.bin');
      const buf = await readFile(latest);
      return new Uint8Array(buf);
    } catch {
      return null;
    }
  }

  async saveUpdateBatch(documentId: string, updates: Uint8Array): Promise<void> {
    try {
      const dir = await this.ensureDoc(documentId);
      const name = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}.bin`;
      await writeFile(join(dir, 'yjs', 'updates', name), updates);
    } catch (err) {
      throw new StorageError(
        `saveUpdateBatch failed: ${err instanceof Error ? err.message : String(err)}`
      );
    }
  }

  async loadUpdatesAfterSnapshot(documentId: string): Promise<Uint8Array[]> {
    try {
      const dir = await this.ensureDoc(documentId);
      const updatesDir = join(dir, 'yjs', 'updates');
      const files = (await readdir(updatesDir)).filter((f) => f.endsWith('.bin')).sort();
      const out: Uint8Array[] = [];
      for (const f of files) {
        const buf = await readFile(join(updatesDir, f));
        if (buf.byteLength > 0) out.push(new Uint8Array(buf));
      }
      return out;
    } catch {
      return [];
    }
  }

  async saveExport(documentId: string, filename: string, content: Uint8Array): Promise<string> {
    try {
      const dir = await this.ensureDoc(documentId);
      const safeName = filename.replace(/[^a-zA-Z0-9._-]+/g, '_');
      const target = join(dir, 'exports', safeName);
      const latest = join(dir, 'exports', 'latest.docx');
      await writeFile(target, content);
      await writeFile(latest, content);
      // Keep a versioned copy
      const versioned = join(dir, 'exports', 'versions');
      await mkdir(versioned, { recursive: true });
      await writeFile(join(versioned, `${Date.now()}-${safeName}`), content);
      return target;
    } catch (err) {
      throw new StorageError(
        `saveExport failed: ${err instanceof Error ? err.message : String(err)}`
      );
    }
  }
}
