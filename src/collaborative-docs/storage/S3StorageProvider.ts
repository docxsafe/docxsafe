/**
 * S3-compatible durable storage. Credentials stay on the server only.
 *
 * Uses the AWS SDK if available at runtime; otherwise throws a clear error
 * so local/demo mode can use LocalStorageProvider instead.
 */

import type { StorageProvider } from './StorageProvider';
import { StorageError } from '../errors';

export interface S3StorageProviderOptions {
  bucket: string;
  prefix?: string;
  region?: string;
  endpoint?: string;
  accessKeyId?: string;
  secretAccessKey?: string;
  forcePathStyle?: boolean;
}

type S3ClientLike = {
  send(command: unknown): Promise<unknown>;
};

/**
 * Lazy S3 provider — imports `@aws-sdk/client-s3` only when constructed.
 * Install `@aws-sdk/client-s3` in the host when using provider: "s3".
 */
export class S3StorageProvider implements StorageProvider {
  private client: S3ClientLike | null = null;
  private readonly bucket: string;
  private readonly prefix: string;
  private readonly opts: S3StorageProviderOptions;

  constructor(opts: S3StorageProviderOptions) {
    if (!opts.bucket?.trim()) {
      throw new StorageError('S3StorageProvider: bucket is required');
    }
    this.bucket = opts.bucket.trim();
    this.prefix = (opts.prefix ?? 'collab').replace(/\/$/, '');
    this.opts = opts;
  }

  private key(documentId: string, ...parts: string[]): string {
    const safe = documentId.replace(/[^a-zA-Z0-9._:-]+/g, '_');
    return [this.prefix, 'documents', safe, ...parts].join('/');
  }

  private async getClient(): Promise<S3ClientLike> {
    if (this.client) return this.client;
    try {
      // Dynamic import so the editor package does not hard-require AWS SDK
      const aws = await import('@aws-sdk/client-s3' as string);
      const { S3Client } = aws as {
        S3Client: new (c: Record<string, unknown>) => S3ClientLike;
      };
      this.client = new S3Client({
        region: this.opts.region ?? 'us-east-1',
        endpoint: this.opts.endpoint,
        forcePathStyle: this.opts.forcePathStyle ?? Boolean(this.opts.endpoint),
        credentials:
          this.opts.accessKeyId && this.opts.secretAccessKey
            ? {
                accessKeyId: this.opts.accessKeyId,
                secretAccessKey: this.opts.secretAccessKey,
              }
            : undefined,
      });
      return this.client;
    } catch {
      throw new StorageError(
        'S3StorageProvider requires @aws-sdk/client-s3. Install it or use LocalStorageProvider.'
      );
    }
  }

  async saveSnapshot(documentId: string, snapshot: Uint8Array): Promise<void> {
    const client = await this.getClient();
    const aws = await import('@aws-sdk/client-s3' as string);
    const { PutObjectCommand } = aws as {
      PutObjectCommand: new (i: Record<string, unknown>) => unknown;
    };
    const body = Buffer.from(snapshot);
    const ts = String(Date.now());
    await client.send(
      new PutObjectCommand({
        Bucket: this.bucket,
        Key: this.key(documentId, 'yjs', 'snapshots', 'latest.bin'),
        Body: body,
      })
    );
    await client.send(
      new PutObjectCommand({
        Bucket: this.bucket,
        Key: this.key(documentId, 'yjs', 'snapshots', `${ts}.bin`),
        Body: body,
      })
    );
  }

  async loadLatestSnapshot(documentId: string): Promise<Uint8Array | null> {
    try {
      const client = await this.getClient();
      const aws = await import('@aws-sdk/client-s3' as string);
      const { GetObjectCommand } = aws as {
        GetObjectCommand: new (i: Record<string, unknown>) => unknown;
      };
      const res = (await client.send(
        new GetObjectCommand({
          Bucket: this.bucket,
          Key: this.key(documentId, 'yjs', 'snapshots', 'latest.bin'),
        })
      )) as { Body?: { transformToByteArray?: () => Promise<Uint8Array> } };
      if (!res.Body?.transformToByteArray) return null;
      return await res.Body.transformToByteArray();
    } catch {
      return null;
    }
  }

  async saveUpdateBatch(documentId: string, updates: Uint8Array): Promise<void> {
    const client = await this.getClient();
    const aws = await import('@aws-sdk/client-s3' as string);
    const { PutObjectCommand } = aws as {
      PutObjectCommand: new (i: Record<string, unknown>) => unknown;
    };
    const name = `${Date.now()}-${Math.random().toString(36).slice(2, 8)}.bin`;
    await client.send(
      new PutObjectCommand({
        Bucket: this.bucket,
        Key: this.key(documentId, 'yjs', 'updates', name),
        Body: Buffer.from(updates),
      })
    );
  }

  async loadUpdatesAfterSnapshot(documentId: string): Promise<Uint8Array[]> {
    try {
      const client = await this.getClient();
      const aws = await import('@aws-sdk/client-s3' as string);
      const { ListObjectsV2Command, GetObjectCommand } = aws as {
        ListObjectsV2Command: new (i: Record<string, unknown>) => unknown;
        GetObjectCommand: new (i: Record<string, unknown>) => unknown;
      };
      const prefix = this.key(documentId, 'yjs', 'updates') + '/';
      const listed = (await client.send(
        new ListObjectsV2Command({ Bucket: this.bucket, Prefix: prefix })
      )) as { Contents?: Array<{ Key?: string }> };
      const keys = (listed.Contents ?? [])
        .map((c) => c.Key)
        .filter((k): k is string => Boolean(k))
        .sort();
      const out: Uint8Array[] = [];
      for (const key of keys) {
        const res = (await client.send(
          new GetObjectCommand({ Bucket: this.bucket, Key: key })
        )) as { Body?: { transformToByteArray?: () => Promise<Uint8Array> } };
        if (res.Body?.transformToByteArray) {
          out.push(await res.Body.transformToByteArray());
        }
      }
      return out;
    } catch {
      return [];
    }
  }

  async saveExport(documentId: string, filename: string, content: Uint8Array): Promise<string> {
    const client = await this.getClient();
    const aws = await import('@aws-sdk/client-s3' as string);
    const { PutObjectCommand } = aws as {
      PutObjectCommand: new (i: Record<string, unknown>) => unknown;
    };
    const safeName = filename.replace(/[^a-zA-Z0-9._-]+/g, '_');
    const key = this.key(documentId, 'exports', safeName);
    await client.send(
      new PutObjectCommand({
        Bucket: this.bucket,
        Key: key,
        Body: Buffer.from(content),
      })
    );
    await client.send(
      new PutObjectCommand({
        Bucket: this.bucket,
        Key: this.key(documentId, 'exports', 'latest.docx'),
        Body: Buffer.from(content),
      })
    );
    return `s3://${this.bucket}/${key}`;
  }
}
