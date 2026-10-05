import { describe, expect, test } from 'bun:test';
import { buildRoomId, parseRoomId } from './roomId';

describe('buildRoomId / parseRoomId', () => {
  test('builds deterministic room ids', () => {
    expect(
      buildRoomId({
        tenantId: 'firm-123',
        matterId: 'M123',
        documentId: 'D456',
      })
    ).toBe('tenant:firm-123:matter:M123:document:D456');
  });

  test('round-trips', () => {
    const id = {
      tenantId: 'demo',
      matterId: 'demo-matter',
      documentId: 'demo-document',
    };
    expect(parseRoomId(buildRoomId(id))).toEqual(id);
  });

  test('rejects empty parts', () => {
    expect(() => buildRoomId({ tenantId: '', matterId: 'm', documentId: 'd' })).toThrow();
  });

  test('parse returns null for invalid', () => {
    expect(parseRoomId('not-a-room')).toBeNull();
  });
});
