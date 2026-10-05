import { describe, expect, test } from 'bun:test';
import { aliasToWordTag, normalizeContentControlTagAttrs } from '../types/content';

describe('Word tag (w:tag) helpers', () => {
  test('aliasToWordTag slugs Title into a Word tag', () => {
    expect(aliasToWordTag('Client name')).toBe('client_name');
    expect(aliasToWordTag('  Invoice #42  ')).toBe('invoice_42');
  });

  test('normalizeContentControlTagAttrs prefers explicit tag', () => {
    expect(normalizeContentControlTagAttrs({ tag: 'sku', alias: 'Product SKU' })).toEqual({
      tag: 'sku',
      alias: 'Product SKU',
    });
  });

  test('normalizeContentControlTagAttrs derives w:tag from alias', () => {
    expect(normalizeContentControlTagAttrs({ tag: '', alias: 'Client name' })).toEqual({
      tag: 'client_name',
      alias: 'Client name',
    });
  });

  test('normalizeContentControlTagAttrs allocates unique tag_N', () => {
    expect(normalizeContentControlTagAttrs({ tag: '', alias: '' }, ['tag_1'])).toEqual({
      tag: 'tag_2',
      alias: '',
    });
  });
});
