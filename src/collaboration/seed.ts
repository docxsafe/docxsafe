/**
 * Seed / bootstrap helpers — first peer loads a DOCX into the shared Y.Doc;
 * late joiners receive the fragment over WebRTC.
 */

import type { Node as PMNode, Schema } from 'prosemirror-model';
import { prosemirrorToYXmlFragment, initProseMirrorDoc } from 'y-prosemirror';

import type { CollaborationSession } from './types';

/**
 * If the shared fragment is empty, copy a local ProseMirror document into it.
 * Safe to call on every peer — only the first non-empty seed wins.
 *
 * @returns true if this peer seeded the fragment
 */
export function seedFragmentFromProseDoc(
  session: CollaborationSession,
  proseDoc: PMNode
): boolean {
  if (session.fragment.length > 0) return false;
  prosemirrorToYXmlFragment(proseDoc, session.fragment);
  return true;
}

/**
 * Replace the shared fragment with a local ProseMirror document.
 * Use on File > Open / New so the room switches to the newly loaded DOCX
 * instead of keeping the previous (often empty) Yjs content.
 */
export function replaceFragmentWithProseDoc(
  session: CollaborationSession,
  proseDoc: PMNode
): void {
  session.doc.transact(() => {
    const frag = session.fragment;
    if (frag.length > 0) {
      frag.delete(0, frag.length);
    }
    prosemirrorToYXmlFragment(proseDoc, frag);
  });
}

/**
 * Build the initial ProseMirror doc (+ mapping) from the Yjs fragment.
 * Call after seeding (or when joining a room that already has content).
 */
export function proseDocFromSession(
  session: CollaborationSession,
  schema: Schema
): { doc: PMNode; mapping: unknown } {
  return initProseMirrorDoc(session.fragment, schema);
}

/** Whether the shared fragment already has content from another peer */
export function isFragmentSeeded(session: CollaborationSession): boolean {
  return session.fragment.length > 0;
}
