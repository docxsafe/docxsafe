/**
 * Suggestion Mode Plugin — records edits as tracked changes (w:ins / w:del)
 *
 * While enabled, every document-changing transaction is post-processed:
 * - Inserted content gets an `insertion` mark (author + date)
 * - Deleted content is restored and marked with a `deletion` mark instead,
 *   unless it was the same author's pending insertion, which really goes away
 *
 * Undo/redo and transactions tagged with SUGGESTION_SKIP_META are left alone,
 * so accept/reject commands and history work normally.
 *
 * Limitations: formatting changes (w:rPrChange) and paragraph-mark deletions
 * are applied directly rather than tracked.
 */

import { Plugin, PluginKey, TextSelection, type Transaction } from 'prosemirror-state';
import { Mapping, ReplaceStep } from 'prosemirror-transform';
import type { Mark, Node as PMNode, Slice } from 'prosemirror-model';

export const suggestionModeKey = new PluginKey('suggestionMode');

/** Set this meta on a transaction to bypass suggestion tracking */
export const SUGGESTION_SKIP_META = 'suggestionModeSkip';

export interface SuggestionModeOptions {
  /** Whether edits should currently be tracked */
  isEnabled: () => boolean;
  /** Author name stamped on revisions */
  getAuthor: () => string;
}

interface PendingDeletion {
  /** Position in the final document where the content was removed */
  pos: number;
  slice: Slice;
  /** Collapsed cursor sat at the end of the removed range (Backspace) */
  backspace: boolean;
}

let revisionCounter = 0;

/** Next revision ID, greater than any revision ID already in the document */
export function nextRevisionId(doc?: PMNode): number {
  doc?.descendants((node) => {
    for (const mark of node.marks) {
      if (mark.type.name === 'insertion' || mark.type.name === 'deletion') {
        revisionCounter = Math.max(revisionCounter, mark.attrs.revisionId as number);
      }
    }
  });
  return ++revisionCounter;
}

/** Revision timestamp in the format Word writes (no milliseconds) */
export function revisionDate(): string {
  return new Date().toISOString().replace(/\.\d{3}Z$/, 'Z');
}

function isHistoryTransaction(tr: Transaction): boolean {
  return tr.getMeta('history$') != null || tr.getMeta('addToHistory') === false;
}

export function createSuggestionModePlugin(options: SuggestionModeOptions): Plugin {
  return new Plugin({
    key: suggestionModeKey,

    appendTransaction(transactions, oldState, newState) {
      if (!options.isEnabled()) return null;
      if (
        !transactions.some((t) => t.docChanged) ||
        transactions.some((t) => t.getMeta(SUGGESTION_SKIP_META) || isHistoryTransaction(t))
      ) {
        return null;
      }

      const { insertion, deletion } = newState.schema.marks;
      if (!insertion || !deletion) return null;

      const allMaps = transactions.flatMap((t) => t.mapping.maps);
      const insertedRanges: { from: number; to: number }[] = [];
      const deletions: PendingDeletion[] = [];
      const oldSelection = oldState.selection;

      let globalIndex = 0;
      for (const t of transactions) {
        for (let i = 0; i < t.steps.length; i++, globalIndex++) {
          const step = t.steps[i];
          if (!(step instanceof ReplaceStep)) continue;
          const { from, to, slice } = step as ReplaceStep & {
            from: number;
            to: number;
            slice: Slice;
          };

          // Maps positions from just after this step into the final document
          const after = new Mapping(allMaps.slice(globalIndex + 1));

          if (slice.size > 0) {
            insertedRanges.push({
              from: after.map(from, 1),
              to: after.map(from + slice.size, -1),
            });
          }
          if (to > from) {
            const docBefore = t.docs[i];
            deletions.push({
              pos: after.map(from, -1),
              slice: docBefore.slice(from, to),
              backspace:
                globalIndex === 0 && oldSelection.empty && oldSelection.from === to,
            });
          }
        }
      }

      if (insertedRanges.length === 0 && deletions.length === 0) return null;

      const author = options.getAuthor() || 'Unknown';
      const date = revisionDate();
      const tr = newState.tr;
      tr.setMeta(SUGGESTION_SKIP_META, true);

      // 1. Mark inserted content (continuing an adjacent revision by the same author)
      for (const range of insertedRanges) {
        if (range.to <= range.from) continue;
        const revisionId = adjacentRevisionId(tr.doc, range.from, 'insertion', author);
        tr.removeMark(range.from, range.to, deletion);
        tr.addMark(range.from, range.to, insertion.create({ revisionId, author, date }));
      }

      // 2. Restore deleted content as deletion-marked text (right-to-left)
      const pureDeletion = insertedRanges.length === 0;
      let cursorTarget: number | null = null;
      for (const del of [...deletions].sort((a, b) => b.pos - a.pos)) {
        if (!hasTrackableContent(del.slice, insertion.name, author)) continue;

        const pos = tr.mapping.map(del.pos, -1);
        const sizeBefore = tr.doc.content.size;
        tr.replace(pos, pos, del.slice);
        const end = pos + (tr.doc.content.size - sizeBefore);

        // The author's own pending insertions are simply removed
        const mapsBeforeCleanup = tr.mapping.maps.length;
        removeOwnInsertions(tr, pos, end, insertion.name, author);
        const restoredEnd = tr.mapping.slice(mapsBeforeCleanup).map(end, -1);

        const revisionId = adjacentRevisionId(tr.doc, pos, 'deletion', author);
        addMarkToUnmarked(tr, pos, restoredEnd, deletion.create({ revisionId, author, date }));

        // Backspace leaves the cursor before the struck text; Delete moves past it
        if (pureDeletion && oldSelection.empty) {
          cursorTarget = del.backspace ? pos : restoredEnd;
        }
      }

      if (cursorTarget != null) {
        const clamped = Math.max(0, Math.min(cursorTarget, tr.doc.content.size));
        tr.setSelection(TextSelection.near(tr.doc.resolve(clamped)));
      }

      return tr.docChanged || tr.selectionSet ? tr : null;
    },
  });
}

/** Reuse the revision ID of an adjacent same-author change so runs merge in Word */
function adjacentRevisionId(doc: PMNode, pos: number, markName: string, author: string): number {
  const $pos = doc.resolve(pos);
  for (const node of [$pos.nodeBefore, $pos.nodeAfter]) {
    const mark = node?.marks.find((m) => m.type.name === markName);
    if (mark && mark.attrs.author === author) return mark.attrs.revisionId as number;
  }
  return nextRevisionId(doc);
}

/** Whether a deleted slice has inline content other than the author's own insertions */
function hasTrackableContent(slice: Slice, insertionName: string, author: string): boolean {
  let trackable = false;
  slice.content.descendants((node) => {
    if (trackable) return false;
    if (node.isInline) {
      const ins = node.marks.find((m) => m.type.name === insertionName);
      if (!ins || ins.attrs.author !== author) trackable = true;
      return false;
    }
    return true;
  });
  return trackable;
}

function removeOwnInsertions(
  tr: Transaction,
  from: number,
  to: number,
  markName: string,
  author: string
): void {
  const ranges: { from: number; to: number }[] = [];
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true;
    const mark = node.marks.find((m) => m.type.name === markName);
    if (mark && mark.attrs.author === author) {
      ranges.push({ from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) });
    }
    return false;
  });
  for (const range of ranges.reverse()) {
    tr.delete(range.from, range.to);
  }
}

function addMarkToUnmarked(tr: Transaction, from: number, to: number, mark: Mark): void {
  const ranges: { from: number; to: number }[] = [];
  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isInline) return true;
    if (!node.marks.some((m) => m.type === mark.type)) {
      ranges.push({ from: Math.max(pos, from), to: Math.min(pos + node.nodeSize, to) });
    }
    return false;
  });
  for (const range of ranges) {
    tr.addMark(range.from, range.to, mark);
  }
}
