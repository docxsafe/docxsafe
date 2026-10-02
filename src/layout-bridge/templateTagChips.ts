/**
 * Template tag chips — split text runs at docxtemplater tag boundaries so the
 * painter can draw each `{tag}` as a pill.
 *
 * Every character keeps its own run position, so measurement, line breaking,
 * caret placement and click mapping are unchanged: the braces are still laid
 * out (drawn transparent as the chip's padding) and the tag stays editable.
 */

import type { FlowBlock, Run, TextRun } from '../layout-engine/types';

export interface TemplateTagRange {
  from: number;
  to: number;
  type: string;
}

type ChipPart = NonNullable<TextRun['templateTag']>['part'];

function partAt(pos: number, tag: TemplateTagRange): ChipPart {
  if (pos === tag.from) return 'open';
  if (pos === tag.to - 1) return 'close';
  return 'body';
}

/** Split one text run into pieces that are uniformly inside/outside tags */
function splitRun(run: TextRun, tags: TemplateTagRange[]): Run[] {
  if (run.pmStart == null || run.pmEnd == null || run.hyperlink) return [run];
  const start = run.pmStart;
  const end = run.pmEnd;
  // Only plain text runs map 1 char = 1 position
  if (end - start !== run.text.length) return [run];

  const overlapping = tags.filter((t) => t.from < end && t.to > start);
  if (overlapping.length === 0) return [run];

  const pieces: Run[] = [];
  let pieceStart = start;
  let current: { tag: TemplateTagRange; part: ChipPart } | null = null;

  const classify = (pos: number) => {
    const tag = overlapping.find((t) => pos >= t.from && pos < t.to);
    return tag ? { tag, part: partAt(pos, tag) } : null;
  };
  const same = (
    a: { tag: TemplateTagRange; part: ChipPart } | null,
    b: { tag: TemplateTagRange; part: ChipPart } | null
  ) => (a === null && b === null) || (!!a && !!b && a.tag === b.tag && a.part === b.part);

  const flush = (pieceEnd: number) => {
    if (pieceEnd <= pieceStart) return;
    const text = run.text.slice(pieceStart - start, pieceEnd - start);
    const piece: TextRun = { ...run, text, pmStart: pieceStart, pmEnd: pieceEnd };
    if (current) piece.templateTag = { type: current.tag.type, part: current.part };
    else delete piece.templateTag;
    pieces.push(piece);
    pieceStart = pieceEnd;
  };

  current = classify(start);
  for (let pos = start + 1; pos < end; pos++) {
    const next = classify(pos);
    if (!same(current, next)) {
      flush(pos);
      current = next;
    }
  }
  flush(end);
  return pieces;
}

function applyToRuns(runs: Run[], tags: TemplateTagRange[]): Run[] {
  const out: Run[] = [];
  for (const run of runs) {
    if (run.kind === 'text') out.push(...splitRun(run, tags));
    else out.push(run);
  }
  return out;
}

/**
 * Return blocks whose text runs are split at template tag boundaries.
 * Blocks without tags are returned unchanged (same object).
 */
export function applyTemplateTagChips(
  blocks: FlowBlock[],
  tags: TemplateTagRange[]
): FlowBlock[] {
  if (tags.length === 0) return blocks;
  return blocks.map((block) => {
    if (block.kind === 'paragraph') {
      return { ...block, runs: applyToRuns(block.runs, tags) };
    }
    if (block.kind === 'table') {
      return {
        ...block,
        rows: block.rows.map((row) => ({
          ...row,
          cells: row.cells.map((cell) => ({
            ...cell,
            blocks: applyTemplateTagChips(cell.blocks, tags),
          })),
        })),
      };
    }
    return block;
  });
}
