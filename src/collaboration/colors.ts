/**
 * Word-like collaboration colors for remote carets / presence.
 */

const PALETTE = [
  '#185abd', // Word blue
  '#c43e1c', // terracotta
  '#0e7a3d', // green
  '#7b2d8e', // purple
  '#b76e00', // amber
  '#0078d4', // accent blue
  '#c2185b', // pink
  '#00838f', // teal
];

/** Pick a stable color from a user id / name */
export function colorForUser(seed: string): string {
  let hash = 0;
  for (let i = 0; i < seed.length; i++) {
    hash = (hash * 31 + seed.charCodeAt(i)) >>> 0;
  }
  return PALETTE[hash % PALETTE.length];
}

/** Generate a short random id */
export function randomUserId(): string {
  if (typeof crypto !== 'undefined' && 'randomUUID' in crypto) {
    return crypto.randomUUID().slice(0, 8);
  }
  return `u-${Math.random().toString(36).slice(2, 10)}`;
}
