/**
 * Shared review types
 */

/**
 * Editing mode. Word calls these Editing, Reviewing (Track Changes on) and Viewing.
 */
export type EditorMode = 'editing' | 'suggesting' | 'viewing';

/** Word's "Display for Review" options */
export type MarkupView = 'simple' | 'all' | 'none' | 'original';
