/**
 * StatusBar — Word-style status bar: page position, word count, Track Changes
 * state, view buttons, and the zoom slider.
 */

import { RibbonIcon } from './RibbonIcons';
import type { EditorMode } from '../review/types';

export interface StatusBarProps {
  currentPage: number;
  totalPages: number;
  wordCount: number;
  mode: EditorMode;
  /** Toggle Track Changes; omitted when switching isn't allowed */
  onToggleTrackChanges?: () => void;
  onShowWordCount: () => void;
  onReadMode?: () => void;
  onPrintLayout?: () => void;
  zoom: number;
  onZoomChange: (zoom: number) => void;
  minZoom?: number;
  maxZoom?: number;
  /** Show the zoom controls (default: true) */
  showZoom?: boolean;
}

const ZOOM_STEP = 0.1;

export function StatusBar({
  currentPage,
  totalPages,
  wordCount,
  mode,
  onToggleTrackChanges,
  onShowWordCount,
  onReadMode,
  onPrintLayout,
  zoom,
  onZoomChange,
  minZoom = 0.5,
  maxZoom = 2,
  showZoom = true,
}: StatusBarProps) {
  const percent = Math.round(zoom * 100);
  const clamp = (z: number) => Math.min(maxZoom, Math.max(minZoom, Math.round(z * 100) / 100));

  return (
    <footer className="ep-statusbar" data-testid="status-bar" aria-label="Status bar">
      <div className="ep-statusbar__left">
        <span className="ep-statusbar__item" data-testid="status-page">
          Page {Math.max(1, currentPage)} of {Math.max(1, totalPages)}
        </span>
        <button
          type="button"
          className="ep-statusbar__item ep-statusbar__btn"
          onClick={onShowWordCount}
          title="Word Count"
          data-testid="status-words"
        >
          {wordCount.toLocaleString()} {wordCount === 1 ? 'word' : 'words'}
        </button>
        {onToggleTrackChanges && mode !== 'viewing' && (
          <button
            type="button"
            className="ep-statusbar__item ep-statusbar__btn"
            onClick={onToggleTrackChanges}
            title="Track Changes (Ctrl+Shift+E)"
            aria-pressed={mode === 'suggesting'}
            data-testid="status-track-changes"
          >
            Track Changes: {mode === 'suggesting' ? 'On' : 'Off'}
          </button>
        )}
        {mode === 'viewing' && <span className="ep-statusbar__item">Read-only</span>}
      </div>

      <div className="ep-statusbar__right">
        {onReadMode && (
          <button
            type="button"
            className="ep-statusbar__icon"
            onClick={onReadMode}
            title="Read Mode"
            aria-label="Read Mode"
            aria-pressed={mode === 'viewing'}
          >
            <RibbonIcon name="readMode" size={18} />
          </button>
        )}
        {onPrintLayout && (
          <button
            type="button"
            className="ep-statusbar__icon"
            onClick={onPrintLayout}
            title="Print Layout"
            aria-label="Print Layout"
            aria-pressed={mode !== 'viewing'}
          >
            <RibbonIcon name="printLayout" size={18} />
          </button>
        )}
        {showZoom && (
          <>
            <span className="ep-statusbar__sep" aria-hidden="true" />
            <button
              type="button"
              className="ep-statusbar__icon"
              onClick={() => onZoomChange(clamp(zoom - ZOOM_STEP))}
              aria-label="Zoom out"
              title="Zoom out"
              disabled={zoom <= minZoom}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true">
                <path d="M5 12h14" stroke="currentColor" strokeWidth="2" strokeLinecap="round" />
              </svg>
            </button>
            <input
              type="range"
              className="ep-statusbar__zoom"
              min={minZoom * 100}
              max={maxZoom * 100}
              step={10}
              value={percent}
              onChange={(e) => onZoomChange(clamp(Number(e.target.value) / 100))}
              aria-label="Zoom"
              aria-valuetext={`${percent}%`}
              data-testid="status-zoom-slider"
            />
            <button
              type="button"
              className="ep-statusbar__icon"
              onClick={() => onZoomChange(clamp(zoom + ZOOM_STEP))}
              aria-label="Zoom in"
              title="Zoom in"
              disabled={zoom >= maxZoom}
            >
              <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true">
                <path
                  d="M5 12h14M12 5v14"
                  stroke="currentColor"
                  strokeWidth="2"
                  strokeLinecap="round"
                />
              </svg>
            </button>
            <button
              type="button"
              className="ep-statusbar__item ep-statusbar__btn ep-statusbar__percent"
              onClick={() => onZoomChange(1)}
              title="Zoom to 100%"
              data-testid="status-zoom"
            >
              {percent}%
            </button>
          </>
        )}
      </div>
    </footer>
  );
}
