/**
 * RibbonDialog — small Word-style modal (title bar, body, OK/Cancel row).
 * Traps focus, closes on Escape and backdrop click, restores focus on close.
 */

import { useEffect, useId, useRef, type FormEvent, type ReactNode } from 'react';
import { RibbonIcon } from './RibbonIcons';

export interface RibbonDialogProps {
  title: string;
  onClose: () => void;
  /** When set, renders OK/Cancel and submits on Enter */
  onSubmit?: () => void;
  submitLabel?: string;
  children: ReactNode;
  testId?: string;
}

export function RibbonDialog({
  title,
  onClose,
  onSubmit,
  submitLabel = 'OK',
  children,
  testId,
}: RibbonDialogProps) {
  const titleId = useId();
  const panelRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const previous = document.activeElement as HTMLElement | null;
    const first = panelRef.current?.querySelector<HTMLElement>(
      'input, select, textarea, button:not(.ep-dialog__close)'
    );
    first?.focus();
    return () => previous?.focus?.();
  }, []);

  const onKeyDown = (e: React.KeyboardEvent) => {
    if (e.key === 'Escape') {
      e.stopPropagation();
      onClose();
      return;
    }
    if (e.key !== 'Tab') return;
    const focusable = Array.from(
      panelRef.current?.querySelectorAll<HTMLElement>(
        'button:not(:disabled), input:not(:disabled), select, textarea, [tabindex="0"]'
      ) ?? []
    );
    if (focusable.length === 0) return;
    const first = focusable[0];
    const last = focusable[focusable.length - 1];
    if (e.shiftKey && document.activeElement === first) {
      e.preventDefault();
      last.focus();
    } else if (!e.shiftKey && document.activeElement === last) {
      e.preventDefault();
      first.focus();
    }
  };

  const submit = (e: FormEvent) => {
    e.preventDefault();
    onSubmit?.();
  };

  return (
    <div
      className="ep-dialog-backdrop"
      onMouseDown={(e) => e.target === e.currentTarget && onClose()}
    >
      <div
        ref={panelRef}
        role="dialog"
        aria-modal="true"
        aria-labelledby={titleId}
        className="ep-dialog"
        data-testid={testId}
        onKeyDown={onKeyDown}
      >
        <div className="ep-dialog__title">
          <h2 id={titleId}>{title}</h2>
          <button type="button" className="ep-dialog__close" aria-label="Close" onClick={onClose}>
            <RibbonIcon name="close" size={16} />
          </button>
        </div>
        <form onSubmit={submit}>
          <div className="ep-dialog__body">{children}</div>
          <div className="ep-dialog__footer">
            {onSubmit ? (
              <>
                <button type="submit" className="ep-dialog__btn is-primary">
                  {submitLabel}
                </button>
                <button type="button" className="ep-dialog__btn" onClick={onClose}>
                  Cancel
                </button>
              </>
            ) : (
              <button type="button" className="ep-dialog__btn is-primary" onClick={onClose}>
                Close
              </button>
            )}
          </div>
        </form>
      </div>
    </div>
  );
}
