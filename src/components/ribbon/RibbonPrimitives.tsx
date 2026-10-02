/**
 * Ribbon building blocks — groups, large/small buttons, dropdown menus and
 * check boxes laid out the way Word's ribbon is. Styling lives in editor.css
 * under `.ep-ribbon`.
 */

import {
  useCallback,
  useEffect,
  useId,
  useRef,
  useState,
  type KeyboardEvent as ReactKeyboardEvent,
  type ReactNode,
} from 'react';
import { RibbonIcon, type RibbonIconName } from './RibbonIcons';

/** Keep the editor selection when pressing ribbon controls */
const keepSelection = (e: { preventDefault: () => void }) => e.preventDefault();

// ============================================================================
// GROUP
// ============================================================================

export function RibbonGroup({
  label,
  children,
  className,
}: {
  label: string;
  children: ReactNode;
  className?: string;
}) {
  return (
    <div className={`ep-ribbon-group ${className ?? ''}`} role="group" aria-label={label}>
      <div className="ep-ribbon-group__body">{children}</div>
      <div className="ep-ribbon-group__label" aria-hidden="true">
        {label}
      </div>
    </div>
  );
}

/** Vertical stack of small controls inside a group */
export function RibbonStack({ children }: { children: ReactNode }) {
  return <div className="ep-ribbon-stack">{children}</div>;
}

/** Horizontal row of small controls inside a stack */
export function RibbonRow({ children }: { children: ReactNode }) {
  return <div className="ep-ribbon-row">{children}</div>;
}

// ============================================================================
// BUTTONS
// ============================================================================

export interface RibbonButtonProps {
  icon: RibbonIconName;
  label: string;
  /** Tooltip; defaults to label */
  title?: string;
  onClick?: () => void;
  disabled?: boolean;
  /** Toggle state (aria-pressed) */
  checked?: boolean;
  size?: 'large' | 'small';
  /** Small buttons: show the text label next to the icon */
  showLabel?: boolean;
  testId?: string;
}

export function RibbonButton({
  icon,
  label,
  title,
  onClick,
  disabled,
  checked,
  size = 'small',
  showLabel = size === 'large',
  testId,
}: RibbonButtonProps) {
  return (
    <button
      type="button"
      className={`ep-ribbon-btn ep-ribbon-btn--${size}`}
      title={title ?? label}
      aria-label={label}
      aria-pressed={checked}
      data-checked={checked || undefined}
      disabled={disabled}
      data-testid={testId}
      onMouseDown={keepSelection}
      onClick={onClick}
    >
      <RibbonIcon name={icon} size={size === 'large' ? 32 : 20} />
      {showLabel && <span className="ep-ribbon-btn__label">{label}</span>}
    </button>
  );
}

// ============================================================================
// MENUS
// ============================================================================

export interface RibbonMenuItem {
  label: string;
  description?: string;
  icon?: RibbonIconName;
  onSelect: () => void;
  disabled?: boolean;
  checked?: boolean;
  testId?: string;
}

export type RibbonMenuEntry = RibbonMenuItem | 'separator';

function useDismiss(open: boolean, close: () => void, ref: React.RefObject<HTMLElement | null>) {
  useEffect(() => {
    if (!open) return;
    const onPointer = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) close();
    };
    document.addEventListener('mousedown', onPointer);
    return () => document.removeEventListener('mousedown', onPointer);
  }, [open, close, ref]);
}

function MenuList({
  id,
  items,
  onClose,
  align = 'left',
  labelledBy,
}: {
  id: string;
  items: RibbonMenuEntry[];
  onClose: (refocus: boolean) => void;
  align?: 'left' | 'right';
  labelledBy: string;
}) {
  const listRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    listRef.current?.querySelector<HTMLButtonElement>('[role^="menuitem"]:not(:disabled)')?.focus();
  }, []);

  const onKeyDown = (e: ReactKeyboardEvent) => {
    const enabled = Array.from(
      listRef.current?.querySelectorAll<HTMLButtonElement>('[role^="menuitem"]:not(:disabled)') ??
        []
    );
    const index = enabled.indexOf(document.activeElement as HTMLButtonElement);
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      const step = e.key === 'ArrowDown' ? 1 : -1;
      enabled[(index + step + enabled.length) % enabled.length]?.focus();
    } else if (e.key === 'Home' || e.key === 'End') {
      e.preventDefault();
      enabled[e.key === 'Home' ? 0 : enabled.length - 1]?.focus();
    } else if (e.key === 'Escape') {
      e.preventDefault();
      onClose(true);
    } else if (e.key === 'Tab') {
      onClose(false);
    }
  };

  return (
    <div
      ref={listRef}
      id={id}
      role="menu"
      aria-labelledby={labelledBy}
      className={`ep-ribbon-menu ep-ribbon-menu--${align}`}
      onKeyDown={onKeyDown}
    >
      {items.map((item, i) =>
        item === 'separator' ? (
          <div key={`sep-${i}`} role="separator" className="ep-ribbon-menu__sep" />
        ) : (
          <button
            key={item.label}
            type="button"
            role={item.checked !== undefined ? 'menuitemradio' : 'menuitem'}
            aria-checked={item.checked}
            disabled={item.disabled}
            data-testid={item.testId}
            className="ep-ribbon-menu__item"
            onMouseDown={keepSelection}
            onClick={() => {
              onClose(false);
              item.onSelect();
            }}
          >
            <span className="ep-ribbon-menu__icon">
              {item.icon ? (
                <RibbonIcon name={item.icon} size={item.description ? 24 : 18} />
              ) : item.checked ? (
                <RibbonIcon name="check" size={16} />
              ) : null}
            </span>
            <span className="ep-ribbon-menu__text">
              <span className="ep-ribbon-menu__label">{item.label}</span>
              {item.description && <span className="ep-ribbon-menu__desc">{item.description}</span>}
            </span>
            {item.icon && item.checked && (
              <RibbonIcon name="check" size={16} className="ep-ribbon-menu__check" />
            )}
          </button>
        )
      )}
    </div>
  );
}

export interface RibbonMenuButtonProps {
  icon?: RibbonIconName;
  label: string;
  title?: string;
  items: RibbonMenuEntry[];
  size?: 'large' | 'small';
  showLabel?: boolean;
  disabled?: boolean;
  /** Split button: clicking the main part runs this, the arrow opens the menu */
  onPrimary?: () => void;
  align?: 'left' | 'right';
  testId?: string;
  /** Custom trigger content (replaces icon + label) */
  children?: ReactNode;
  className?: string;
}

export function RibbonMenuButton({
  icon,
  label,
  title,
  items,
  size = 'small',
  showLabel = size === 'large',
  disabled,
  onPrimary,
  align = 'left',
  testId,
  children,
  className,
}: RibbonMenuButtonProps) {
  const [open, setOpen] = useState(false);
  const rootRef = useRef<HTMLDivElement>(null);
  const triggerRef = useRef<HTMLButtonElement>(null);
  const menuId = useId();
  const triggerId = useId();

  const close = useCallback((refocus?: boolean) => {
    setOpen(false);
    if (refocus) triggerRef.current?.focus();
  }, []);
  useDismiss(open, () => close(false), rootRef);

  const content = children ?? (
    <>
      {icon && <RibbonIcon name={icon} size={size === 'large' ? 32 : 20} />}
      {showLabel && <span className="ep-ribbon-btn__label">{label}</span>}
    </>
  );

  const triggerProps = {
    'aria-haspopup': 'menu' as const,
    'aria-expanded': open,
    'aria-controls': open ? menuId : undefined,
    onMouseDown: keepSelection,
    onClick: () => setOpen((v) => !v),
    onKeyDown: (e: ReactKeyboardEvent) => {
      if (e.key === 'ArrowDown') {
        e.preventDefault();
        setOpen(true);
      }
    },
  };

  return (
    <div
      ref={rootRef}
      className={`ep-ribbon-menu-root ep-ribbon-menu-root--${size} ${onPrimary ? 'is-split' : ''} ${className ?? ''}`}
    >
      {onPrimary ? (
        <div className={`ep-ribbon-split ep-ribbon-split--${size}`} data-open={open || undefined}>
          <button
            type="button"
            className={`ep-ribbon-btn ep-ribbon-btn--${size} ep-ribbon-split__main`}
            title={title ?? label}
            aria-label={label}
            disabled={disabled}
            data-testid={testId}
            onMouseDown={keepSelection}
            onClick={onPrimary}
          >
            {content}
          </button>
          <button
            ref={triggerRef}
            id={triggerId}
            type="button"
            className="ep-ribbon-split__arrow"
            aria-label={`${label} options`}
            title={`${label} options`}
            disabled={disabled}
            data-testid={testId ? `${testId}-menu` : undefined}
            {...triggerProps}
          >
            <RibbonIcon name="chevronDown" size={12} />
          </button>
        </div>
      ) : (
        <button
          ref={triggerRef}
          id={triggerId}
          type="button"
          className={`ep-ribbon-btn ep-ribbon-btn--${size} ep-ribbon-btn--menu`}
          title={title ?? label}
          aria-label={label}
          disabled={disabled}
          data-testid={testId}
          data-open={open || undefined}
          {...triggerProps}
        >
          {content}
          <RibbonIcon name="chevronDown" size={12} className="ep-ribbon-btn__caret" />
        </button>
      )}
      {open && (
        <MenuList id={menuId} items={items} onClose={close} align={align} labelledBy={triggerId} />
      )}
    </div>
  );
}

// ============================================================================
// CHECKBOX
// ============================================================================

export function RibbonCheckbox({
  label,
  checked,
  onChange,
  disabled,
  testId,
}: {
  label: string;
  checked: boolean;
  onChange: (checked: boolean) => void;
  disabled?: boolean;
  testId?: string;
}) {
  return (
    <label className="ep-ribbon-check" data-disabled={disabled || undefined}>
      <input
        type="checkbox"
        checked={checked}
        disabled={disabled}
        data-testid={testId}
        onChange={(e) => onChange(e.target.checked)}
      />
      <span>{label}</span>
    </label>
  );
}
