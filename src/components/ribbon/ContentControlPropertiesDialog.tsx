/**
 * Content Control Properties — set a Word tag's Title (w:alias), Tag (w:tag),
 * and Lock (w:lock), like Word's Developer > Properties dialog.
 */

import { useState } from 'react';
import { RibbonDialog } from './RibbonDialog';
import {
  sdtLockFlags,
  sdtLockFromFlags,
  aliasToWordTag,
  normalizeContentControlTagAttrs,
  type SdtLock,
} from '../../types/content';

export interface ContentControlPropertiesInitial {
  tag: string;
  alias: string;
  lock?: SdtLock | null;
}

export interface ContentControlPropertiesResult {
  tag: string;
  alias: string;
  lock: SdtLock | null;
}

export function ContentControlPropertiesDialog({
  initial,
  onClose,
  onSave,
}: {
  initial: ContentControlPropertiesInitial;
  onClose: () => void;
  onSave: (attrs: ContentControlPropertiesResult) => void;
}) {
  const initialFlags = sdtLockFlags(initial.lock);
  const [alias, setAlias] = useState(initial.alias);
  const [tag, setTag] = useState(initial.tag);
  const [cannotDelete, setCannotDelete] = useState(initialFlags.cannotDelete);
  const [cannotEditContents, setCannotEditContents] = useState(initialFlags.cannotEditContents);

  return (
    <RibbonDialog
      title="Content Control Properties"
      onClose={onClose}
      onSubmit={() => {
        const normalized = normalizeContentControlTagAttrs({ alias, tag });
        onSave({
          alias: normalized.alias,
          tag: normalized.tag,
          lock: sdtLockFromFlags(cannotDelete, cannotEditContents) ?? null,
        });
      }}
      testId="control-properties-dialog"
    >
      <label>
        Title <span className="ep-dialog__hint">(w:alias)</span>
        <input
          type="text"
          value={alias}
          onChange={(e) => {
            const next = e.target.value;
            setAlias(next);
            if (!tag || tag === aliasToWordTag(alias)) {
              setTag(next.trim() ? aliasToWordTag(next) : '');
            }
          }}
          placeholder="e.g. Client name"
          data-testid="tag-title-input"
        />
      </label>
      <label>
        Tag <span className="ep-dialog__hint">(w:tag)</span>
        <input
          type="text"
          value={tag}
          onChange={(e) => setTag(e.target.value)}
          placeholder="e.g. client_name"
          spellCheck={false}
          data-testid="tag-tag-input"
        />
      </label>
      <fieldset className="ep-dialog__fieldset" data-testid="tag-lock-options">
        <legend>Locking</legend>
        <label className="ep-dialog__check">
          <input
            type="checkbox"
            checked={cannotDelete}
            onChange={(e) => setCannotDelete(e.target.checked)}
            data-testid="tag-lock-cannot-delete"
          />
          Content control cannot be deleted
        </label>
        <label className="ep-dialog__check">
          <input
            type="checkbox"
            checked={cannotEditContents}
            onChange={(e) => setCannotEditContents(e.target.checked)}
            data-testid="tag-lock-cannot-edit"
          />
          Contents cannot be edited
        </label>
      </fieldset>
    </RibbonDialog>
  );
}
