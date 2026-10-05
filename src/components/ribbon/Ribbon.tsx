/**
 * Ribbon — Word-style tabbed command surface.
 *
 * Tabs: File (backstage), Home, Insert, Layout, Review, View, Developer, plus
 * contextual Table Layout / Picture Format tabs. The Home tab is the existing
 * Toolbar in its `ribbon` variant and stays mounted so its shortcuts keep
 * working while other tabs are shown. The right end of the tab row holds the
 * Comments button and the Editing / Reviewing / Viewing mode menu.
 */

import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import type { SectionProperties } from '../../types/document';
import type { TableAction } from '../ui/TableToolbar';
import { TableGridPicker } from '../ui/TableGridPicker';
import { TableOptionsDropdown } from '../ui/TableOptionsDropdown';
import {
  RibbonButton,
  RibbonCheckbox,
  RibbonGroup,
  RibbonMenuButton,
  RibbonStack,
  type RibbonMenuEntry,
} from './RibbonPrimitives';
import { RibbonIcon, type RibbonIconName } from './RibbonIcons';
import type { EditorMode, MarkupView } from '../review/types';

// ============================================================================
// TYPES
// ============================================================================

export type RibbonLayout = 'simplified' | 'classic';

export type RibbonTabId =
  'home' | 'insert' | 'layout' | 'review' | 'view' | 'developer' | 'tableLayout' | 'pictureFormat';

export interface PageSetupChange {
  pageWidth?: number;
  pageHeight?: number;
  orientation?: 'portrait' | 'landscape';
  marginTop?: number;
  marginBottom?: number;
  marginLeft?: number;
  marginRight?: number;
}

export interface RibbonProps {
  /** Rendered Home tab content (Toolbar variant="ribbon") */
  home: ReactNode;
  documentTitle: string;

  mode: EditorMode;
  availableModes: EditorMode[];
  onModeChange: (mode: EditorMode) => void;

  // File
  /** File > New: start a blank document */
  onNew: () => void;
  /** File > Open: pick a .docx */
  onOpen: () => void;
  /** File > Save (Ctrl+S) */
  onSave: () => void;
  /** Shown on File > Info and as "•" next to the name */
  hasUnsavedChanges: boolean;
  onDownload: () => void;
  onPrint: (withMarkup: boolean) => void;
  stats: { words: number; characters: number; paragraphs: number; pages: number };

  // Insert
  onInsertTable?: (rows: number, columns: number) => void;
  onInsertImage?: () => void;
  onInsertLink: () => void;
  onInsertPageBreak: () => void;
  hasHeader: boolean;
  hasFooter: boolean;
  onHeaderFooter: (
    kind: 'header' | 'footer',
    action: 'blank' | 'pageNumber' | 'pageXofY' | 'edit' | 'remove'
  ) => void;

  /** 'simplified' = single thin row (default), 'classic' = full ribbon with group labels */
  layout: RibbonLayout;
  onLayoutChange: (layout: RibbonLayout) => void;

  // Layout
  sectionProperties?: SectionProperties | null;
  onPageSetup: (change: PageSetupChange) => void;

  // Review
  canComment: boolean;
  onNewComment: () => void;
  onDeleteComment: () => void;
  onDeleteAllComments: () => void;
  onPreviousComment: () => void;
  onNextComment: () => void;
  commentCount: number;
  showComments: boolean;
  onShowCommentsChange: (show: boolean) => void;
  markupView: MarkupView;
  onMarkupViewChange: (view: MarkupView) => void;
  reviewingPaneOpen: boolean;
  onReviewingPaneChange: (open: boolean) => void;
  onAccept: (moveNext: boolean) => void;
  onAcceptAll: () => void;
  onReject: (moveNext: boolean) => void;
  onRejectAll: () => void;
  onPreviousChange: () => void;
  onNextChange: () => void;
  revisionCount: number;
  onWordCount: () => void;

  // View
  showRuler: boolean;
  onShowRulerChange: (show: boolean) => void;
  zoom: number;
  onZoomChange: (zoom: number) => void;
  onOnePage: () => void;
  onPageWidth: () => void;

  // Developer (Word tags)
  onInsertContentControl: () => void;
  canEditControlProperties: boolean;
  onContentControlProperties: () => void;
  onShowTagsPane: () => void;

  // Contextual
  tableContext?: {
    isInTable: boolean;
    rowCount?: number;
    columnCount?: number;
    canSplitCell?: boolean;
    hasMultiCellSelection?: boolean;
  } | null;
  onTableAction?: (action: TableAction) => void;
  imageContext?: { wrapType: string; displayMode: string; cssFloat: string | null } | null;
  onImageWrapType?: (wrapType: string) => void;
  onImageTransform?: (action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV') => void;
  onOpenImagePosition?: () => void;
  onOpenImageProperties?: () => void;
}

// ============================================================================
// CONSTANTS
// ============================================================================

const MODE_INFO: Record<EditorMode, { label: string; description: string; icon: RibbonIconName }> =
  {
    editing: { label: 'Editing', description: 'Edit document directly', icon: 'editing' },
    suggesting: {
      label: 'Reviewing',
      description: 'Edits become suggestions (Track Changes)',
      icon: 'reviewing',
    },
    viewing: { label: 'Viewing', description: 'Read or print the final document', icon: 'viewing' },
  };

const MARKUP_LABELS: Record<MarkupView, string> = {
  simple: 'Simple Markup',
  all: 'All Markup',
  none: 'No Markup',
  original: 'Original',
};

/** 1 inch = 1440 twips */
const IN = 1440;
const MM = 1440 / 25.4;

const MARGIN_PRESETS: {
  label: string;
  detail: string;
  top: number;
  bottom: number;
  left: number;
  right: number;
}[] = [
  {
    label: 'Normal',
    detail: 'Top 1" · Bottom 1" · Left 1" · Right 1"',
    top: IN,
    bottom: IN,
    left: IN,
    right: IN,
  },
  {
    label: 'Narrow',
    detail: 'Top 0.5" · Bottom 0.5" · Left 0.5" · Right 0.5"',
    top: IN / 2,
    bottom: IN / 2,
    left: IN / 2,
    right: IN / 2,
  },
  {
    label: 'Moderate',
    detail: 'Top 1" · Bottom 1" · Left 0.75" · Right 0.75"',
    top: IN,
    bottom: IN,
    left: 0.75 * IN,
    right: 0.75 * IN,
  },
  {
    label: 'Wide',
    detail: 'Top 1" · Bottom 1" · Left 2" · Right 2"',
    top: IN,
    bottom: IN,
    left: 2 * IN,
    right: 2 * IN,
  },
];

const PAGE_SIZES: { label: string; detail: string; width: number; height: number }[] = [
  { label: 'Letter', detail: '8.5" × 11"', width: 8.5 * IN, height: 11 * IN },
  { label: 'Legal', detail: '8.5" × 14"', width: 8.5 * IN, height: 14 * IN },
  { label: 'Executive', detail: '7.25" × 10.5"', width: 7.25 * IN, height: 10.5 * IN },
  {
    label: 'A4',
    detail: '21 cm × 29.7 cm',
    width: Math.round(210 * MM),
    height: Math.round(297 * MM),
  },
  {
    label: 'A5',
    detail: '14.8 cm × 21 cm',
    width: Math.round(148 * MM),
    height: Math.round(210 * MM),
  },
];

const ZOOM_LEVELS = [0.5, 0.75, 1, 1.25, 1.5, 2];

const near = (a: number | undefined, b: number) => a != null && Math.abs(a - b) < 20;

// ============================================================================
// BACKSTAGE (File tab)
// ============================================================================

type BackstagePage = 'new' | 'open' | 'info' | 'print' | 'export';

function Backstage({ props, onClose }: { props: RibbonProps; onClose: () => void }) {
  const [page, setPage] = useState<BackstagePage>('info');
  const [printMarkup, setPrintMarkup] = useState(props.markupView === 'all');
  const backRef = useRef<HTMLButtonElement>(null);

  useEffect(() => {
    backRef.current?.focus();
    const onKey = (e: KeyboardEvent) => e.key === 'Escape' && onClose();
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [onClose]);

  const hasMarkup = props.revisionCount > 0 || props.commentCount > 0;

  const navItem = (id: BackstagePage, label: string) => (
    <button
      type="button"
      className="ep-backstage__nav-item"
      aria-current={page === id ? 'page' : undefined}
      data-testid={`backstage-${id}`}
      onClick={() => setPage(id)}
    >
      {label}
    </button>
  );

  return (
    <div
      className="ep-backstage"
      role="dialog"
      aria-modal="true"
      aria-label="File"
      data-testid="backstage"
    >
      <nav className="ep-backstage__nav" aria-label="File">
        <button
          ref={backRef}
          type="button"
          className="ep-backstage__back"
          onClick={onClose}
          aria-label="Back to document"
        >
          <svg
            width="20"
            height="20"
            viewBox="0 0 24 24"
            fill="none"
            stroke="currentColor"
            strokeWidth="1.75"
            strokeLinecap="round"
            strokeLinejoin="round"
            aria-hidden="true"
          >
            <circle cx="12" cy="12" r="9" />
            <path d="M13 8l-4 4 4 4M9 12h7" />
          </svg>
        </button>
        {navItem('new', 'New')}
        {navItem('open', 'Open')}
        {navItem('info', 'Info')}
        <button
          type="button"
          className="ep-backstage__nav-item"
          data-testid="backstage-save"
          title="Save (Ctrl+S)"
          onClick={() => {
            props.onSave();
            onClose();
          }}
        >
          Save
        </button>
        <button
          type="button"
          className="ep-backstage__nav-item"
          data-testid="backstage-download"
          onClick={() => {
            props.onDownload();
            onClose();
          }}
        >
          Download a Copy
        </button>
        {navItem('print', 'Print')}
        {navItem('export', 'Export')}
      </nav>

      <div className="ep-backstage__content">
        {page === 'new' && (
          <section>
            <h1 className="ep-backstage__title">New</h1>
            <button
              type="button"
              className="ep-backstage__tile"
              data-testid="backstage-new-blank"
              onClick={() => {
                onClose();
                props.onNew();
              }}
            >
              <span className="ep-backstage__tile-page" aria-hidden="true" />
              <span>Blank document</span>
            </button>
          </section>
        )}

        {page === 'open' && (
          <section>
            <h1 className="ep-backstage__title">Open</h1>
            <button
              type="button"
              className="ep-backstage__card"
              data-testid="backstage-browse"
              onClick={() => {
                onClose();
                props.onOpen();
              }}
            >
              <RibbonIcon name="open" size={36} />
              <span>
                <strong>Browse</strong>
                <span>Open a Word document (.docx) from this device. Ctrl+O</span>
              </span>
            </button>
          </section>
        )}

        {page === 'info' && (
          <section>
            <h1 className="ep-backstage__title">Info</h1>
            <p className="ep-backstage__doc">
              {props.documentTitle}
              {props.hasUnsavedChanges && (
                <span className="ep-backstage__unsaved"> · Unsaved changes</span>
              )}
            </p>
            <dl className="ep-backstage__props">
              <dt>Pages</dt>
              <dd>{props.stats.pages}</dd>
              <dt>Words</dt>
              <dd>{props.stats.words.toLocaleString()}</dd>
              <dt>Characters</dt>
              <dd>{props.stats.characters.toLocaleString()}</dd>
              <dt>Paragraphs</dt>
              <dd>{props.stats.paragraphs.toLocaleString()}</dd>
              <dt>Comments</dt>
              <dd>{props.commentCount}</dd>
              <dt>Tracked changes</dt>
              <dd>{props.revisionCount}</dd>
            </dl>
            {props.revisionCount > 0 && (
              <p className="ep-backstage__note">
                This document contains tracked changes. Review them before sharing.
              </p>
            )}
          </section>
        )}

        {page === 'print' && (
          <section>
            <h1 className="ep-backstage__title">Print</h1>
            <button
              type="button"
              className="ep-backstage__print"
              data-testid="backstage-print-button"
              onClick={() => {
                onClose();
                props.onPrint(printMarkup);
              }}
            >
              <RibbonIcon name="print" size={40} />
              <span>Print</span>
            </button>
            {hasMarkup && (
              <RibbonCheckbox
                label="Print Markup (tracked changes)"
                checked={printMarkup}
                onChange={setPrintMarkup}
                testId="backstage-print-markup"
              />
            )}
            <p className="ep-backstage__note">
              To create a PDF, choose <strong>Save as PDF</strong> as the destination in the print
              dialog.
            </p>
          </section>
        )}

        {page === 'export' && (
          <section>
            <h1 className="ep-backstage__title">Export</h1>
            <button
              type="button"
              className="ep-backstage__card"
              data-testid="backstage-pdf"
              onClick={() => {
                onClose();
                props.onPrint(printMarkup);
              }}
            >
              <RibbonIcon name="pdf" size={36} />
              <span>
                <strong>Create PDF Document</strong>
                <span>
                  Preserves layout, formatting and fonts. Choose “Save as PDF” in the print dialog.
                </span>
              </span>
            </button>
            <button
              type="button"
              className="ep-backstage__card"
              onClick={() => {
                props.onDownload();
                onClose();
              }}
            >
              <RibbonIcon name="download" size={36} />
              <span>
                <strong>Download as Word Document (.docx)</strong>
                <span>Includes comments, tracked changes and tags.</span>
              </span>
            </button>
          </section>
        )}
      </div>
    </div>
  );
}

// ============================================================================
// RIBBON
// ============================================================================

export function Ribbon(props: RibbonProps) {
  const { mode, tableContext, imageContext } = props;
  const viewing = mode === 'viewing';
  const [activeTab, setActiveTab] = useState<RibbonTabId>('home');
  const [backstageOpen, setBackstageOpen] = useState(false);
  const [collapsed, setCollapsed] = useState(false);

  // Ctrl+F1 shows/hides the ribbon commands, as in Word
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key === 'F1') {
        e.preventDefault();
        setCollapsed((c) => !c);
      }
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, []);
  const tabRefs = useRef(new Map<string, HTMLButtonElement>());

  const tabs = useMemo(() => {
    const list: { id: RibbonTabId; label: string; contextual?: boolean }[] = viewing
      ? [
          { id: 'review', label: 'Review' },
          { id: 'view', label: 'View' },
        ]
      : [
          { id: 'home', label: 'Home' },
          { id: 'insert', label: 'Insert' },
          { id: 'layout', label: 'Layout' },
          { id: 'review', label: 'Review' },
          { id: 'view', label: 'View' },
          { id: 'developer', label: 'Developer' },
        ];
    if (!viewing && tableContext?.isInTable && props.onTableAction) {
      list.push({ id: 'tableLayout', label: 'Table Layout', contextual: true });
    }
    if (!viewing && imageContext && props.onImageWrapType) {
      list.push({ id: 'pictureFormat', label: 'Picture Format', contextual: true });
    }
    return list;
  }, [viewing, tableContext?.isInTable, imageContext, props.onTableAction, props.onImageWrapType]);

  // Fall back to a visible tab when the current one disappears (mode/context change)
  const selected: RibbonTabId = tabs.some((t) => t.id === activeTab)
    ? activeTab
    : viewing
      ? 'view'
      : 'home';

  const selectTab = useCallback((id: RibbonTabId) => {
    setActiveTab(id);
    setCollapsed(false);
  }, []);

  const onTabKeyDown = (e: React.KeyboardEvent, index: number) => {
    if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
    e.preventDefault();
    const next = tabs[(index + (e.key === 'ArrowRight' ? 1 : -1) + tabs.length) % tabs.length];
    selectTab(next.id);
    tabRefs.current.get(next.id)?.focus();
  };

  // ------------------------------------------------------------------------
  // Menus
  // ------------------------------------------------------------------------

  const sp = props.sectionProperties ?? {};
  const isLandscape = sp.orientation === 'landscape';

  const marginItems: RibbonMenuEntry[] = MARGIN_PRESETS.map((m) => ({
    label: m.label,
    description: m.detail,
    icon: 'margins',
    checked:
      near(sp.marginTop, m.top) &&
      near(sp.marginBottom, m.bottom) &&
      near(sp.marginLeft, m.left) &&
      near(sp.marginRight, m.right),
    testId: `margins-${m.label.toLowerCase()}`,
    onSelect: () =>
      props.onPageSetup({
        marginTop: m.top,
        marginBottom: m.bottom,
        marginLeft: m.left,
        marginRight: m.right,
      }),
  }));

  const swapToOrientation = (orientation: 'portrait' | 'landscape') => {
    const w = sp.pageWidth ?? 12240;
    const h = sp.pageHeight ?? 15840;
    const long = Math.max(w, h);
    const short = Math.min(w, h);
    props.onPageSetup({
      orientation,
      pageWidth: orientation === 'landscape' ? long : short,
      pageHeight: orientation === 'landscape' ? short : long,
    });
  };

  const orientationItems: RibbonMenuEntry[] = [
    {
      label: 'Portrait',
      icon: 'size',
      checked: !isLandscape,
      testId: 'orientation-portrait',
      onSelect: () => swapToOrientation('portrait'),
    },
    {
      label: 'Landscape',
      icon: 'orientation',
      checked: isLandscape,
      testId: 'orientation-landscape',
      onSelect: () => swapToOrientation('landscape'),
    },
  ];

  const sizeItems: RibbonMenuEntry[] = PAGE_SIZES.map((s) => {
    const w = isLandscape ? s.height : s.width;
    const h = isLandscape ? s.width : s.height;
    return {
      label: s.label,
      description: s.detail,
      icon: 'size',
      checked: near(sp.pageWidth, w) && near(sp.pageHeight, h),
      testId: `size-${s.label.toLowerCase()}`,
      onSelect: () => props.onPageSetup({ pageWidth: w, pageHeight: h }),
    };
  });

  const markupItems: RibbonMenuEntry[] = (
    ['simple', 'all', 'none', 'original'] as MarkupView[]
  ).map((v) => ({
    label: MARKUP_LABELS[v],
    checked: props.markupView === v,
    testId: `markup-${v}`,
    onSelect: () => props.onMarkupViewChange(v),
  }));

  const modeItems: RibbonMenuEntry[] = props.availableModes.map((m) => ({
    label: MODE_INFO[m].label,
    description: MODE_INFO[m].description,
    icon: MODE_INFO[m].icon,
    checked: mode === m,
    testId: `mode-${m}`,
    onSelect: () => props.onModeChange(m),
  }));

  const noRevisions = props.revisionCount === 0;
  const noComments = props.commentCount === 0;

  // ------------------------------------------------------------------------
  // Tab panels
  // ------------------------------------------------------------------------

  const insertPanel = (
    <>
      <RibbonGroup label="Pages">
        <RibbonButton
          size="large"
          icon="pageBreak"
          label="Page Break"
          title="Page Break (Ctrl+Enter)"
          onClick={props.onInsertPageBreak}
          testId="ribbon-page-break"
        />
      </RibbonGroup>
      {props.onInsertTable && (
        <RibbonGroup label="Tables">
          <div className="ep-ribbon-table-picker">
            <TableGridPicker onInsert={props.onInsertTable} tooltip="Insert Table" />
            <span className="ep-ribbon-btn__label" aria-hidden="true">
              Table
            </span>
          </div>
        </RibbonGroup>
      )}
      {props.onInsertImage && (
        <RibbonGroup label="Illustrations">
          <RibbonButton
            size="large"
            icon="picture"
            label="Pictures"
            title="Insert Pictures"
            onClick={props.onInsertImage}
            testId="ribbon-insert-picture"
          />
        </RibbonGroup>
      )}
      <RibbonGroup label="Links">
        <RibbonButton
          size="large"
          icon="link"
          label="Link"
          title="Insert Link (Ctrl+K)"
          onClick={props.onInsertLink}
          testId="toolbar-insert-link"
        />
      </RibbonGroup>
      <RibbonGroup label="Comments">
        <RibbonButton
          size="large"
          icon="newComment"
          label="Comment"
          title="Insert Comment (Ctrl+Alt+M)"
          disabled={!props.canComment}
          onClick={props.onNewComment}
          testId="ribbon-insert-comment"
        />
      </RibbonGroup>
      <RibbonGroup label="Header & Footer">
        <RibbonMenuButton
          size="large"
          icon="header"
          label="Header"
          testId="ribbon-header"
          items={[
            {
              label: 'Blank',
              description: 'An empty header at the top of every page',
              icon: 'header',
              onSelect: () => props.onHeaderFooter('header', 'blank'),
              testId: 'header-blank',
            },
            {
              label: 'Page Number',
              description: 'Centered page number',
              icon: 'header',
              onSelect: () => props.onHeaderFooter('header', 'pageNumber'),
              testId: 'header-page-number',
            },
            'separator',
            {
              label: 'Edit Header',
              onSelect: () => props.onHeaderFooter('header', 'edit'),
              testId: 'header-edit',
            },
            {
              label: 'Remove Header',
              disabled: !props.hasHeader,
              onSelect: () => props.onHeaderFooter('header', 'remove'),
              testId: 'header-remove',
            },
          ]}
        />
        <RibbonMenuButton
          size="large"
          icon="footer"
          label="Footer"
          testId="ribbon-footer"
          items={[
            {
              label: 'Blank',
              description: 'An empty footer at the bottom of every page',
              icon: 'footer',
              onSelect: () => props.onHeaderFooter('footer', 'blank'),
              testId: 'footer-blank',
            },
            {
              label: 'Page Number',
              description: 'Centered page number',
              icon: 'footer',
              onSelect: () => props.onHeaderFooter('footer', 'pageNumber'),
              testId: 'footer-page-number',
            },
            {
              label: 'Page X of Y',
              description: 'e.g. “Page 2 of 8”',
              icon: 'footer',
              onSelect: () => props.onHeaderFooter('footer', 'pageXofY'),
              testId: 'footer-page-x-of-y',
            },
            'separator',
            {
              label: 'Edit Footer',
              onSelect: () => props.onHeaderFooter('footer', 'edit'),
              testId: 'footer-edit',
            },
            {
              label: 'Remove Footer',
              disabled: !props.hasFooter,
              onSelect: () => props.onHeaderFooter('footer', 'remove'),
              testId: 'footer-remove',
            },
          ]}
        />
        <RibbonMenuButton
          size="large"
          icon="pageNumber"
          label="Page Number"
          testId="ribbon-page-number"
          items={[
            {
              label: 'Current Position',
              description: 'Insert page number at the cursor (header/footer)',
              icon: 'pageNumber',
              onSelect: () => props.onHeaderFooter('footer', 'pageNumber'),
              testId: 'page-number-current',
            },
            {
              label: 'Page X of Y at Cursor',
              description: 'e.g. “Page 1 of 1” at the cursor',
              icon: 'pageNumber',
              onSelect: () => props.onHeaderFooter('footer', 'pageXofY'),
              testId: 'page-number-x-of-y-current',
            },
            'separator',
            {
              label: 'Top of Page',
              description: 'Plain number, centered header',
              icon: 'header',
              onSelect: () => props.onHeaderFooter('header', 'pageNumber'),
              testId: 'page-number-top',
            },
            {
              label: 'Bottom of Page',
              description: 'Plain number, centered footer',
              icon: 'footer',
              onSelect: () => props.onHeaderFooter('footer', 'pageNumber'),
              testId: 'page-number-bottom',
            },
            {
              label: 'Bottom of Page — Page X of Y',
              icon: 'footer',
              onSelect: () => props.onHeaderFooter('footer', 'pageXofY'),
              testId: 'page-number-bottom-x-of-y',
            },
          ]}
        />
      </RibbonGroup>
    </>
  );

  const layoutPanel = (
    <>
      <RibbonGroup label="Page Setup">
        <RibbonMenuButton
          size="large"
          icon="margins"
          label="Margins"
          items={marginItems}
          testId="ribbon-margins"
        />
        <RibbonMenuButton
          size="large"
          icon="orientation"
          label="Orientation"
          items={orientationItems}
          testId="ribbon-orientation"
        />
        <RibbonMenuButton
          size="large"
          icon="size"
          label="Size"
          items={sizeItems}
          testId="ribbon-size"
        />
        <RibbonMenuButton
          size="large"
          icon="pageBreak"
          label="Breaks"
          items={[
            {
              label: 'Page',
              description: 'Mark the point at which one page ends and the next begins.',
              icon: 'pageBreak',
              onSelect: props.onInsertPageBreak,
            },
          ]}
          testId="ribbon-breaks"
        />
      </RibbonGroup>
    </>
  );

  const reviewPanel = (
    <>
      <RibbonGroup label="Proofing">
        <RibbonButton
          size="large"
          icon="wordCount"
          label="Word Count"
          onClick={props.onWordCount}
          testId="ribbon-word-count"
        />
      </RibbonGroup>
      <RibbonGroup label="Comments">
        <RibbonButton
          size="large"
          icon="newComment"
          label="New Comment"
          title="New Comment (Ctrl+Alt+M)"
          disabled={viewing || !props.canComment}
          onClick={props.onNewComment}
          testId="review-add-comment"
        />
        <RibbonMenuButton
          size="large"
          icon="deleteComment"
          label="Delete"
          disabled={viewing || noComments}
          onPrimary={props.onDeleteComment}
          items={[
            { label: 'Delete', icon: 'deleteComment', onSelect: props.onDeleteComment },
            {
              label: 'Delete All Comments in Document',
              onSelect: props.onDeleteAllComments,
              testId: 'delete-all-comments',
            },
          ]}
          testId="ribbon-delete-comment"
        />
        <RibbonStack>
          <RibbonButton
            icon="previous"
            label="Previous"
            title="Previous Comment"
            showLabel
            disabled={noComments}
            onClick={props.onPreviousComment}
            testId="ribbon-prev-comment"
          />
          <RibbonButton
            icon="next"
            label="Next"
            title="Next Comment"
            showLabel
            disabled={noComments}
            onClick={props.onNextComment}
            testId="ribbon-next-comment"
          />
        </RibbonStack>
        <RibbonButton
          size="large"
          icon="showComments"
          label="Show Comments"
          checked={props.showComments}
          onClick={() => props.onShowCommentsChange(!props.showComments)}
          testId="ribbon-show-comments"
        />
      </RibbonGroup>
      <RibbonGroup label="Tracking">
        <RibbonButton
          size="large"
          icon="trackChanges"
          label="Track Changes"
          title="Track Changes (Ctrl+Shift+E)"
          checked={mode === 'suggesting'}
          disabled={viewing || !props.availableModes.includes('suggesting')}
          onClick={() => props.onModeChange(mode === 'suggesting' ? 'editing' : 'suggesting')}
          testId="ribbon-track-changes"
        />
        <RibbonStack>
          <RibbonMenuButton
            label={MARKUP_LABELS[props.markupView]}
            title="Display for Review"
            items={markupItems}
            testId="ribbon-markup-view"
            className="ep-ribbon-select"
          >
            <RibbonIcon name="markup" size={18} />
            <span className="ep-ribbon-btn__label">{MARKUP_LABELS[props.markupView]}</span>
          </RibbonMenuButton>
          <RibbonButton
            icon="reviewingPane"
            label="Reviewing Pane"
            showLabel
            checked={props.reviewingPaneOpen}
            onClick={() => props.onReviewingPaneChange(!props.reviewingPaneOpen)}
            testId="review-toggle-sidebar"
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Changes">
        <RibbonMenuButton
          size="large"
          icon="accept"
          label="Accept"
          title="Accept and Move to Next"
          disabled={viewing || noRevisions}
          onPrimary={() => props.onAccept(true)}
          items={[
            { label: 'Accept and Move to Next', onSelect: () => props.onAccept(true) },
            {
              label: 'Accept This Change',
              onSelect: () => props.onAccept(false),
              testId: 'accept-this-change',
            },
            'separator',
            {
              label: 'Accept All Changes',
              onSelect: props.onAcceptAll,
              testId: 'accept-all-changes',
            },
            {
              label: 'Accept All Changes and Stop Tracking',
              onSelect: () => {
                props.onAcceptAll();
                props.onModeChange('editing');
              },
            },
          ]}
          testId="ribbon-accept"
        />
        <RibbonMenuButton
          size="large"
          icon="reject"
          label="Reject"
          title="Reject and Move to Next"
          disabled={viewing || noRevisions}
          onPrimary={() => props.onReject(true)}
          items={[
            { label: 'Reject and Move to Next', onSelect: () => props.onReject(true) },
            {
              label: 'Reject Change',
              onSelect: () => props.onReject(false),
              testId: 'reject-this-change',
            },
            'separator',
            {
              label: 'Reject All Changes',
              onSelect: props.onRejectAll,
              testId: 'reject-all-changes',
            },
            {
              label: 'Reject All Changes and Stop Tracking',
              onSelect: () => {
                props.onRejectAll();
                props.onModeChange('editing');
              },
            },
          ]}
          testId="ribbon-reject"
        />
        <RibbonStack>
          <RibbonButton
            icon="previous"
            label="Previous"
            title="Previous Change"
            showLabel
            disabled={noRevisions}
            onClick={props.onPreviousChange}
            testId="ribbon-prev-change"
          />
          <RibbonButton
            icon="next"
            label="Next"
            title="Next Change"
            showLabel
            disabled={noRevisions}
            onClick={props.onNextChange}
            testId="ribbon-next-change"
          />
        </RibbonStack>
      </RibbonGroup>
    </>
  );

  const viewPanel = (
    <>
      <RibbonGroup label="Views">
        {props.availableModes.includes('viewing') && (
          <RibbonButton
            size="large"
            icon="readMode"
            label="Read Mode"
            checked={viewing}
            onClick={() => props.onModeChange('viewing')}
            testId="ribbon-read-mode"
          />
        )}
        <RibbonButton
          size="large"
          icon="printLayout"
          label="Print Layout"
          checked={!viewing}
          disabled={props.availableModes.length === 1 && viewing}
          onClick={() =>
            viewing &&
            props.onModeChange(props.availableModes.find((m) => m !== 'viewing') ?? 'viewing')
          }
          testId="ribbon-print-layout"
        />
      </RibbonGroup>
      <RibbonGroup label="Show">
        <RibbonStack>
          <RibbonCheckbox
            label="Ruler"
            checked={props.showRuler}
            disabled={viewing}
            onChange={props.onShowRulerChange}
            testId="ribbon-ruler"
          />
          <RibbonCheckbox
            label="Comments"
            checked={props.showComments}
            onChange={props.onShowCommentsChange}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Zoom">
        <RibbonMenuButton
          size="large"
          icon="zoom"
          label="Zoom"
          items={ZOOM_LEVELS.map((z) => ({
            label: `${Math.round(z * 100)}%`,
            checked: Math.abs(props.zoom - z) < 0.001,
            onSelect: () => props.onZoomChange(z),
          }))}
          testId="ribbon-zoom"
        />
        <RibbonButton
          size="large"
          icon="printLayout"
          label="100%"
          onClick={() => props.onZoomChange(1)}
          testId="ribbon-zoom-100"
        />
        <RibbonStack>
          <RibbonButton icon="onePage" label="One Page" showLabel onClick={props.onOnePage} />
          <RibbonButton icon="pageWidth" label="Page Width" showLabel onClick={props.onPageWidth} />
        </RibbonStack>
      </RibbonGroup>
    </>
  );

  const developerPanel = (
    <>
      <RibbonGroup label="Controls">
        <RibbonStack>
          <RibbonButton
            icon="tag"
            label="Rich Text Content Control"
            showLabel
            onClick={props.onInsertContentControl}
            testId="review-insert-tag"
          />
          <RibbonButton
            icon="markup"
            label="Properties"
            title="Content Control Properties"
            showLabel
            disabled={!props.canEditControlProperties}
            onClick={props.onContentControlProperties}
            testId="ribbon-control-properties"
          />
        </RibbonStack>
        <RibbonButton
          size="large"
          icon="reviewingPane"
          label="Tags Pane"
          title="Show all Word tags in this document"
          onClick={props.onShowTagsPane}
          testId="ribbon-tags-pane"
        />
      </RibbonGroup>
    </>
  );

  const tableAction = (action: TableAction) => () => props.onTableAction?.(action);
  const tablePanel = tableContext?.isInTable && props.onTableAction && (
    <>
      <RibbonGroup label="Table">
        <RibbonMenuButton
          size="large"
          icon="table"
          label="Select"
          items={[
            { label: 'Select Row', onSelect: tableAction('selectRow') },
            { label: 'Select Column', onSelect: tableAction('selectColumn') },
            { label: 'Select Table', onSelect: tableAction('selectTable') },
          ]}
        />
      </RibbonGroup>
      <RibbonGroup label="Rows & Columns">
        <RibbonMenuButton
          size="large"
          icon="deleteTable"
          label="Delete"
          items={[
            { label: 'Delete Columns', onSelect: tableAction('deleteColumn') },
            { label: 'Delete Rows', onSelect: tableAction('deleteRow') },
            { label: 'Delete Table', onSelect: tableAction('deleteTable') },
          ]}
          testId="ribbon-table-delete"
        />
        <RibbonButton
          size="large"
          icon="insertRowAbove"
          label="Insert Above"
          onClick={tableAction('addRowAbove')}
        />
        <RibbonStack>
          <RibbonButton
            icon="insertRowBelow"
            label="Insert Below"
            showLabel
            onClick={tableAction('addRowBelow')}
          />
          <RibbonButton
            icon="insertColLeft"
            label="Insert Left"
            showLabel
            onClick={tableAction('addColumnLeft')}
          />
          <RibbonButton
            icon="insertColRight"
            label="Insert Right"
            showLabel
            onClick={tableAction('addColumnRight')}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Merge">
        <RibbonStack>
          <RibbonButton
            icon="merge"
            label="Merge Cells"
            showLabel
            disabled={!tableContext.hasMultiCellSelection}
            onClick={tableAction('mergeCells')}
          />
          <RibbonButton
            icon="split"
            label="Split Cells"
            showLabel
            disabled={!tableContext.canSplitCell}
            onClick={tableAction('splitCell')}
          />
        </RibbonStack>
      </RibbonGroup>
      <RibbonGroup label="Borders & Shading">
        <TableOptionsDropdown
          onAction={(a) => props.onTableAction?.(a)}
          tableContext={tableContext}
          tooltip="Table options"
        />
      </RibbonGroup>
    </>
  );

  const wrapItems: RibbonMenuEntry[] = imageContext
    ? [
        {
          label: 'In Line with Text',
          checked: imageContext.wrapType === 'inline',
          onSelect: () => props.onImageWrapType?.('inline'),
        },
        {
          label: 'Square (Left)',
          checked: imageContext.displayMode === 'float' && imageContext.cssFloat === 'left',
          onSelect: () => props.onImageWrapType?.('wrapRight'),
        },
        {
          label: 'Square (Right)',
          checked: imageContext.displayMode === 'float' && imageContext.cssFloat === 'right',
          onSelect: () => props.onImageWrapType?.('wrapLeft'),
        },
        {
          label: 'Top and Bottom',
          checked: imageContext.wrapType === 'topAndBottom',
          onSelect: () => props.onImageWrapType?.('topAndBottom'),
        },
        {
          label: 'Behind Text',
          checked: imageContext.wrapType === 'behind',
          onSelect: () => props.onImageWrapType?.('behind'),
        },
        {
          label: 'In Front of Text',
          checked: imageContext.wrapType === 'inFront',
          onSelect: () => props.onImageWrapType?.('inFront'),
        },
      ]
    : [];

  const picturePanel = imageContext && props.onImageWrapType && (
    <>
      <RibbonGroup label="Accessibility">
        {props.onOpenImageProperties && (
          <RibbonButton
            size="large"
            icon="picture"
            label="Alt Text"
            title="Alt text and border"
            onClick={props.onOpenImageProperties}
          />
        )}
      </RibbonGroup>
      <RibbonGroup label="Arrange">
        {props.onOpenImagePosition && (
          <RibbonButton
            size="large"
            icon="margins"
            label="Position"
            onClick={props.onOpenImagePosition}
          />
        )}
        <RibbonMenuButton size="large" icon="pageWidth" label="Wrap Text" items={wrapItems} />
        {props.onImageTransform && (
          <RibbonMenuButton
            size="large"
            icon="redo"
            label="Rotate"
            items={[
              { label: 'Rotate Right 90°', onSelect: () => props.onImageTransform?.('rotateCW') },
              { label: 'Rotate Left 90°', onSelect: () => props.onImageTransform?.('rotateCCW') },
              { label: 'Flip Vertical', onSelect: () => props.onImageTransform?.('flipV') },
              { label: 'Flip Horizontal', onSelect: () => props.onImageTransform?.('flipH') },
            ]}
          />
        )}
      </RibbonGroup>
    </>
  );

  const panels: Partial<Record<RibbonTabId, ReactNode>> = {
    insert: insertPanel,
    layout: layoutPanel,
    review: reviewPanel,
    view: viewPanel,
    developer: developerPanel,
    tableLayout: tablePanel,
    pictureFormat: picturePanel,
  };

  return (
    <div className={`ep-ribbon ep-ribbon--${props.layout}`} data-testid="ribbon">
      <div className="ep-ribbon__tabs">
        <button
          type="button"
          className="ep-ribbon__file"
          aria-haspopup="dialog"
          aria-expanded={backstageOpen}
          data-testid="ribbon-tab-file"
          onClick={() => setBackstageOpen(true)}
        >
          File
        </button>
        <div role="tablist" aria-label="Ribbon tabs" className="ep-ribbon__tablist">
          {tabs.map((tab, i) => (
            <button
              key={tab.id}
              ref={(el) => {
                if (el) tabRefs.current.set(tab.id, el);
              }}
              type="button"
              role="tab"
              id={`ep-ribbon-tab-${tab.id}`}
              aria-selected={selected === tab.id && !collapsed}
              aria-controls={`ep-ribbon-panel-${tab.id}`}
              tabIndex={selected === tab.id ? 0 : -1}
              className={`ep-ribbon__tab${tab.contextual ? ' is-contextual' : ''}`}
              data-testid={`ribbon-tab-${tab.id}`}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() =>
                selected === tab.id && collapsed ? setCollapsed(false) : selectTab(tab.id)
              }
              onDoubleClick={() => setCollapsed((c) => !c)}
              onKeyDown={(e) => onTabKeyDown(e, i)}
            >
              {tab.label}
            </button>
          ))}
        </div>

        <div className="ep-ribbon__end">
          <button
            type="button"
            className="ep-ribbon__comments"
            title={props.showComments ? 'Hide Comments' : 'Show Comments'}
            aria-pressed={props.showComments}
            aria-label={props.showComments ? 'Hide Comments' : 'Show Comments'}
            onMouseDown={(e) => e.preventDefault()}
            onClick={() => props.onShowCommentsChange(!props.showComments)}
            data-testid="ribbon-comments-button"
          >
            <RibbonIcon name="showComments" size={18} />
            <span>Comments</span>
            {props.commentCount > 0 && (
              <span className="ep-ribbon__count">{props.commentCount}</span>
            )}
          </button>
          {props.availableModes.length > 1 && (
            <RibbonMenuButton
              label={`Mode: ${MODE_INFO[mode].label}`}
              title="Change mode"
              items={modeItems}
              align="right"
              testId="mode-menu"
              className={`ep-ribbon-mode is-${mode}`}
            >
              <RibbonIcon name={MODE_INFO[mode].icon} size={18} />
              <span className="ep-ribbon-btn__label">{MODE_INFO[mode].label}</span>
            </RibbonMenuButton>
          )}
          <RibbonMenuButton
            label="Ribbon display options"
            title="Ribbon display options (Ctrl+F1 collapses)"
            align="right"
            testId="ribbon-display-options"
            className="ep-ribbon__display"
            items={[
              {
                label: 'Simplified ribbon',
                description: 'Commands on a single, thin line',
                checked: props.layout === 'simplified' && !collapsed,
                onSelect: () => {
                  props.onLayoutChange('simplified');
                  setCollapsed(false);
                },
                testId: 'ribbon-layout-simplified',
              },
              {
                label: 'Classic ribbon',
                description: 'Full ribbon with large buttons and group names',
                checked: props.layout === 'classic' && !collapsed,
                onSelect: () => {
                  props.onLayoutChange('classic');
                  setCollapsed(false);
                },
                testId: 'ribbon-layout-classic',
              },
              'separator',
              {
                label: 'Show tabs only',
                description: 'Click a tab to show its commands',
                checked: collapsed,
                onSelect: () => setCollapsed(true),
                testId: 'ribbon-layout-tabs-only',
              },
            ]}
          >
            <RibbonIcon name="chevronDown" size={16} />
          </RibbonMenuButton>
        </div>
      </div>

      {/* Home stays mounted (hidden when inactive) so its shortcuts and pickers persist */}
      {!viewing && (
        <div
          role="tabpanel"
          id="ep-ribbon-panel-home"
          aria-labelledby="ep-ribbon-tab-home"
          className="ep-ribbon__panel"
          hidden={collapsed || selected !== 'home'}
        >
          {props.home}
        </div>
      )}
      {!collapsed && selected !== 'home' && (
        <div
          role="tabpanel"
          id={`ep-ribbon-panel-${selected}`}
          aria-labelledby={`ep-ribbon-tab-${selected}`}
          className="ep-ribbon__panel"
        >
          <div
            className="ep-ribbon-panel__groups"
            role="toolbar"
            aria-label={tabs.find((t) => t.id === selected)?.label}
          >
            {panels[selected]}
          </div>
        </div>
      )}

      {backstageOpen && <Backstage props={props} onClose={() => setBackstageOpen(false)} />}
    </div>
  );
}
