/**
 * DocxEditor Component
 *
 * Main component integrating all editor features:
 * - Toolbar for formatting
 * - ProseMirror-based editor for content editing
 * - VariablePanel for template variables
 * - Zoom control
 * - Error boundary
 * - Loading states
 */

import {
  useRef,
  useCallback,
  useState,
  useEffect,
  useMemo,
  forwardRef,
  useImperativeHandle,
} from 'react';
import type { CSSProperties, ReactNode } from 'react';
import type {
  Document,
  Theme,
  HeaderFooter,
  Paragraph,
  ParagraphContent,
  Run,
} from '../types/document';

import { Toolbar, type SelectionFormatting, type FormattingAction } from './Toolbar';
import { pointsToHalfPoints } from './ui/FontSizePicker';
import { VariablePanel } from './VariablePanel';
import { ErrorBoundary, ErrorProvider } from './ErrorBoundary';
import type { TableAction } from './ui/TableToolbar';
import { mapHexToHighlightName } from './toolbarUtils';
import {
  PageNumberIndicator,
  type PageIndicatorPosition,
  type PageIndicatorVariant,
} from './ui/PageNumberIndicator';
import {
  PageNavigator,
  type PageNavigatorPosition,
  type PageNavigatorVariant,
} from './ui/PageNavigator';
import { HorizontalRuler } from './ui/HorizontalRuler';
import { VerticalRuler } from './ui/VerticalRuler';
import { type PrintOptions } from './ui/PrintPreview';
import type { EditorMode } from './review/types';
import { ReviewSidebar } from './review/ReviewSidebar';
import { useReview } from './review/useReview';
import { CommentBalloons } from './review/CommentBalloons';
import { Ribbon, type PageSetupChange, type RibbonLayout } from './ribbon/Ribbon';
import { StatusBar } from './ribbon/StatusBar';
import { RibbonDialog } from './ribbon/RibbonDialog';
import { ContentControlPropertiesDialog } from './ribbon/ContentControlPropertiesDialog';
import { insertPageBreak } from '../prosemirror/commands/pageBreak';
import {
  FindReplaceDialog,
  useFindReplace,
  findInDocument,
  scrollToMatch,
  type FindMatch,
  type FindOptions,
  type FindResult,
} from './dialogs/FindReplaceDialog';
import { HyperlinkDialog, useHyperlinkDialog, type HyperlinkData } from './dialogs/HyperlinkDialog';
import { TablePropertiesDialog } from './dialogs/TablePropertiesDialog';
import { ImagePositionDialog, type ImagePositionData } from './dialogs/ImagePositionDialog';
import { ImagePropertiesDialog, type ImagePropertiesData } from './dialogs/ImagePropertiesDialog';
import {
  HeaderFooterEditor,
  insertPageNumberInView,
} from './HeaderFooterEditor';
import type { EditorView as PMEditorView } from 'prosemirror-view';
import { undo as pmUndo, redo as pmRedo } from 'prosemirror-history';
import type { DocxCollaborationConfig } from '../collaboration';
import { CollaborationPresence } from '../collaboration/CollaborationPresence';
import { collabUndo, collabRedo } from '../collaboration/plugins';
import {
  replaceFragmentWithProseDoc,
  seedFragmentFromProseDoc,
  isFragmentSeeded,
} from '../collaboration/seed';
import { whenRoomReady } from '../collaboration/ready';
import { toProseDoc } from '../prosemirror/conversion';
import '../collaboration/collaboration.css';
import { FootnotePropertiesDialog } from './dialogs/FootnotePropertiesDialog';
import { getBuiltinTableStyle, type TableStylePreset } from './ui/TableStyleGallery';
import { DocumentAgent } from '../agent/DocumentAgent';
import {
  DefaultLoadingIndicator,
  DefaultPlaceholder,
  ParseError,
  extractVariableNames,
  extractVariables,
} from './DocxEditorHelpers';
import { parseDocx } from '../docx/parser';
import { createEmptyDocx } from '../docx/rezip';
import { type DocxInput } from '../utils/docxInput';
import { onFontsLoaded, loadDocumentFonts } from '../utils/fontLoader';
import { executeCommand } from '../agent/executor';
import { useTableSelection } from '../hooks/useTableSelection';
import { useDocumentHistory } from '../hooks/useHistory';

// Extension system
import { createStarterKit } from '../prosemirror/extensions/StarterKit';
import { ExtensionManager } from '../prosemirror/extensions/ExtensionManager';

// ProseMirror editor
import {
  type SelectionState,
  TextSelection,
  extractSelectionState,
  toggleBold,
  toggleItalic,
  toggleUnderline,
  toggleStrike,
  toggleSuperscript,
  toggleSubscript,
  setTextColor,
  setHighlight,
  setFontSize,
  setFontFamily,
  setAlignment,
  setLineSpacing,
  toggleBulletList,
  toggleNumberedList,
  increaseIndent,
  decreaseIndent,
  increaseListLevel,
  decreaseListLevel,
  clearFormatting,
  applyStyle,
  createStyleResolver,
  // Hyperlink commands
  getHyperlinkAttrs,
  getSelectedText,
  setHyperlink,
  removeHyperlink,
  insertHyperlink,
  // Table commands
  getTableContext,
  insertTable,
  addRowAbove,
  addRowBelow,
  deleteRow as pmDeleteRow,
  addColumnLeft,
  addColumnRight,
  deleteColumn as pmDeleteColumn,
  deleteTable as pmDeleteTable,
  selectTable as pmSelectTable,
  selectRow as pmSelectRow,
  selectColumn as pmSelectColumn,
  mergeCells as pmMergeCells,
  splitCell as pmSplitCell,
  setCellBorder,
  setCellVerticalAlign,
  setCellMargins,
  setCellTextDirection,
  toggleNoWrap,
  setRowHeight,
  toggleHeaderRow,
  distributeColumns,
  autoFitContents,
  setTableProperties,
  applyTableStyle,
  removeTableBorders,
  setAllTableBorders,
  setOutsideTableBorders,
  setInsideTableBorders,
  setCellFillColor,
  setTableBorderColor,
  type TableContextInfo,
} from '../prosemirror';

// Paginated editor
import { PagedEditor, type PagedEditorRef } from '../paged-editor/PagedEditor';

// Plugin API types
import type { RenderedDomContext } from '../plugin-api/types';

// ============================================================================
// TYPES
// ============================================================================

/**
 * DocxEditor props
 */
export interface DocxEditorProps {
  /** Document data — ArrayBuffer, Uint8Array, Blob, or File */
  documentBuffer?: DocxInput | null;
  /** Pre-parsed document (alternative to documentBuffer) */
  document?: Document | null;
  /** Callback when document is saved */
  onSave?: (buffer: ArrayBuffer) => void;
  /** Callback when document changes */
  onChange?: (document: Document) => void;
  /** Callback when selection changes */
  onSelectionChange?: (state: SelectionState | null) => void;
  /** Callback on error */
  onError?: (error: Error) => void;
  /** Callback when fonts are loaded */
  onFontsLoaded?: () => void;
  /** External ProseMirror plugins (from PluginHost) */
  externalPlugins?: import('prosemirror-state').Plugin[];
  /**
   * Real-time co-editing via Yjs + y-webrtc.
   * Create a session with `useCollaboration` / `createCollaborationSession`
   * from `decidendi-editor/collaboration`, then pass it here.
   */
  collaboration?: DocxCollaborationConfig | null;
  /** Callback when editor view is ready (for PluginHost) */
  onEditorViewReady?: (view: import('prosemirror-view').EditorView) => void;
  /** Theme for styling */
  theme?: Theme | null;
  /** Whether to show toolbar (default: true) */
  showToolbar?: boolean;
  /** Whether to show variable panel (default: true) */
  showVariablePanel?: boolean;
  /** Whether to show zoom control (default: true) */
  showZoomControl?: boolean;
  /** Whether to show page number indicator (default: true) */
  showPageNumbers?: boolean;
  /** Whether to enable interactive page navigation (default: true) */
  enablePageNavigation?: boolean;
  /** Position of page number indicator (default: 'bottom-center') */
  pageNumberPosition?: PageIndicatorPosition | PageNavigatorPosition;
  /** Variant of page number indicator (default: 'default') */
  pageNumberVariant?: PageIndicatorVariant | PageNavigatorVariant;
  /** Whether to show page margin guides/boundaries (default: false) */
  showMarginGuides?: boolean;
  /** Color for margin guides (default: '#c0c0c0') */
  marginGuideColor?: string;
  /** Whether to show horizontal ruler (default: false) */
  showRuler?: boolean;
  /** Unit for ruler display (default: 'inch') */
  rulerUnit?: 'inch' | 'cm';
  /** Initial zoom level (default: 1.0) */
  initialZoom?: number;
  /** Whether the editor is read-only. When true, hides toolbar, rulers, and variable panel */
  readOnly?: boolean;
  /** Custom toolbar actions */
  toolbarExtra?: ReactNode;
  /** Variable panel position (default: 'right') */
  variablePanelPosition?: 'left' | 'right';
  /** Variable descriptions */
  variableDescriptions?: Record<string, string>;
  /** Additional CSS class name */
  className?: string;
  /** Additional inline styles */
  style?: CSSProperties;
  /** Placeholder when no document */
  placeholder?: ReactNode;
  /** Loading indicator */
  loadingIndicator?: ReactNode;
  /**
   * Editing mode: 'editing' (direct edits), 'suggesting' (edits become tracked
   * changes), or 'viewing' (read-only). Controlled when set; otherwise use defaultMode.
   */
  mode?: EditorMode;
  /** Initial mode when `mode` is uncontrolled (default: 'editing') */
  defaultMode?: EditorMode;
  /** Called when the user switches mode */
  onModeChange?: (mode: EditorMode) => void;
  /** Modes offered in the mode switcher (default: all three) */
  availableModes?: EditorMode[];
  /** Author name used for new comments and suggestions (default: 'Anonymous') */
  author?: string;
  /**
   * Ribbon layout: 'simplified' (single thin row, default) or 'classic' (full
   * ribbon with large buttons and group names). Users can switch from the ribbon.
   */
  ribbonLayout?: RibbonLayout;
  /** Draw docxtemplater `{tags}` as chips in the document (default: true) */
  templateTagChips?: boolean;
  /** Show the Word-style status bar (page, word count, Track Changes, zoom) (default: true) */
  showStatusBar?: boolean;
  /** Document name shown in File > Info and used for downloads (default: 'Document') */
  documentName?: string;
  /** Called after File > New created a blank document */
  onNew?: () => void;
  /** Called after File > Open loaded a file */
  onOpen?: (file: File) => void;
  /** Open the review sidebar on load (default: opens when the document has comments) */
  defaultReviewSidebarOpen?: boolean;
  /** @deprecated Print is always available from File > Print in the ribbon */
  showPrintButton?: boolean;
  /** Print options for print preview */
  printOptions?: PrintOptions;
  /** Callback when print is triggered */
  onPrint?: () => void;
  /** Callback when content is copied */
  onCopy?: () => void;
  /** Callback when content is cut */
  onCut?: () => void;
  /** Callback when content is pasted */
  onPaste?: () => void;
  /**
   * Callback when rendered DOM context is ready (for plugin overlays).
   * Used by PluginHost to get access to the rendered page DOM for positioning.
   */
  onRenderedDomContextReady?: (context: RenderedDomContext) => void;
  /**
   * Plugin overlays to render inside the editor viewport.
   * Passed from PluginHost to render plugin-specific overlays.
   */
  pluginOverlays?: ReactNode;
}

/**
 * DocxEditor ref interface
 */
export interface DocxEditorRef {
  /** Get the DocumentAgent for programmatic access */
  getAgent: () => DocumentAgent | null;
  /** Get the current document */
  getDocument: () => Document | null;
  /** Get the editor ref */
  getEditorRef: () => PagedEditorRef | null;
  /** Save the document to buffer */
  save: () => Promise<ArrayBuffer | null>;
  /** Set zoom level */
  setZoom: (zoom: number) => void;
  /** Get current zoom level */
  getZoom: () => number;
  /** Focus the editor */
  focus: () => void;
  /** Get current page number */
  getCurrentPage: () => number;
  /** Get total page count */
  getTotalPages: () => number;
  /** Scroll to a specific page */
  scrollToPage: (pageNumber: number) => void;
  /** Open print preview */
  openPrintPreview: () => void;
  /** Print the document directly */
  print: () => void;
  /** Current editing mode */
  getMode: () => EditorMode;
  /** Switch editing mode */
  setMode: (mode: EditorMode) => void;
}

/**
 * Editor internal state
 */
interface EditorState {
  isLoading: boolean;
  parseError: string | null;
  zoom: number;
  variableValues: Record<string, string>;
  isApplyingVariables: boolean;
  /** Current selection formatting for toolbar */
  selectionFormatting: SelectionFormatting;
  /** Current page number (1-indexed) */
  currentPage: number;
  /** Total page count */
  totalPages: number;
  /** ProseMirror table context (for showing table toolbar) */
  pmTableContext: TableContextInfo | null;
  /** Image context when cursor is on an image node */
  pmImageContext: {
    pos: number;
    wrapType: string;
    displayMode: string;
    cssFloat: string | null;
    transform: string | null;
    alt: string | null;
    borderWidth: number | null;
    borderColor: string | null;
    borderStyle: string | null;
  } | null;
}

// ============================================================================
// MAIN COMPONENT
// ============================================================================

/**
 * DocxEditor - Complete DOCX editor component
 */
const ALL_MODES: EditorMode[] = ['editing', 'suggesting', 'viewing'];

export type HeaderFooterAction = 'blank' | 'pageNumber' | 'pageXofY' | 'edit' | 'remove';

/** Content for Word-style built-in header/footer designs */
function headerFooterDesign(design: 'blank' | 'pageNumber' | 'pageXofY'): Paragraph[] {
  const text = (value: string): Run => ({ type: 'run', content: [{ type: 'text', text: value }] });
  const field = (fieldType: 'PAGE' | 'NUMPAGES'): ParagraphContent => ({
    type: 'simpleField',
    instruction: ` ${fieldType} `,
    fieldType,
    content: [text('1')],
  });
  if (design === 'blank') return [{ type: 'paragraph', content: [] }];
  const content: ParagraphContent[] =
    design === 'pageNumber'
      ? [field('PAGE')]
      : [text('Page '), field('PAGE'), text(' of '), field('NUMPAGES')];
  return [{ type: 'paragraph', formatting: { alignment: 'center' }, content }];
}

export const DocxEditor = forwardRef<DocxEditorRef, DocxEditorProps>(function DocxEditor(
  {
    documentBuffer,
    document: initialDocument,
    onSave,
    onChange,
    onSelectionChange,
    onError,
    onFontsLoaded: onFontsLoadedCallback,
    theme,
    showToolbar = true,
    showVariablePanel = true,
    showZoomControl = true,
    showPageNumbers = true,
    enablePageNavigation = true,
    pageNumberPosition = 'bottom-center',
    pageNumberVariant = 'default',
    showMarginGuides: _showMarginGuides = false,
    marginGuideColor: _marginGuideColor,
    showRuler = false,
    rulerUnit = 'inch',
    initialZoom = 1.0,
    readOnly: readOnlyProp = false,
    toolbarExtra,
    variablePanelPosition = 'right',
    variableDescriptions,
    className = '',
    style,
    placeholder,
    loadingIndicator,
    mode: modeProp,
    defaultMode = 'editing',
    onModeChange,
    availableModes = ALL_MODES,
    author = 'Anonymous',
    ribbonLayout: ribbonLayoutProp = 'simplified',
    showStatusBar = true,
    templateTagChips = true,
    documentName: documentNameProp = 'Document',
    onNew,
    onOpen,
    defaultReviewSidebarOpen = false,
    showPrintButton: _showPrintButton = true,
    printOptions: _printOptions,
    onPrint,
    onCopy: _onCopy,
    onCut: _onCut,
    onPaste: _onPaste,
    externalPlugins: externalPluginsProp,
    collaboration = null,
    onEditorViewReady,
    onRenderedDomContextReady,
    pluginOverlays,
  },
  ref
) {
  // State
  const [state, setState] = useState<EditorState>({
    isLoading: !!documentBuffer,
    parseError: null,
    zoom: initialZoom,
    variableValues: {},
    isApplyingVariables: false,
    selectionFormatting: {},
    currentPage: 1,
    totalPages: 1,
    pmTableContext: null,
    pmImageContext: null,
  });

  // Table properties dialog state
  const [tablePropsOpen, setTablePropsOpen] = useState(false);
  // Image position dialog state
  const [imagePositionOpen, setImagePositionOpen] = useState(false);
  // Image properties dialog state
  const [imagePropsOpen, setImagePropsOpen] = useState(false);
  // Footnote properties dialog state
  const [footnotePropsOpen, setFootnotePropsOpen] = useState(false);
  // Header/footer editing state
  const [hfEditPosition, setHfEditPosition] = useState<'header' | 'footer' | null>(null);

  // History hook for undo/redo - start with null document
  const history = useDocumentHistory<Document | null>(initialDocument || null, {
    maxEntries: 100,
    groupingInterval: 500,
    enableKeyboardShortcuts: true,
  });

  // Extension manager — built once; disable local history when collaborating
  // (y-prosemirror yUndoPlugin owns undo/redo instead).
  const collaborationSession = collaboration?.session ?? null;
  const [collabPeers, setCollabPeers] = useState<
    import('../collaboration').CollaborationAwarenessUser[]
  >([]);
  const [collabConnected, setCollabConnected] = useState(false);
  /**
   * Only bind y-prosemirror after the room has had a chance to receive peer
   * content (and we've seeded from the open document if the room is empty).
   * Until then the editor keeps editing the local open document.
   */
  const [collabBound, setCollabBound] = useState(false);
  const activeCollabSession = collabBound ? collaborationSession : null;

  useEffect(() => {
    if (!collaborationSession) {
      setCollabPeers([]);
      setCollabConnected(false);
      return;
    }
    const offPeers = collaborationSession.onPeersChange(setCollabPeers);
    const offStatus = collaborationSession.onStatusChange(setCollabConnected);
    return () => {
      offPeers();
      offStatus();
    };
  }, [collaborationSession]);

  const extensionManager = useMemo(() => {
    const mgr = new ExtensionManager(
      createStarterKit({
        disable: activeCollabSession ? ['history'] : [],
      })
    );
    mgr.buildSchema();
    mgr.initializeRuntime();
    return mgr;
  }, [activeCollabSession]);

  // Refs
  const pagedEditorRef = useRef<PagedEditorRef>(null);
  /** Nested PM view while editing header/footer — ribbon formats this view */
  const hfViewRef = useRef<PMEditorView | null>(null);
  const scrollContainerRef = useRef<HTMLDivElement>(null);
  const [rulerVisible, setRulerVisible] = useState(showRuler);
  const [ribbonLayout, setRibbonLayout] = useState<RibbonLayout>(ribbonLayoutProp);
  const [documentName, setDocumentName] = useState(documentNameProp);
  useEffect(() => setDocumentName(documentNameProp), [documentNameProp]);
  /** Bumped when File > New / Open loads a document, to remount the editor view */
  const [documentKey, setDocumentKey] = useState(0);

  // When collaboration starts: wait for peer sync, then seed from the document
  // already open (File → Open first, then Collaborate). If the room already has
  // content from a peer, bind to that instead of overwriting it.
  const collabRoomKey = collaborationSession?.roomId ?? '';
  const prevCollabRoomRef = useRef<string>('');
  useEffect(() => {
    const prevRoom = prevCollabRoomRef.current;
    prevCollabRoomRef.current = collabRoomKey;

    if (!collaborationSession) {
      setCollabBound(false);
      // Left a room — remount with local history plugins
      if (prevRoom) setDocumentKey((k) => k + 1);
      return;
    }

    setCollabBound(false);
    const cancel = whenRoomReady(collaborationSession, () => {
      const doc = history.state;
      if (doc && !isFragmentSeeded(collaborationSession)) {
        try {
          const pmDoc = toProseDoc(doc, { styles: doc.package.styles ?? undefined });
          seedFragmentFromProseDoc(collaborationSession, pmDoc);
        } catch (err) {
          console.warn('Failed to seed collaboration room from open document:', err);
        }
      }
      setCollabBound(true);
      // Remount the PM view so y-prosemirror plugins bind to this session
      setDocumentKey((k) => k + 1);
    });
    return cancel;
    // eslint-disable-next-line react-hooks/exhaustive-deps -- only when joining/leaving a room
  }, [collabRoomKey]);
  /** Undo depth at the last save — differs from the current depth when there are unsaved changes */
  const [savedUndoCount, setSavedUndoCount] = useState(0);
  const [pendingDiscard, setPendingDiscard] = useState<(() => void) | null>(null);
  const openInputRef = useRef<HTMLInputElement>(null);
  const [wordCountOpen, setWordCountOpen] = useState(false);

  // Review: editing modes, comments, suggestions, Word tags
  const review = useReview({
    getView: () => pagedEditorRef.current?.getView() ?? null,
    author,
    mode: readOnlyProp ? 'viewing' : modeProp,
    defaultMode,
    onModeChange,
    defaultSidebarOpen: defaultReviewSidebarOpen,
  });
  // Viewing mode is read-only; the readOnly prop locks the editor into it
  const readOnly = readOnlyProp || review.mode === 'viewing';
  const modeOptions = readOnlyProp ? (['viewing'] as EditorMode[]) : availableModes;
  // Comment balloons occupy the markup area to the right of the page
  const showBalloons =
    review.showComments &&
    (review.markupView === 'all' || review.markupView === 'simple') &&
    (review.liveComments.length > 0 || review.draft != null);
  const toggleTrackChanges = useCallback(() => {
    if (review.mode === 'viewing') return;
    review.setMode(review.mode === 'suggesting' ? 'editing' : 'suggesting');
  }, [review]);
  const externalPlugins = useMemo(
    () => [review.suggestionPlugin, review.sdtLockPlugin, ...(externalPluginsProp ?? [])],
    [review.suggestionPlugin, review.sdtLockPlugin, externalPluginsProp]
  );
  const agentRef = useRef<DocumentAgent | null>(null);
  const containerRef = useRef<HTMLDivElement>(null);
  // Save the last known selection for restoring after toolbar interactions
  const lastSelectionRef = useRef<{ from: number; to: number } | null>(null);
  // React state mirror so New Comment can stay enabled after selection collapses
  const [hasSavedSelection, setHasSavedSelection] = useState(false);
  const imageInputRef = useRef<HTMLInputElement>(null);

  // Helper to get the active editor's view (header/footer takes priority)
  const getActiveEditorView = useCallback(() => {
    return hfViewRef.current ?? pagedEditorRef.current?.getView() ?? null;
  }, []);

  // Helper to focus the active editor
  const focusActiveEditor = useCallback(() => {
    if (hfViewRef.current) {
      hfViewRef.current.focus();
      return;
    }
    pagedEditorRef.current?.focus();
  }, []);

  // Helper to undo in the active editor
  const undoActiveEditor = useCallback(() => {
    const hf = hfViewRef.current;
    if (hf) {
      pmUndo(hf.state, hf.dispatch);
      return;
    }
    const view = pagedEditorRef.current?.getView();
    if (activeCollabSession && view) {
      collabUndo(view.state);
      return;
    }
    pagedEditorRef.current?.undo();
  }, [activeCollabSession]);

  // Helper to redo in the active editor
  const redoActiveEditor = useCallback(() => {
    const hf = hfViewRef.current;
    if (hf) {
      pmRedo(hf.state, hf.dispatch);
      return;
    }
    const view = pagedEditorRef.current?.getView();
    if (activeCollabSession && view) {
      collabRedo(view.state);
      return;
    }
    pagedEditorRef.current?.redo();
  }, [activeCollabSession]);

  // Find/Replace hook
  const findReplace = useFindReplace();

  // Hyperlink dialog hook
  const hyperlinkDialog = useHyperlinkDialog();

  // Parse document buffer
  useEffect(() => {
    if (!documentBuffer) {
      if (initialDocument) {
        history.reset(initialDocument);
        review.loadComments(initialDocument);
        setState((prev) => ({ ...prev, isLoading: false }));
        // Load fonts for initial document
        loadDocumentFonts(initialDocument).catch((err) => {
          console.warn('Failed to load document fonts:', err);
        });
      }
      return;
    }

    setState((prev) => ({ ...prev, isLoading: true, parseError: null }));

    const parseDocument = async () => {
      try {
        const doc = await parseDocx(documentBuffer);
        // Reset history with parsed document (clears undo/redo stacks)
        history.reset(doc);
        review.loadComments(doc);
        setState((prev) => ({
          ...prev,
          isLoading: false,
          parseError: null,
        }));

        // Extract initial variable values
        if (doc.package.document) {
          const variables = extractVariables(doc);
          setState((prev) => ({ ...prev, variableValues: variables }));
        }

        // Load fonts used in the document from Google Fonts
        loadDocumentFonts(doc).catch((err) => {
          console.warn('Failed to load document fonts:', err);
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : 'Failed to parse document';
        setState((prev) => ({
          ...prev,
          isLoading: false,
          parseError: message,
        }));
        onError?.(error instanceof Error ? error : new Error(message));
      }
    };

    parseDocument();
  }, [documentBuffer, initialDocument, onError]); // eslint-disable-line react-hooks/exhaustive-deps

  // Update document when initialDocument changes
  useEffect(() => {
    if (initialDocument && !documentBuffer) {
      history.reset(initialDocument);
      review.loadComments(initialDocument);
    }
  }, [initialDocument, documentBuffer]); // eslint-disable-line react-hooks/exhaustive-deps

  // Create/update agent when document changes
  useEffect(() => {
    if (history.state) {
      agentRef.current = new DocumentAgent(history.state);
    } else {
      agentRef.current = null;
    }
  }, [history.state]);

  // Listen for font loading
  useEffect(() => {
    const cleanup = onFontsLoaded(() => {
      onFontsLoadedCallback?.();
    });
    return cleanup;
  }, [onFontsLoadedCallback]);

  // Handle document change
  const handleDocumentChange = useCallback(
    (newDocument: Document) => {
      history.push(newDocument);
      onChange?.(newDocument);
    },
    [onChange, history]
  );

  // Handle selection changes from ProseMirror
  const handleSelectionChange = useCallback(
    (selectionState: SelectionState | null) => {
      // Save selection for restoring after toolbar interactions
      const view = getActiveEditorView();
      if (view) {
        const { from, to } = view.state.selection;
        // Only save non-empty selections (when text is actually selected)
        if (from !== to) {
          lastSelectionRef.current = { from, to };
          setHasSavedSelection(true);
        }
      }

      // Also check table context from ProseMirror
      let pmTableCtx: TableContextInfo | null = null;
      if (view) {
        pmTableCtx = getTableContext(view.state);
        if (!pmTableCtx.isInTable) {
          pmTableCtx = null;
        }
      }

      // Check if cursor is on an image (NodeSelection)
      let pmImageCtx: typeof state.pmImageContext = null;
      if (view) {
        const sel = view.state.selection;
        // NodeSelection has a `node` property
        const selectedNode = (
          sel as { node?: { type: { name: string }; attrs: Record<string, unknown> } }
        ).node;
        if (selectedNode?.type.name === 'image') {
          pmImageCtx = {
            pos: sel.from,
            wrapType: (selectedNode.attrs.wrapType as string) ?? 'inline',
            displayMode: (selectedNode.attrs.displayMode as string) ?? 'inline',
            cssFloat: (selectedNode.attrs.cssFloat as string) ?? null,
            transform: (selectedNode.attrs.transform as string) ?? null,
            alt: (selectedNode.attrs.alt as string) ?? null,
            borderWidth: (selectedNode.attrs.borderWidth as number) ?? null,
            borderColor: (selectedNode.attrs.borderColor as string) ?? null,
            borderStyle: (selectedNode.attrs.borderStyle as string) ?? null,
          };
        }
      }

      if (!selectionState) {
        setState((prev) => ({
          ...prev,
          selectionFormatting: {},
          pmTableContext: pmTableCtx,
          pmImageContext: pmImageCtx,
        }));
        return;
      }

      // Update toolbar formatting from ProseMirror selection
      const { textFormatting, paragraphFormatting } = selectionState;

      // Extract font family (prefer ascii, fall back to hAnsi)
      const fontFamily = textFormatting.fontFamily?.ascii || textFormatting.fontFamily?.hAnsi;

      // Extract text color as hex string
      const textColor = textFormatting.color?.rgb ? `#${textFormatting.color.rgb}` : undefined;

      // Build list state from numPr
      const numPr = paragraphFormatting.numPr;
      const listState = numPr
        ? {
            type: (numPr.numId === 1 ? 'bullet' : 'numbered') as 'bullet' | 'numbered',
            level: numPr.ilvl ?? 0,
            isInList: true,
            numId: numPr.numId,
          }
        : undefined;

      const formatting: SelectionFormatting = {
        bold: textFormatting.bold,
        italic: textFormatting.italic,
        underline: !!textFormatting.underline,
        strike: textFormatting.strike,
        superscript: textFormatting.vertAlign === 'superscript',
        subscript: textFormatting.vertAlign === 'subscript',
        fontFamily,
        fontSize: textFormatting.fontSize,
        color: textColor,
        highlight: textFormatting.highlight,
        alignment: paragraphFormatting.alignment,
        lineSpacing: paragraphFormatting.lineSpacing,
        listState,
        styleId: selectionState.styleId ?? undefined,
        indentLeft: paragraphFormatting.indentLeft,
      };
      setState((prev) => ({
        ...prev,
        selectionFormatting: formatting,
        pmTableContext: pmTableCtx,
        pmImageContext: pmImageCtx,
      }));

      // Notify parent
      onSelectionChange?.(selectionState);
    },
    [onSelectionChange]
  );

  // Table selection hook
  const tableSelection = useTableSelection({
    document: history.state,
    onChange: handleDocumentChange,
    onSelectionChange: (_context) => {
      // Could notify parent of table selection changes
    },
  });

  // Word shortcuts: Ctrl+Alt+M new comment, Ctrl+Shift+E Track Changes, Ctrl+Enter page break
  const { withComments, startComment } = review;
  /** Restore saved selection (if needed) then open the comment composer */
  const handleNewComment = useCallback(() => {
    const view = pagedEditorRef.current?.getView();
    if (!view || readOnly) return;
    // Capture before focus — focusing the hidden PM can move the live selection
    // away from the text the user highlighted on the painted page.
    const saved = lastSelectionRef.current;
    view.focus();
    if (saved && saved.from !== saved.to) {
      try {
        view.dispatch(
          view.state.tr.setSelection(TextSelection.create(view.state.doc, saved.from, saved.to))
        );
      } catch {
        // startComment still receives the saved range as fallback
      }
    }
    startComment(saved);
  }, [readOnly, startComment]);
  const shortcutHandlersRef = useRef({
    startComment: handleNewComment,
    toggleTrackChanges: () => {},
    insertPageBreak: () => {},
    save: () => {},
    open: () => {},
  });
  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;
    const handleWordShortcuts = (e: KeyboardEvent) => {
      const mod = e.ctrlKey || e.metaKey;
      if (!mod) return;
      // Leave shortcuts alone while typing in comment boxes, dialogs and inputs
      const target = e.target as HTMLElement | null;
      if (target?.closest('input, textarea, select, .ep-balloons, .ep-dialog, .ep-backstage')) {
        return;
      }
      const handlers = shortcutHandlersRef.current;
      if (e.altKey && e.code === 'KeyM') {
        e.preventDefault();
        handlers.startComment();
      } else if (e.shiftKey && !e.altKey && e.code === 'KeyE') {
        e.preventDefault();
        handlers.toggleTrackChanges();
      } else if (!e.shiftKey && !e.altKey && e.key === 'Enter') {
        e.preventDefault();
        handlers.insertPageBreak();
      } else if (!e.shiftKey && !e.altKey && e.code === 'KeyS') {
        e.preventDefault();
        handlers.save();
      } else if (!e.shiftKey && !e.altKey && e.code === 'KeyO') {
        e.preventDefault();
        handlers.open();
      }
    };
    container.addEventListener('keydown', handleWordShortcuts);
    return () => container.removeEventListener('keydown', handleWordShortcuts);
  }, [state.isLoading]);

  // Keyboard shortcuts for Find/Replace (Ctrl+F, Ctrl+H) and delete table selection
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      // Check for Ctrl+F (Find) or Ctrl+H (Replace)
      const isMac = navigator.platform.toUpperCase().indexOf('MAC') >= 0;
      const cmdOrCtrl = isMac ? e.metaKey : e.ctrlKey;

      // Delete selected table from layout selection (non-ProseMirror selection)
      if (!cmdOrCtrl && !e.shiftKey && !e.altKey) {
        if (e.key === 'Delete' || e.key === 'Backspace') {
          // If full table is selected via ProseMirror CellSelection, delete it.
          const view = pagedEditorRef.current?.getView();
          if (view) {
            const sel = view.state.selection as { $anchorCell?: unknown; forEachCell?: unknown };
            const isCellSel = '$anchorCell' in sel && typeof sel.forEachCell === 'function';
            if (isCellSel) {
              const context = getTableContext(view.state);
              if (context.isInTable && context.table) {
                let totalCells = 0;
                context.table.descendants((node) => {
                  if (node.type.name === 'tableCell' || node.type.name === 'tableHeader') {
                    totalCells += 1;
                  }
                });
                let selectedCells = 0;
                (sel as { forEachCell: (fn: () => void) => void }).forEachCell(() => {
                  selectedCells += 1;
                });
                if (totalCells > 0 && selectedCells >= totalCells) {
                  e.preventDefault();
                  pmDeleteTable(view.state, view.dispatch);
                  return;
                }
              }
            }
          }

          if (tableSelection.state.tableIndex !== null) {
            e.preventDefault();
            tableSelection.handleAction('deleteTable');
            return;
          }
        }
      }

      if (cmdOrCtrl && !e.shiftKey && !e.altKey) {
        if (e.key.toLowerCase() === 'f') {
          e.preventDefault();
          // Get selected text if any
          const selection = window.getSelection();
          const selectedText = selection && !selection.isCollapsed ? selection.toString() : '';
          findReplace.openFind(selectedText);
        } else if (e.key.toLowerCase() === 'h') {
          e.preventDefault();
          // Get selected text if any
          const selection = window.getSelection();
          const selectedText = selection && !selection.isCollapsed ? selection.toString() : '';
          findReplace.openReplace(selectedText);
        } else if (e.key.toLowerCase() === 'k') {
          e.preventDefault();
          // Open hyperlink dialog
          const view = pagedEditorRef.current?.getView();
          if (view) {
            const selectedText = getSelectedText(view.state);
            const existingLink = getHyperlinkAttrs(view.state);
            if (existingLink) {
              hyperlinkDialog.openEdit({
                url: existingLink.href,
                displayText: selectedText,
                tooltip: existingLink.tooltip,
              });
            } else {
              hyperlinkDialog.openInsert(selectedText);
            }
          }
        }
      }
    };

    document.addEventListener('keydown', handleKeyDown);
    return () => {
      document.removeEventListener('keydown', handleKeyDown);
    };
  }, [findReplace, hyperlinkDialog, tableSelection]);

  // Handle table insert from toolbar
  const handleInsertTable = useCallback(
    (rows: number, columns: number) => {
      const view = getActiveEditorView();
      if (!view) return;
      insertTable(rows, columns)(view.state, view.dispatch);
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor]
  );

  // Trigger file picker for image insert
  const handleInsertImageClick = useCallback(() => {
    imageInputRef.current?.click();
  }, []);

  // Handle file selection for image insert
  const handleImageFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (!file) return;

      const view = getActiveEditorView();
      if (!view) return;

      const reader = new FileReader();
      reader.onload = () => {
        const dataUrl = reader.result as string;

        // Create an Image element to get natural dimensions
        const img = new Image();
        img.onload = () => {
          let width = img.naturalWidth;
          let height = img.naturalHeight;

          // Constrain to reasonable max width (612px ~ 6.375 inches at 96dpi)
          const maxWidth = 612;
          if (width > maxWidth) {
            const scale = maxWidth / width;
            width = maxWidth;
            height = Math.round(height * scale);
          }

          const rId = `rId_img_${Date.now()}`;
          const imageNode = view.state.schema.nodes.image.create({
            src: dataUrl,
            alt: file.name,
            width,
            height,
            rId,
            wrapType: 'inline',
            displayMode: 'inline',
          });

          const { from } = view.state.selection;
          const tr = view.state.tr.insert(from, imageNode);
          view.dispatch(tr.scrollIntoView());
          focusActiveEditor();
        };
        img.src = dataUrl;
      };
      reader.readAsDataURL(file);

      // Reset the input so the same file can be selected again
      e.target.value = '';
    },
    [getActiveEditorView, focusActiveEditor]
  );

  // Handle shape insertion
  // Handle image wrap type change
  const handleImageWrapType = useCallback(
    (wrapType: string) => {
      const view = getActiveEditorView();
      if (!view || !state.pmImageContext) return;

      const pos = state.pmImageContext.pos;
      const node = view.state.doc.nodeAt(pos);
      if (!node || node.type.name !== 'image') return;

      // Map wrap type to display mode + cssFloat
      let displayMode = 'inline';
      let cssFloat: string | null = null;

      switch (wrapType) {
        case 'inline':
          displayMode = 'inline';
          cssFloat = null;
          break;
        case 'square':
        case 'tight':
        case 'through':
          displayMode = 'float';
          cssFloat = 'left';
          break;
        case 'topAndBottom':
          displayMode = 'block';
          cssFloat = null;
          break;
        case 'behind':
        case 'inFront':
          displayMode = 'float';
          cssFloat = 'none';
          break;
        case 'wrapLeft':
          displayMode = 'float';
          cssFloat = 'right';
          wrapType = 'square';
          break;
        case 'wrapRight':
          displayMode = 'float';
          cssFloat = 'left';
          wrapType = 'square';
          break;
      }

      const tr = view.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        wrapType,
        displayMode,
        cssFloat,
      });
      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor, state.pmImageContext]
  );

  // Handle image transform (rotate/flip)
  const handleImageTransform = useCallback(
    (action: 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV') => {
      const view = getActiveEditorView();
      if (!view || !state.pmImageContext) return;

      const pos = state.pmImageContext.pos;
      const node = view.state.doc.nodeAt(pos);
      if (!node || node.type.name !== 'image') return;

      const currentTransform = (node.attrs.transform as string) || '';

      // Parse current rotation and flip state
      const rotateMatch = currentTransform.match(/rotate\((-?\d+(?:\.\d+)?)deg\)/);
      let rotation = rotateMatch ? parseFloat(rotateMatch[1]) : 0;
      let hasFlipH = /scaleX\(-1\)/.test(currentTransform);
      let hasFlipV = /scaleY\(-1\)/.test(currentTransform);

      switch (action) {
        case 'rotateCW':
          rotation = (rotation + 90) % 360;
          break;
        case 'rotateCCW':
          rotation = (rotation - 90 + 360) % 360;
          break;
        case 'flipH':
          hasFlipH = !hasFlipH;
          break;
        case 'flipV':
          hasFlipV = !hasFlipV;
          break;
      }

      // Build new transform string
      const parts: string[] = [];
      if (rotation !== 0) parts.push(`rotate(${rotation}deg)`);
      if (hasFlipH) parts.push('scaleX(-1)');
      if (hasFlipV) parts.push('scaleY(-1)');
      const newTransform = parts.length > 0 ? parts.join(' ') : null;

      const tr = view.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        transform: newTransform,
      });
      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor, state.pmImageContext]
  );

  // Open image position dialog
  const handleOpenImagePosition = useCallback(() => {
    setImagePositionOpen(true);
  }, []);

  // Apply image position changes
  const handleApplyImagePosition = useCallback(
    (data: ImagePositionData) => {
      const view = getActiveEditorView();
      if (!view || !state.pmImageContext) return;

      const pos = state.pmImageContext.pos;
      const node = view.state.doc.nodeAt(pos);
      if (!node || node.type.name !== 'image') return;

      const tr = view.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        position: {
          horizontal: data.horizontal,
          vertical: data.vertical,
        },
        distTop: data.distTop ?? node.attrs.distTop,
        distBottom: data.distBottom ?? node.attrs.distBottom,
        distLeft: data.distLeft ?? node.attrs.distLeft,
        distRight: data.distRight ?? node.attrs.distRight,
      });
      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor, state.pmImageContext]
  );

  // Open image properties dialog
  const handleOpenImageProperties = useCallback(() => {
    setImagePropsOpen(true);
  }, []);

  // Apply image properties (alt text + border)
  const handleApplyImageProperties = useCallback(
    (data: ImagePropertiesData) => {
      const view = getActiveEditorView();
      if (!view || !state.pmImageContext) return;

      const pos = state.pmImageContext.pos;
      const node = view.state.doc.nodeAt(pos);
      if (!node || node.type.name !== 'image') return;

      const tr = view.state.tr.setNodeMarkup(pos, undefined, {
        ...node.attrs,
        alt: data.alt ?? null,
        borderWidth: data.borderWidth ?? null,
        borderColor: data.borderColor ?? null,
        borderStyle: data.borderStyle ?? null,
      });
      view.dispatch(tr.scrollIntoView());
      focusActiveEditor();
    },
    [getActiveEditorView, focusActiveEditor, state.pmImageContext]
  );

  // Handle footnote/endnote properties update
  const handleApplyFootnoteProperties = useCallback(
    (
      footnotePr: import('../types/document').FootnoteProperties,
      endnotePr: import('../types/document').EndnoteProperties
    ) => {
      if (!history.state?.package) return;
      const newDoc = {
        ...history.state.package.document,
        finalSectionProperties: {
          ...history.state.package.document.finalSectionProperties,
          footnotePr,
          endnotePr,
        },
      };
      history.push({
        ...history.state,
        package: {
          ...history.state.package,
          document: newDoc,
        },
      });
    },
    [history]
  );

  // Handle table action from Toolbar - use ProseMirror commands
  const handleTableAction = useCallback(
    (action: TableAction) => {
      const view = getActiveEditorView();
      if (!view) return;

      switch (action) {
        case 'addRowAbove':
          addRowAbove(view.state, view.dispatch);
          break;
        case 'addRowBelow':
          addRowBelow(view.state, view.dispatch);
          break;
        case 'addColumnLeft':
          addColumnLeft(view.state, view.dispatch);
          break;
        case 'addColumnRight':
          addColumnRight(view.state, view.dispatch);
          break;
        case 'deleteRow':
          pmDeleteRow(view.state, view.dispatch);
          break;
        case 'deleteColumn':
          pmDeleteColumn(view.state, view.dispatch);
          break;
        case 'deleteTable':
          pmDeleteTable(view.state, view.dispatch);
          break;
        case 'selectTable':
          pmSelectTable(view.state, view.dispatch);
          break;
        case 'selectRow':
          pmSelectRow(view.state, view.dispatch);
          break;
        case 'selectColumn':
          pmSelectColumn(view.state, view.dispatch);
          break;
        case 'mergeCells':
          pmMergeCells(view.state, view.dispatch);
          break;
        case 'splitCell':
          pmSplitCell(view.state, view.dispatch);
          break;
        // Border actions
        case 'borderAll':
          setAllTableBorders(view.state, view.dispatch);
          break;
        case 'borderOutside':
          setOutsideTableBorders(view.state, view.dispatch);
          break;
        case 'borderInside':
          setInsideTableBorders(view.state, view.dispatch);
          break;
        case 'borderNone':
          removeTableBorders(view.state, view.dispatch);
          break;
        // Per-side border actions (toggle with default style)
        case 'borderTop':
          setCellBorder('top', { style: 'single', size: 4, color: { rgb: '000000' } })(
            view.state,
            view.dispatch
          );
          break;
        case 'borderBottom':
          setCellBorder('bottom', { style: 'single', size: 4, color: { rgb: '000000' } })(
            view.state,
            view.dispatch
          );
          break;
        case 'borderLeft':
          setCellBorder('left', { style: 'single', size: 4, color: { rgb: '000000' } })(
            view.state,
            view.dispatch
          );
          break;
        case 'borderRight':
          setCellBorder('right', { style: 'single', size: 4, color: { rgb: '000000' } })(
            view.state,
            view.dispatch
          );
          break;
        default:
          // Handle complex actions (with parameters)
          if (typeof action === 'object') {
            if (action.type === 'cellFillColor') {
              setCellFillColor(action.color)(view.state, view.dispatch);
            } else if (action.type === 'borderColor') {
              setTableBorderColor(action.color)(view.state, view.dispatch);
            } else if (action.type === 'cellBorder') {
              setCellBorder(action.side, {
                style: action.style,
                size: action.size,
                color: { rgb: action.color.replace(/^#/, '') },
              })(view.state, view.dispatch);
            } else if (action.type === 'cellVerticalAlign') {
              setCellVerticalAlign(action.align)(view.state, view.dispatch);
            } else if (action.type === 'cellMargins') {
              setCellMargins(action.margins)(view.state, view.dispatch);
            } else if (action.type === 'cellTextDirection') {
              setCellTextDirection(action.direction)(view.state, view.dispatch);
            } else if (action.type === 'toggleNoWrap') {
              toggleNoWrap()(view.state, view.dispatch);
            } else if (action.type === 'rowHeight') {
              setRowHeight(action.height, action.rule)(view.state, view.dispatch);
            } else if (action.type === 'toggleHeaderRow') {
              toggleHeaderRow()(view.state, view.dispatch);
            } else if (action.type === 'distributeColumns') {
              distributeColumns()(view.state, view.dispatch);
            } else if (action.type === 'autoFitContents') {
              autoFitContents()(view.state, view.dispatch);
            } else if (action.type === 'openTableProperties') {
              setTablePropsOpen(true);
            } else if (action.type === 'tableProperties') {
              setTableProperties(action.props)(view.state, view.dispatch);
            } else if (action.type === 'applyTableStyle') {
              // Resolve style data from built-in presets or document styles
              let preset: TableStylePreset | undefined = getBuiltinTableStyle(action.styleId);
              if (!preset && history.state?.package.styles) {
                const styleResolver = createStyleResolver(history.state.package.styles);
                const docStyle = styleResolver.getStyle(action.styleId);
                if (docStyle) {
                  // Convert to preset inline (same as documentStyleToPreset)
                  preset = { id: docStyle.styleId, name: docStyle.name ?? docStyle.styleId };
                  if (docStyle.tblPr?.borders) {
                    const b = docStyle.tblPr.borders;
                    preset.tableBorders = {};
                    for (const side of [
                      'top',
                      'bottom',
                      'left',
                      'right',
                      'insideH',
                      'insideV',
                    ] as const) {
                      const bs = b[side];
                      if (bs) {
                        preset.tableBorders[side] = {
                          style: bs.style,
                          size: bs.size,
                          color: bs.color?.rgb ? { rgb: bs.color.rgb } : undefined,
                        };
                      }
                    }
                  }
                  if (docStyle.tblStylePr) {
                    preset.conditionals = {};
                    for (const cond of docStyle.tblStylePr) {
                      const entry: Record<string, unknown> = {};
                      if (cond.tcPr?.shading?.fill)
                        entry.backgroundColor = `#${cond.tcPr.shading.fill}`;
                      if (cond.tcPr?.borders) {
                        const borders: Record<string, unknown> = {};
                        for (const s of ['top', 'bottom', 'left', 'right'] as const) {
                          const bs2 = cond.tcPr.borders[s];
                          if (bs2)
                            borders[s] = {
                              style: bs2.style,
                              size: bs2.size,
                              color: bs2.color?.rgb ? { rgb: bs2.color.rgb } : undefined,
                            };
                        }
                        entry.borders = borders;
                      }
                      if (cond.rPr?.bold) entry.bold = true;
                      if (cond.rPr?.color?.rgb) entry.color = `#${cond.rPr.color.rgb}`;
                      // eslint-disable-next-line @typescript-eslint/no-explicit-any
                      (preset.conditionals as any)[cond.type] = entry;
                    }
                  }
                  preset.look = { firstRow: true, lastRow: false, noHBand: false, noVBand: true };
                }
              }
              if (preset) {
                applyTableStyle({
                  styleId: preset.id,
                  tableBorders: preset.tableBorders,
                  conditionals: preset.conditionals,
                  look: preset.look,
                })(view.state, view.dispatch);
              }
            }
          } else {
            // Fallback to legacy table selection handler for other actions
            tableSelection.handleAction(action);
          }
      }

      focusActiveEditor();
    },
    [tableSelection, getActiveEditorView, focusActiveEditor]
  );

  // Handle formatting action from toolbar (body or header/footer view)
  const handleFormat = useCallback((action: FormattingAction) => {
    const view = getActiveEditorView();
    if (!view) return;

    // Focus editor first to ensure we can dispatch commands
    view.focus();

    // Restore selection if it was lost during toolbar interaction
    // This happens when user clicks on dropdown menus (font picker, style picker, etc.)
    const { from, to } = view.state.selection;
    const isEmptySelection = from === to;
    const savedSelection = lastSelectionRef.current;

    if (isEmptySelection && savedSelection && savedSelection.from !== savedSelection.to) {
      // Selection was lost - restore it before applying the format
      try {
        const tr = view.state.tr.setSelection(
          TextSelection.create(view.state.doc, savedSelection.from, savedSelection.to)
        );
        view.dispatch(tr);
      } catch (e) {
        // If restoration fails (e.g., positions are invalid after doc change), continue with current selection
        console.warn('Could not restore selection:', e);
      }
    }

    // Handle simple toggle actions
    if (action === 'bold') {
      toggleBold(view.state, view.dispatch);
      return;
    }
    if (action === 'italic') {
      toggleItalic(view.state, view.dispatch);
      return;
    }
    if (action === 'underline') {
      toggleUnderline(view.state, view.dispatch);
      return;
    }
    if (action === 'strikethrough') {
      toggleStrike(view.state, view.dispatch);
      return;
    }
    if (action === 'superscript') {
      toggleSuperscript(view.state, view.dispatch);
      return;
    }
    if (action === 'subscript') {
      toggleSubscript(view.state, view.dispatch);
      return;
    }
    if (action === 'bulletList') {
      toggleBulletList(view.state, view.dispatch);
      return;
    }
    if (action === 'numberedList') {
      toggleNumberedList(view.state, view.dispatch);
      return;
    }
    if (action === 'indent') {
      // Try list indent first, then paragraph indent
      if (!increaseListLevel(view.state, view.dispatch)) {
        increaseIndent()(view.state, view.dispatch);
      }
      return;
    }
    if (action === 'outdent') {
      // Try list outdent first, then paragraph outdent
      if (!decreaseListLevel(view.state, view.dispatch)) {
        decreaseIndent()(view.state, view.dispatch);
      }
      return;
    }
    if (action === 'clearFormatting') {
      clearFormatting(view.state, view.dispatch);
      return;
    }
    if (action === 'insertLink') {
      // Get the selected text for the hyperlink dialog
      const selectedText = getSelectedText(view.state);
      // Check if we're editing an existing link
      const existingLink = getHyperlinkAttrs(view.state);
      if (existingLink) {
        hyperlinkDialog.openEdit({
          url: existingLink.href,
          displayText: selectedText,
          tooltip: existingLink.tooltip,
        });
      } else {
        hyperlinkDialog.openInsert(selectedText);
      }
      return;
    }

    // Handle object-based actions
    if (typeof action === 'object') {
      switch (action.type) {
        case 'alignment':
          setAlignment(action.value)(view.state, view.dispatch);
          break;
        case 'textColor':
          // action.value can be a string like "#FF0000" or a color name
          setTextColor({ rgb: action.value.replace('#', '') })(view.state, view.dispatch);
          break;
        case 'highlightColor': {
          // Convert hex to OOXML named highlight value (e.g., 'FFFF00' → 'yellow')
          const highlightName = action.value ? mapHexToHighlightName(action.value) : '';
          setHighlight(highlightName || action.value)(view.state, view.dispatch);
          break;
        }
        case 'fontSize':
          // Convert points to half-points (OOXML uses half-points for font sizes)
          setFontSize(pointsToHalfPoints(action.value))(view.state, view.dispatch);
          break;
        case 'fontFamily':
          setFontFamily(action.value)(view.state, view.dispatch);
          break;
        case 'lineSpacing':
          setLineSpacing(action.value)(view.state, view.dispatch);
          break;
        case 'applyStyle': {
          // Resolve style to get its formatting properties
          const styleResolver = history.state?.package.styles
            ? createStyleResolver(history.state.package.styles)
            : null;

          if (styleResolver) {
            const resolved = styleResolver.resolveParagraphStyle(action.value);
            applyStyle(action.value, {
              paragraphFormatting: resolved.paragraphFormatting,
              runFormatting: resolved.runFormatting,
            })(view.state, view.dispatch);
          } else {
            // No styles available, just set the styleId
            applyStyle(action.value)(view.state, view.dispatch);
          }
          break;
        }
      }
    }
  }, [getActiveEditorView]);

  // Handle variable values change
  const handleVariableValuesChange = useCallback((values: Record<string, string>) => {
    setState((prev) => ({ ...prev, variableValues: values }));
  }, []);

  // Handle apply variables
  const handleApplyVariables = useCallback(
    async (values: Record<string, string>) => {
      if (!agentRef.current) return;

      setState((prev) => ({ ...prev, isApplyingVariables: true }));

      try {
        const newDoc = agentRef.current.setVariables(values).getDocument();
        handleDocumentChange(newDoc);
      } catch (error) {
        onError?.(error instanceof Error ? error : new Error('Failed to apply variables'));
      } finally {
        setState((prev) => ({ ...prev, isApplyingVariables: false }));
      }
    },
    [handleDocumentChange, onError]
  );

  // Handle zoom change
  const handleZoomChange = useCallback((zoom: number) => {
    setState((prev) => ({ ...prev, zoom }));
  }, []);

  // Handle hyperlink dialog submit
  const handleHyperlinkSubmit = useCallback(
    (data: HyperlinkData) => {
      const view = getActiveEditorView();
      if (!view) return;

      const url = data.url || '';
      const tooltip = data.tooltip;

      // Check if we have a selection
      const { empty } = view.state.selection;

      if (empty && data.displayText) {
        // No selection but display text provided - insert new linked text
        insertHyperlink(data.displayText, url, tooltip)(view.state, view.dispatch);
      } else if (!empty) {
        // Have selection - apply hyperlink to it
        setHyperlink(url, tooltip)(view.state, view.dispatch);
      } else if (data.displayText) {
        // Empty selection but display text provided
        insertHyperlink(data.displayText, url, tooltip)(view.state, view.dispatch);
      }

      hyperlinkDialog.close();
      focusActiveEditor();
    },
    [hyperlinkDialog, getActiveEditorView, focusActiveEditor]
  );

  // Handle hyperlink removal
  const handleHyperlinkRemove = useCallback(() => {
    const view = getActiveEditorView();
    if (!view) return;

    removeHyperlink(view.state, view.dispatch);
    hyperlinkDialog.close();
    focusActiveEditor();
  }, [hyperlinkDialog, getActiveEditorView, focusActiveEditor]);

  // Handle margin changes from rulers
  const createMarginHandler = useCallback(
    (property: 'marginLeft' | 'marginRight' | 'marginTop' | 'marginBottom') =>
      (marginTwips: number) => {
        if (!history.state || readOnly) return;
        const newDoc = {
          ...history.state,
          package: {
            ...history.state.package,
            document: {
              ...history.state.package.document,
              finalSectionProperties: {
                ...history.state.package.document.finalSectionProperties,
                [property]: marginTwips,
              },
            },
          },
        };
        handleDocumentChange(newDoc);
      },
    [history.state, readOnly, handleDocumentChange]
  );

  const handleLeftMarginChange = useMemo(
    () => createMarginHandler('marginLeft'),
    [createMarginHandler]
  );
  const handleRightMarginChange = useMemo(
    () => createMarginHandler('marginRight'),
    [createMarginHandler]
  );
  const handleTopMarginChange = useMemo(
    () => createMarginHandler('marginTop'),
    [createMarginHandler]
  );
  const handleBottomMarginChange = useMemo(
    () => createMarginHandler('marginBottom'),
    [createMarginHandler]
  );

  // Handle page navigation (from PageNavigator)
  // TODO: Implement page navigation in ProseMirror
  const handlePageNavigate = useCallback((_pageNumber: number) => {
    // Page navigation not yet implemented for ProseMirror
  }, []);

  // Handle save
  const handleSave = useCallback(async (): Promise<ArrayBuffer | null> => {
    if (!history.state) return null;

    try {
      const buffer = await new DocumentAgent(withComments(history.state)).toBuffer();
      setSavedUndoCount(history.undoCount);
      onSave?.(buffer);
      return buffer;
    } catch (error) {
      onError?.(error instanceof Error ? error : new Error('Failed to save document'));
      return null;
    }
  }, [onSave, onError, history.state, history.undoCount, withComments]);

  // Handle error from editor
  const handleEditorError = useCallback(
    (error: Error) => {
      onError?.(error);
    },
    [onError]
  );

  /**
   * Print (and Save as PDF via the print dialog). `withMarkup` prints tracked
   * changes as markup; otherwise the final text (or original, in Original view).
   */
  const handlePrint = useCallback(
    (withMarkup?: boolean) => {
      const printMarkup = withMarkup ?? review.markupView === 'all';
      // Find the pages container and clone its content into a clean print window
      const pagesEl = containerRef.current?.querySelector('.paged-editor__pages');
      if (!pagesEl) {
        window.print();
        onPrint?.();
        return;
      }

      const printWindow = window.open('', '_blank');
      if (!printWindow) {
        // Popup blocked — fall back to window.print()
        window.print();
        onPrint?.();
        return;
      }

      // Collect all @font-face rules from the current page
      const fontFaceRules: string[] = [];
      for (const sheet of Array.from(document.styleSheets)) {
        try {
          for (const rule of Array.from(sheet.cssRules)) {
            if (rule instanceof CSSFontFaceRule) {
              fontFaceRules.push(rule.cssText);
            }
          }
        } catch {
          // Cross-origin stylesheets can't be read — skip
        }
      }

      // Clone pages and remove transforms/shadows
      const pagesClone = pagesEl.cloneNode(true) as HTMLElement;
      pagesClone.style.cssText = 'display: block; margin: 0; padding: 0;';
      for (const page of Array.from(pagesClone.querySelectorAll('.layout-page'))) {
        const el = page as HTMLElement;
        el.style.boxShadow = 'none';
        el.style.margin = '0';
      }

      // Tracked changes print as shown on screen: with markup, or as the final text
      const reviewCss = printMarkup
        ? `.docx-revision-ins { color: #1b5e20 !important; text-decoration-line: underline !important; }
.docx-revision-del { color: #b71c1c !important; text-decoration-line: line-through !important; }`
        : review.markupView === 'original'
          ? `.docx-revision-ins { display: none !important; }`
          : `.docx-revision-del { display: none !important; }`;
      const title = (document.title || 'Document').replace(/[<>&]/g, '');

      printWindow.document.write(`<!DOCTYPE html>
<html><head><title>${title}</title>
<style>
${fontFaceRules.join('\n')}
${reviewCss}
* { margin: 0; padding: 0; }
body { background: white; }
.layout-page { break-after: page; }
.layout-page:last-child { break-after: auto; }
@page { margin: 0; size: auto; }
</style>
</head><body>${pagesClone.outerHTML}</body></html>`);
      printWindow.document.close();

      // Wait for fonts/images then print
      printWindow.onload = () => {
        printWindow.print();
        printWindow.close();
      };

      // Fallback if onload doesn't fire (some browsers)
      setTimeout(() => {
        if (!printWindow.closed) {
          printWindow.print();
          printWindow.close();
        }
      }, 1000);

      onPrint?.();
    },
    [onPrint, review.markupView]
  );
  const handleDirectPrint = useCallback(() => handlePrint(), [handlePrint]);

  // File > Download a Copy
  const handleDownload = useCallback(async () => {
    const buffer = await handleSave();
    if (!buffer) return;
    const blob = new Blob([buffer], {
      type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = /\.docx$/i.test(documentName) ? documentName : `${documentName}.docx`;
    link.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }, [handleSave, documentName]);

  const hasUnsavedChanges = history.undoCount !== savedUndoCount;

  /** Run an action that replaces the document, asking first if there are unsaved changes */
  const confirmDiscard = useCallback(
    (action: () => void) => {
      if (hasUnsavedChanges) setPendingDiscard(() => action);
      else action();
    },
    [hasUnsavedChanges]
  );

  /** Load a DOCX buffer as the current document (File > New / Open) */
  const loadDocumentBuffer = useCallback(
    async (buffer: ArrayBuffer, name: string) => {
      try {
        const doc = await parseDocx(buffer);
        history.reset(doc);
        review.loadComments(doc);
        // Collaboration keeps content in Yjs — replace the shared fragment so
        // Open/New doesn't leave the previous (often empty) room document on screen.
        if (collaborationSession) {
          try {
            const pmDoc = toProseDoc(doc, { styles: doc.package.styles ?? undefined });
            replaceFragmentWithProseDoc(collaborationSession, pmDoc);
          } catch (err) {
            console.warn('Failed to sync opened document into collaboration room:', err);
          }
        }
        setDocumentName(name);
        setDocumentKey((k) => k + 1);
        setSavedUndoCount(0);
        setState((prev) => ({
          ...prev,
          parseError: null,
          variableValues: doc.package.document ? extractVariables(doc) : {},
        }));
        loadDocumentFonts(doc).catch((err) => {
          console.warn('Failed to load document fonts:', err);
        });
        return true;
      } catch (error) {
        onError?.(error instanceof Error ? error : new Error('Failed to open document'));
        return false;
      }
    },
    [history, review, onError, collaborationSession]
  );

  // File > New: blank document
  const handleNewDocument = useCallback(() => {
    confirmDiscard(async () => {
      if (await loadDocumentBuffer(await createEmptyDocx(), 'Document1.docx')) onNew?.();
    });
  }, [confirmDiscard, loadDocumentBuffer, onNew]);

  // File > Open: pick a .docx
  const handleOpenDocument = useCallback(() => {
    confirmDiscard(() => openInputRef.current?.click());
  }, [confirmDiscard]);

  const handleOpenFileChange = useCallback(
    async (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      e.target.value = '';
      if (!file) return;
      if (await loadDocumentBuffer(await file.arrayBuffer(), file.name)) onOpen?.(file);
    },
    [loadDocumentBuffer, onOpen]
  );

  // File > Save: hand the DOCX to the host (onSave), or download it when there's no host handler
  const handleSaveCommand = useCallback(async () => {
    if (onSave) await handleSave();
    else await handleDownload();
  }, [onSave, handleSave, handleDownload]);

  // Insert > Page Break (Ctrl+Enter)
  const handleInsertPageBreak = useCallback(() => {
    const view = pagedEditorRef.current?.getView();
    if (!view || readOnly) return;
    insertPageBreak(view);
    pagedEditorRef.current?.focus();
  }, [readOnly]);

  shortcutHandlersRef.current = {
    startComment: handleNewComment,
    toggleTrackChanges,
    insertPageBreak: handleInsertPageBreak,
    save: () => void handleSaveCommand(),
    open: handleOpenDocument,
  };

  // Track page count and the page in view (status bar "Page X of Y")
  useEffect(() => {
    const scroller = scrollContainerRef.current;
    if (!scroller) return;
    let frame = 0;
    const update = () => {
      cancelAnimationFrame(frame);
      frame = requestAnimationFrame(() => {
        const pages = scroller.querySelectorAll('.layout-page');
        const total = Math.max(1, pages.length);
        const viewRect = scroller.getBoundingClientRect();
        const probe = viewRect.top + Math.min(viewRect.height / 3, 200);
        let current = 1;
        pages.forEach((page, i) => {
          if (page.getBoundingClientRect().top <= probe) current = i + 1;
        });
        setState((prev) =>
          prev.totalPages === total && prev.currentPage === current
            ? prev
            : { ...prev, totalPages: total, currentPage: current }
        );
      });
    };
    const pagesEl = scroller.querySelector('.paged-editor__pages');
    const observer = new MutationObserver(update);
    if (pagesEl) observer.observe(pagesEl, { childList: true });
    scroller.addEventListener('scroll', update, { passive: true });
    update();
    return () => {
      cancelAnimationFrame(frame);
      observer.disconnect();
      scroller.removeEventListener('scroll', update);
    };
  }, [state.isLoading, history.state]);

  // Layout > Margins / Orientation / Size
  const handlePageSetup = useCallback(
    (change: PageSetupChange) => {
      if (!history.state || readOnly) return;
      handleDocumentChange({
        ...history.state,
        package: {
          ...history.state.package,
          document: {
            ...history.state.package.document,
            finalSectionProperties: {
              ...history.state.package.document.finalSectionProperties,
              ...change,
            },
          },
        },
      });
    },
    [history.state, readOnly, handleDocumentChange]
  );

  // View > One Page / Page Width
  const handleFitZoom = useCallback(
    (fit: 'page' | 'width') => {
      const container = scrollContainerRef.current;
      const sp = history.state?.package.document?.finalSectionProperties;
      if (!container) return;
      const pageWidthPx = ((sp?.pageWidth ?? 12240) / 1440) * 96;
      const pageHeightPx = ((sp?.pageHeight ?? 15840) / 1440) * 96;
      const balloonSpace = showBalloons ? 300 : 0;
      const byWidth = (container.clientWidth - 48 - balloonSpace) / pageWidthPx;
      const byHeight = (container.clientHeight - 48) / pageHeightPx;
      const zoom = fit === 'width' ? byWidth : Math.min(byWidth, byHeight);
      setState((prev) => ({
        ...prev,
        zoom: Math.min(2, Math.max(0.5, Math.round(zoom * 100) / 100)),
      }));
    },
    [history.state, showBalloons]
  );

  // ============================================================================
  // FIND/REPLACE HANDLERS
  // ============================================================================

  // Store the current find result for navigation
  const findResultRef = useRef<FindResult | null>(null);

  // Handle find operation
  const handleFind = useCallback(
    (searchText: string, options: FindOptions): FindResult | null => {
      if (!history.state || !searchText.trim()) {
        findResultRef.current = null;
        return null;
      }

      const matches = findInDocument(history.state, searchText, options);
      const result: FindResult = {
        matches,
        totalCount: matches.length,
        currentIndex: 0,
      };

      findResultRef.current = result;
      findReplace.setMatches(matches, 0);

      // Scroll to first match
      if (matches.length > 0 && containerRef.current) {
        scrollToMatch(containerRef.current, matches[0]);
      }

      return result;
    },
    [history.state, findReplace]
  );

  // Handle find next
  const handleFindNext = useCallback((): FindMatch | null => {
    if (!findResultRef.current || findResultRef.current.matches.length === 0) {
      return null;
    }

    const newIndex = findReplace.goToNextMatch();
    const match = findResultRef.current.matches[newIndex];

    // Scroll to the match
    if (match && containerRef.current) {
      scrollToMatch(containerRef.current, match);
    }

    return match || null;
  }, [findReplace]);

  // Handle find previous
  const handleFindPrevious = useCallback((): FindMatch | null => {
    if (!findResultRef.current || findResultRef.current.matches.length === 0) {
      return null;
    }

    const newIndex = findReplace.goToPreviousMatch();
    const match = findResultRef.current.matches[newIndex];

    // Scroll to the match
    if (match && containerRef.current) {
      scrollToMatch(containerRef.current, match);
    }

    return match || null;
  }, [findReplace]);

  // Handle replace current match
  const handleReplace = useCallback(
    (replaceText: string): boolean => {
      if (!history.state || !findResultRef.current || findResultRef.current.matches.length === 0) {
        return false;
      }

      const currentMatch = findResultRef.current.matches[findResultRef.current.currentIndex];
      if (!currentMatch) return false;

      // Execute replace command
      try {
        const newDoc = executeCommand(history.state, {
          type: 'replaceText',
          range: {
            start: {
              paragraphIndex: currentMatch.paragraphIndex,
              offset: currentMatch.startOffset,
            },
            end: {
              paragraphIndex: currentMatch.paragraphIndex,
              offset: currentMatch.endOffset,
            },
          },
          text: replaceText,
        });

        handleDocumentChange(newDoc);
        return true;
      } catch (error) {
        console.error('Replace failed:', error);
        return false;
      }
    },
    [history.state, handleDocumentChange]
  );

  // Handle replace all matches
  const handleReplaceAll = useCallback(
    (searchText: string, replaceText: string, options: FindOptions): number => {
      if (!history.state || !searchText.trim()) {
        return 0;
      }

      // Find all matches first
      const matches = findInDocument(history.state, searchText, options);
      if (matches.length === 0) return 0;

      // Replace from end to start to maintain correct indices
      let doc = history.state;
      const sortedMatches = [...matches].sort((a, b) => {
        if (a.paragraphIndex !== b.paragraphIndex) {
          return b.paragraphIndex - a.paragraphIndex;
        }
        return b.startOffset - a.startOffset;
      });

      for (const match of sortedMatches) {
        try {
          doc = executeCommand(doc, {
            type: 'replaceText',
            range: {
              start: {
                paragraphIndex: match.paragraphIndex,
                offset: match.startOffset,
              },
              end: {
                paragraphIndex: match.paragraphIndex,
                offset: match.endOffset,
              },
            },
            text: replaceText,
          });
        } catch (error) {
          console.error('Replace failed for match:', match, error);
        }
      }

      handleDocumentChange(doc);
      findResultRef.current = null;
      findReplace.setMatches([], 0);

      return matches.length;
    },
    [history.state, handleDocumentChange, findReplace]
  );

  // Expose ref methods
  useImperativeHandle(
    ref,
    () => ({
      getAgent: () => agentRef.current,
      getDocument: () => history.state,
      getEditorRef: () => pagedEditorRef.current,
      save: handleSave,
      setZoom: (zoom: number) => setState((prev) => ({ ...prev, zoom })),
      getZoom: () => state.zoom,
      focus: () => {
        pagedEditorRef.current?.focus();
      },
      getCurrentPage: () => state.currentPage,
      getTotalPages: () => state.totalPages,
      scrollToPage: (_pageNumber: number) => {
        // TODO: Implement page navigation in ProseMirror
      },
      openPrintPreview: handleDirectPrint,
      print: handleDirectPrint,
      getMode: () => review.mode,
      setMode: review.setMode,
    }),
    [
      history.state,
      state.zoom,
      state.currentPage,
      state.totalPages,
      handleSave,
      handleDirectPrint,
      review.mode,
      review.setMode,
    ]
  );

  // Get detected variables from document
  const detectedVariables = useMemo(() => {
    if (!history.state) return [];
    return extractVariableNames(history.state);
  }, [history.state]);

  // Get header and footer content from document
  const { headerContent, footerContent } = useMemo<{
    headerContent: HeaderFooter | null;
    footerContent: HeaderFooter | null;
  }>(() => {
    if (!history.state?.package) {
      return { headerContent: null, footerContent: null };
    }

    const pkg = history.state.package;
    const sectionProps = pkg.document?.finalSectionProperties;
    const headers = pkg.headers;
    const footers = pkg.footers;

    let header: HeaderFooter | null = null;
    let footer: HeaderFooter | null = null;

    // Get default header from section references
    if (headers && sectionProps?.headerReferences) {
      const defaultRef = sectionProps.headerReferences.find((r) => r.type === 'default');
      if (defaultRef?.rId) {
        header = headers.get(defaultRef.rId) ?? null;
      }
    }

    // Get default footer from section references
    if (footers && sectionProps?.footerReferences) {
      const defaultRef = sectionProps.footerReferences.find((r) => r.type === 'default');
      if (defaultRef?.rId) {
        footer = footers.get(defaultRef.rId) ?? null;
      }
    }

    return { headerContent: header, footerContent: footer };
  }, [history.state]);

  // Pages container for in-document header/footer editing
  const [pagesRoot, setPagesRoot] = useState<HTMLElement | null>(null);

  // Handle header/footer save — update document package with edited content
  const handleHeaderFooterSave = useCallback(
    (content: (import('../types/document').Paragraph | import('../types/document').Table)[]) => {
      if (!hfEditPosition || !history.state?.package) {
        setHfEditPosition(null);
        return;
      }

      const pkg = history.state.package;
      const sectionProps = pkg.document?.finalSectionProperties;
      const refs =
        hfEditPosition === 'header'
          ? sectionProps?.headerReferences
          : sectionProps?.footerReferences;
      const defaultRef = refs?.find((r) => r.type === 'default');
      const map = hfEditPosition === 'header' ? pkg.headers : pkg.footers;

      if (defaultRef?.rId && map) {
        const existing = map.get(defaultRef.rId);
        if (existing) {
          const updated: HeaderFooter = {
            ...existing,
            content,
            modified: true,
          };
          // Copy the map so earlier history entries (undo) keep their version
          const nextMap = new Map(map);
          nextMap.set(defaultRef.rId, updated);

          const newDoc: Document = {
            ...history.state,
            package: {
              ...pkg,
              [hfEditPosition === 'header' ? 'headers' : 'footers']: nextMap,
            },
          };
          history.push(newDoc);
        }
      }

      hfViewRef.current = null;
      setHfEditPosition(null);
    },
    [hfEditPosition, history]
  );

  // Insert > Header / Footer / Page Number (Word's built-in designs)
  const handleHeaderFooterCommand = useCallback(
    (kind: 'header' | 'footer', action: HeaderFooterAction) => {
      if (!history.state?.package || readOnly) return;

      // While editing a header/footer, page-number commands insert at the cursor
      // (Word-style) instead of replacing the whole part — any Insert → Page Number
      // item targets the band being edited.
      if (
        hfEditPosition &&
        hfViewRef.current &&
        (action === 'pageNumber' || action === 'pageXofY')
      ) {
        insertPageNumberInView(hfViewRef.current, action);
        return;
      }

      const pkg = history.state.package;
      const sectionProps = pkg.document?.finalSectionProperties ?? {};
      const refsKey = kind === 'header' ? 'headerReferences' : 'footerReferences';
      const mapKey = kind === 'header' ? 'headers' : 'footers';
      const refs = sectionProps[refsKey] ?? [];
      const existingRef = refs.find((r) => r.type === 'default');
      const map = new Map(pkg[mapKey] ?? []);

      if (action === 'edit') {
        if (existingRef && map.get(existingRef.rId)) {
          setHfEditPosition(kind);
        } else {
          handleHeaderFooterCommand(kind, 'blank');
        }
        return;
      }

      if (action === 'remove') {
        if (!existingRef) return;
        // Leaving edit mode if we're removing the part being edited
        if (hfEditPosition === kind) {
          hfViewRef.current = null;
          setHfEditPosition(null);
        }
        map.delete(existingRef.rId);
        history.push({
          ...history.state,
          package: {
            ...pkg,
            [mapKey]: map,
            document: {
              ...pkg.document,
              finalSectionProperties: {
                ...sectionProps,
                [refsKey]: refs.filter((r) => r !== existingRef),
              },
            },
          },
        });
        return;
      }

      const content = headerFooterDesign(action);
      let rId = existingRef?.rId;
      if (!rId) {
        // Unique relationship id not used by the package or other parts
        const used = new Set<string>([
          ...(pkg.headers?.keys() ?? []),
          ...(pkg.footers?.keys() ?? []),
          ...(pkg.relationships?.keys() ?? []),
        ]);
        let n = 1;
        while (used.has(`rIdDec${kind}${n}`)) n++;
        rId = `rIdDec${kind}${n}`;
      }
      map.set(rId, { type: kind, hdrFtrType: 'default', content, modified: true });
      history.push({
        ...history.state,
        package: {
          ...pkg,
          [mapKey]: map,
          document: {
            ...pkg.document,
            finalSectionProperties: {
              ...sectionProps,
              [refsKey]: existingRef ? refs : [...refs, { type: 'default', rId }],
            },
          },
        },
      });
      // Blank opens the in-document editor; page-number designs apply immediately
      // (user can double-click or Edit Header/Footer to format with the ribbon).
      if (action === 'blank') setHfEditPosition(kind);
    },
    [history, readOnly, hfEditPosition]
  );

  // Handle header/footer double-click — enter in-document edit mode
  const handleHeaderFooterDoubleClick = useCallback(
    (position: 'header' | 'footer') => {
      if (readOnly) return;
      const hf = position === 'header' ? headerContent : footerContent;
      if (hf) {
        setHfEditPosition(position);
        return;
      }
      // No header/footer yet — create a blank one, then edit in place
      handleHeaderFooterCommand(position, 'blank');
    },
    [headerContent, footerContent, readOnly, handleHeaderFooterCommand]
  );

  // Container styles - using overflow: auto so sticky toolbar works
  const containerStyle: CSSProperties = {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    width: '100%',
    backgroundColor: 'var(--doc-bg-subtle)',
    ...style,
  };

  const mainContentStyle: CSSProperties = {
    display: 'flex',
    flex: 1,
    minHeight: 0, // Allow flex item to shrink below content size
    flexDirection: variablePanelPosition === 'left' ? 'row-reverse' : 'row',
  };

  const editorContainerStyle: CSSProperties = {
    flex: 1,
    overflow: 'auto', // This is the scroll container - sticky toolbar will stick to this
    position: 'relative',
  };

  const variablePanelStyle: CSSProperties = {
    width: '300px',
    borderLeft: variablePanelPosition === 'right' ? '1px solid var(--doc-border)' : undefined,
    borderRight: variablePanelPosition === 'left' ? '1px solid var(--doc-border)' : undefined,
    overflow: 'auto',
    backgroundColor: 'white',
  };

  // Render loading state
  if (state.isLoading) {
    return (
      <div
        className={`ep-root docx-editor docx-editor-loading ${className}`}
        style={containerStyle}
        data-testid="docx-editor"
      >
        {loadingIndicator || <DefaultLoadingIndicator />}
      </div>
    );
  }

  // Render error state
  if (state.parseError) {
    return (
      <div
        className={`ep-root docx-editor docx-editor-error ${className}`}
        style={containerStyle}
        data-testid="docx-editor"
      >
        <ParseError message={state.parseError} />
      </div>
    );
  }

  // Render placeholder when no document
  if (!history.state) {
    return (
      <div
        className={`ep-root docx-editor docx-editor-empty ${className}`}
        style={containerStyle}
        data-testid="docx-editor"
      >
        {placeholder || <DefaultPlaceholder />}
      </div>
    );
  }

  return (
    <ErrorProvider>
      <ErrorBoundary onError={handleEditorError}>
        <div
          ref={containerRef}
          className={`ep-root docx-editor ${className}`}
          style={containerStyle}
          data-testid="docx-editor"
        >
          {/* Ribbon (Word-style tabs) */}
          {showToolbar && (
            <Ribbon
              home={
                <Toolbar
                  variant="ribbon"
                  currentFormatting={state.selectionFormatting}
                  onFormat={handleFormat}
                  onUndo={undoActiveEditor}
                  onRedo={redoActiveEditor}
                  canUndo={true}
                  canRedo={true}
                  disabled={readOnly}
                  documentStyles={history.state?.package.styles?.styles}
                  theme={history.state?.package.theme || theme}
                  onRefocusEditor={focusActiveEditor}
                  onFind={() => findReplace.openFind()}
                  onReplace={() => findReplace.openReplace()}
                >
                  {collaborationSession && (
                    <CollaborationPresence
                      user={collaboration?.user ?? collaborationSession.user}
                      peers={collabPeers}
                      connected={collabConnected}
                      roomId={collaborationSession.roomId}
                    />
                  )}
                  {toolbarExtra}
                </Toolbar>
              }
              documentTitle={documentName}
              mode={review.mode}
              availableModes={modeOptions}
              onModeChange={review.setMode}
              onNew={handleNewDocument}
              onOpen={handleOpenDocument}
              onSave={() => void handleSaveCommand()}
              hasUnsavedChanges={hasUnsavedChanges}
              onDownload={handleDownload}
              onPrint={(withMarkup) => handlePrint(withMarkup)}
              stats={{ ...review.stats, pages: state.totalPages }}
              onInsertTable={handleInsertTable}
              onInsertImage={handleInsertImageClick}
              onInsertLink={() => handleFormat('insertLink')}
              onInsertPageBreak={handleInsertPageBreak}
              hasHeader={headerContent != null}
              hasFooter={footerContent != null}
              onHeaderFooter={handleHeaderFooterCommand}
              layout={ribbonLayout}
              onLayoutChange={setRibbonLayout}
              sectionProperties={history.state?.package.document?.finalSectionProperties}
              onPageSetup={handlePageSetup}
              canComment={(review.hasSelection || hasSavedSelection) && !readOnly}
              onNewComment={handleNewComment}
              onDeleteComment={review.deleteActiveComment}
              onDeleteAllComments={review.deleteAllComments}
              onPreviousComment={() => review.goToComment(-1)}
              onNextComment={() => review.goToComment(1)}
              commentCount={review.openCommentCount}
              showComments={review.showComments}
              onShowCommentsChange={review.setShowComments}
              markupView={review.markupView}
              onMarkupViewChange={review.setMarkupView}
              reviewingPaneOpen={review.sidebar.open}
              onReviewingPaneChange={review.sidebar.setOpen}
              onAccept={review.acceptAtSelection}
              onAcceptAll={review.acceptAll}
              onReject={review.rejectAtSelection}
              onRejectAll={review.rejectAll}
              onPreviousChange={() => review.goToRevision(-1)}
              onNextChange={() => review.goToRevision(1)}
              revisionCount={review.revisions.length}
              onWordCount={() => setWordCountOpen(true)}
              showRuler={rulerVisible}
              onShowRulerChange={setRulerVisible}
              zoom={state.zoom}
              onZoomChange={handleZoomChange}
              onOnePage={() => handleFitZoom('page')}
              onPageWidth={() => handleFitZoom('width')}
              onInsertContentControl={review.insertControlNow}
              canEditControlProperties={review.controlAtSelection != null && !readOnly}
              onContentControlProperties={review.openTagProperties}
              onShowTagsPane={() => review.openPane('tags')}
              tableContext={state.pmTableContext}
              onTableAction={handleTableAction}
              imageContext={state.pmImageContext}
              onImageWrapType={handleImageWrapType}
              onImageTransform={handleImageTransform}
              onOpenImagePosition={handleOpenImagePosition}
              onOpenImageProperties={handleOpenImageProperties}
            />
          )}

          {/* Main content area */}
          <div style={mainContentStyle}>
            {/* Reviewing Pane: comments, suggestions, Word tags */}
            {review.sidebar.open && <ReviewSidebar {...review.sidebar.props} />}

            {/* Editor container - this is the scroll container */}
            <div
              ref={scrollContainerRef}
              className={showBalloons ? 'ep-markup-area' : undefined}
              style={editorContainerStyle}
            >
              {/* Horizontal Ruler - sticky at top of the scroll container */}
              {rulerVisible && !readOnly && (
                <div className="ep-ruler-row sticky top-0 z-40 flex justify-center px-5 py-1 overflow-x-auto flex-shrink-0 bg-doc-bg">
                  <HorizontalRuler
                    sectionProps={history.state?.package.document?.finalSectionProperties}
                    zoom={state.zoom}
                    unit={rulerUnit}
                    editable={!readOnly}
                    onLeftMarginChange={handleLeftMarginChange}
                    onRightMarginChange={handleRightMarginChange}
                  />
                </div>
              )}

              {/* Vertical Ruler - fixed on left edge (hidden in read-only mode) */}
              {rulerVisible && !readOnly && (
                <div
                  style={{
                    position: 'absolute',
                    left: 0,
                    top: 0,
                    paddingTop: 20,
                    zIndex: 10,
                  }}
                >
                  <VerticalRuler
                    sectionProps={history.state?.package.document?.finalSectionProperties}
                    zoom={state.zoom}
                    unit={rulerUnit}
                    editable={!readOnly}
                    onTopMarginChange={handleTopMarginChange}
                    onBottomMarginChange={handleBottomMarginChange}
                  />
                </div>
              )}

              {/* Editor content area */}
              {review.commentHighlightCss && <style>{review.commentHighlightCss}</style>}
              <div
                className={`docx-markup-${review.markupView}`}
                style={{ position: 'relative' }}
                onMouseDown={(e) => {
                  // Focus editor when clicking on the background area (not the editor itself)
                  // Using mouseDown for immediate response before focus can be lost
                  if (e.target === e.currentTarget) {
                    e.preventDefault();
                    pagedEditorRef.current?.focus();
                  }
                }}
              >
                <PagedEditor
                  key={documentKey}
                  ref={pagedEditorRef}
                  document={history.state}
                  styles={history.state?.package.styles}
                  theme={history.state?.package.theme || theme}
                  sectionProperties={history.state?.package.document?.finalSectionProperties}
                  headerContent={headerContent}
                  footerContent={footerContent}
                  onHeaderFooterDoubleClick={handleHeaderFooterDoubleClick}
                  zoom={state.zoom}
                  readOnly={readOnly}
                  templateTagChips={templateTagChips}
                  extensionManager={extensionManager}
                  collaborationSession={activeCollabSession}
                  onDocumentChange={handleDocumentChange}
                  onSelectionChange={(_from, _to) => {
                    // Extract full selection state from PM and use the standard handler
                    const view = pagedEditorRef.current?.getView();
                    review.syncFromView(view ?? null);
                    if (view) {
                      const selectionState = extractSelectionState(view.state);
                      handleSelectionChange(selectionState);
                    } else {
                      handleSelectionChange(null);
                    }
                  }}
                  onNewComment={readOnly || review.mode === 'viewing' ? undefined : handleNewComment}
                  hideSelectionCommentButton={!!review.draft || review.mode === 'viewing'}
                  externalPlugins={externalPlugins}
                  onReady={(ref) => {
                    review.syncFromView(ref.getView());
                    onEditorViewReady?.(ref.getView()!);
                    setPagesRoot(
                      (containerRef.current?.querySelector(
                        '.paged-editor__pages'
                      ) as HTMLElement | null) ?? null
                    );
                  }}
                  onRenderedDomContextReady={onRenderedDomContextReady}
                  pluginOverlays={
                    showBalloons || pluginOverlays ? (
                      <>
                        {pluginOverlays}
                        {showBalloons && (
                          <CommentBalloons
                            comments={review.liveComments}
                            anchors={review.anchors}
                            activeCommentId={review.activeCommentId}
                            readOnly={readOnly}
                            markupView={review.markupView}
                            draft={review.draft}
                            onSubmitDraft={review.submitDraft}
                            onCancelDraft={review.cancelDraft}
                            onActivate={review.setActiveCommentId}
                            onSelect={review.selectComment}
                            onReply={review.reply}
                            onResolve={review.resolve}
                            onDelete={review.deleteComment}
                          />
                        )}
                      </>
                    ) : undefined
                  }
                />

                {/* Page navigation / indicator (the status bar shows the page when enabled) */}
                {showPageNumbers &&
                  !showStatusBar &&
                  state.totalPages > 0 &&
                  (enablePageNavigation ? (
                    <PageNavigator
                      currentPage={state.currentPage}
                      totalPages={state.totalPages}
                      onNavigate={handlePageNavigate}
                      position={pageNumberPosition as PageNavigatorPosition}
                      variant={pageNumberVariant as PageNavigatorVariant}
                      floating
                    />
                  ) : (
                    <PageNumberIndicator
                      currentPage={state.currentPage}
                      totalPages={state.totalPages}
                      position={pageNumberPosition as PageIndicatorPosition}
                      variant={pageNumberVariant as PageIndicatorVariant}
                      floating
                    />
                  ))}
              </div>
            </div>

            {/* Variable panel (hidden in read-only mode) */}
            {showVariablePanel && !readOnly && detectedVariables.length > 0 && (
              <div style={variablePanelStyle}>
                <VariablePanel
                  variables={detectedVariables}
                  values={state.variableValues}
                  onValuesChange={handleVariableValuesChange}
                  onApply={handleApplyVariables}
                  isApplying={state.isApplyingVariables}
                  descriptions={variableDescriptions}
                  disabled={readOnly}
                />
              </div>
            )}
          </div>

          {/* Status bar */}
          {showStatusBar && (
            <StatusBar
              currentPage={state.currentPage}
              totalPages={state.totalPages}
              wordCount={review.stats.words}
              mode={review.mode}
              onToggleTrackChanges={
                modeOptions.includes('suggesting') ? toggleTrackChanges : undefined
              }
              onShowWordCount={() => setWordCountOpen(true)}
              onReadMode={
                modeOptions.includes('viewing') ? () => review.setMode('viewing') : undefined
              }
              onPrintLayout={
                modeOptions.some((m) => m !== 'viewing')
                  ? () =>
                      review.mode === 'viewing' &&
                      review.setMode(modeOptions.find((m) => m !== 'viewing') ?? 'viewing')
                  : undefined
              }
              zoom={state.zoom}
              onZoomChange={handleZoomChange}
              showZoom={showZoomControl}
            />
          )}

          {wordCountOpen && (
            <RibbonDialog
              title="Word Count"
              onClose={() => setWordCountOpen(false)}
              testId="word-count-dialog"
            >
              <dl className="ep-dialog__stats">
                <dt>Pages</dt>
                <dd>{Math.max(1, state.totalPages).toLocaleString()}</dd>
                <dt>Words</dt>
                <dd data-testid="word-count-words">{review.stats.words.toLocaleString()}</dd>
                <dt>Characters (no spaces)</dt>
                <dd>{review.stats.charactersNoSpaces.toLocaleString()}</dd>
                <dt>Characters (with spaces)</dt>
                <dd>{review.stats.characters.toLocaleString()}</dd>
                <dt>Paragraphs</dt>
                <dd>{review.stats.paragraphs.toLocaleString()}</dd>
              </dl>
            </RibbonDialog>
          )}

          <input
            ref={openInputRef}
            type="file"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            onChange={handleOpenFileChange}
            style={{ display: 'none' }}
            data-testid="file-open-input"
          />

          {pendingDiscard && (
            <RibbonDialog
              title="Discard changes?"
              submitLabel="Discard"
              onClose={() => setPendingDiscard(null)}
              onSubmit={() => {
                const action = pendingDiscard;
                setPendingDiscard(null);
                action();
              }}
              testId="discard-dialog"
            >
              <p style={{ margin: 0 }}>
                {documentName} has unsaved changes. Save it first (File › Save) or discard the
                changes to continue.
              </p>
            </RibbonDialog>
          )}

          {review.tagDialog && (
            <ContentControlPropertiesDialog
              initial={review.tagDialog}
              onClose={review.closeTagDialog}
              onSave={review.saveTagDialog}
            />
          )}

          {/* Find/Replace Dialog */}
          <FindReplaceDialog
            isOpen={findReplace.state.isOpen}
            onClose={findReplace.close}
            onFind={handleFind}
            onFindNext={handleFindNext}
            onFindPrevious={handleFindPrevious}
            onReplace={handleReplace}
            onReplaceAll={handleReplaceAll}
            initialSearchText={findReplace.state.searchText}
            replaceMode={findReplace.state.replaceMode}
            currentResult={findResultRef.current}
          />

          {/* Hyperlink Dialog */}
          <HyperlinkDialog
            isOpen={hyperlinkDialog.state.isOpen}
            onClose={hyperlinkDialog.close}
            onSubmit={handleHyperlinkSubmit}
            onRemove={hyperlinkDialog.state.isEditing ? handleHyperlinkRemove : undefined}
            initialData={hyperlinkDialog.state.initialData}
            selectedText={hyperlinkDialog.state.selectedText}
            isEditing={hyperlinkDialog.state.isEditing}
          />
          <TablePropertiesDialog
            isOpen={tablePropsOpen}
            onClose={() => setTablePropsOpen(false)}
            onApply={(props) => {
              const view = getActiveEditorView();
              if (view) {
                setTableProperties(props)(view.state, view.dispatch);
              }
            }}
            currentProps={state.pmTableContext?.table?.attrs as Record<string, unknown> | undefined}
          />
          <ImagePositionDialog
            isOpen={imagePositionOpen}
            onClose={() => setImagePositionOpen(false)}
            onApply={handleApplyImagePosition}
          />
          <ImagePropertiesDialog
            isOpen={imagePropsOpen}
            onClose={() => setImagePropsOpen(false)}
            onApply={handleApplyImageProperties}
            currentData={
              state.pmImageContext
                ? {
                    alt: state.pmImageContext.alt ?? undefined,
                    borderWidth: state.pmImageContext.borderWidth ?? undefined,
                    borderColor: state.pmImageContext.borderColor ?? undefined,
                    borderStyle: state.pmImageContext.borderStyle ?? undefined,
                  }
                : undefined
            }
          />
          <FootnotePropertiesDialog
            isOpen={footnotePropsOpen}
            onClose={() => setFootnotePropsOpen(false)}
            onApply={handleApplyFootnoteProperties}
            footnotePr={history.state?.package.document?.finalSectionProperties?.footnotePr}
            endnotePr={history.state?.package.document?.finalSectionProperties?.endnotePr}
          />
          {/* Header/Footer in-document editor — ribbon formats this view */}
          {hfEditPosition && (
            <HeaderFooterEditor
              headerFooter={
                (hfEditPosition === 'header' ? headerContent : footerContent) ?? {
                  type: hfEditPosition,
                  hdrFtrType: 'default',
                  content: [{ type: 'paragraph', content: [] }],
                }
              }
              position={hfEditPosition}
              styles={history.state?.package.styles}
              pagesRoot={pagesRoot}
              onSave={handleHeaderFooterSave}
              onClose={() => {
                hfViewRef.current = null;
                setHfEditPosition(null);
              }}
              onViewReady={(view) => {
                hfViewRef.current = view;
              }}
              onSelectionChange={(view) => {
                const selectionState = extractSelectionState(view.state);
                handleSelectionChange(selectionState);
              }}
            />
          )}
          {/* Hidden file input for image insertion */}
          <input
            ref={imageInputRef}
            type="file"
            accept="image/*"
            style={{ display: 'none' }}
            onChange={handleImageFileChange}
          />
        </div>
      </ErrorBoundary>
    </ErrorProvider>
  );
});

// ============================================================================
// EXPORTS
// ============================================================================

export default DocxEditor;
