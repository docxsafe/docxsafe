/**
 * DocXSafeEditor — drop-in React component for any app.
 *
 * Wraps {@link DocxEditor} with a simpler API:
 * - `source` accepts File / Blob / ArrayBuffer / URL / parsed Document
 * - optional `collaboration` object (no need to call useCollaboration yourself)
 * - fills its parent by default (`height: '100%'`)
 *
 * @example
 * ```tsx
 * import { DocXSafeEditor } from 'docxsafe-editor';
 * import 'docxsafe-editor/styles.css';
 *
 * export function App() {
 *   return (
 *     <div style={{ height: '100vh' }}>
 *       <DocXSafeEditor
 *         source="/contracts/msa.docx"
 *         author="Alice"
 *         onSave={async (buffer) => {
 *           await fetch('/api/docs/1', { method: 'PUT', body: buffer });
 *         }}
 *       />
 *     </div>
 *   );
 * }
 * ```
 */

import {
  forwardRef,
  useCallback,
  useEffect,
  useImperativeHandle,
  useMemo,
  useRef,
  useState,
  type CSSProperties,
} from 'react';
import type { Document } from '../types/document';
import { useCollaboration, colorForUser, randomUserId } from '../collaboration';
import type { CollaborationUser } from '../collaboration/types';
import { type DocxInput, toArrayBuffer } from '../utils/docxInput';
import { DocxEditor, type DocxEditorProps, type DocxEditorRef } from './DocxEditor';

// ============================================================================
// TYPES
// ============================================================================

/** Anything the drop-in editor can open */
export type DocXSafeEditorSource = DocxInput | Document | string | null;

export interface DocXSafeCollaborationProps {
  /** Shared room id (peers with the same id edit together) */
  roomId: string;
  /** Collaboration WebSocket URL, e.g. `ws://localhost:1234` or `/collab-ws` */
  websocketUrl: string;
  /** Local user shown on carets / comments */
  user: {
    id?: string;
    name: string;
    color?: string;
  };
  /** Optional auth token forwarded to the collab server */
  token?: string;
}

export interface DocXSafeEditorProps extends Omit<
  DocxEditorProps,
  'document' | 'documentBuffer' | 'collaboration' | 'style' | 'className'
> {
  /**
   * Document to open.
   * - `File` / `Blob` / `ArrayBuffer` / `Uint8Array` — binary .docx
   * - `string` — URL fetched as ArrayBuffer
   * - parsed `Document` — skip parsing
   * - `null` / omit — blank document (user can File → Open)
   */
  source?: DocXSafeEditorSource;
  /** CSS height of the editor shell (default: `'100%'`) */
  height?: string | number;
  /** Extra class on the outer shell */
  className?: string;
  /** Extra styles on the outer shell */
  style?: CSSProperties;
  /**
   * Optional real-time collaboration. When set, peers in the same `roomId`
   * share edits, carets, comments, and Word tags.
   */
  collaboration?: DocXSafeCollaborationProps | null;
  /** Called once the inner editor ref is ready */
  onReady?: (api: DocXSafeEditorHandle) => void;
}

/** Public handle — same as DocxEditorRef plus download helper */
export interface DocXSafeEditorHandle extends DocxEditorRef {
  /** Trigger a browser download of the current .docx */
  download: (filename?: string) => Promise<void>;
}

export type DocXSafeEditorRef = DocXSafeEditorHandle;

// ============================================================================
// HELPERS
// ============================================================================

function isParsedDocument(value: unknown): value is Document {
  return (
    !!value &&
    typeof value === 'object' &&
    'package' in value &&
    !!(value as Document).package?.document
  );
}

async function resolveSource(
  source: DocXSafeEditorSource | undefined
): Promise<{ document?: Document | null; buffer?: ArrayBuffer | null; error?: Error }> {
  if (source == null) {
    return { document: null, buffer: null };
  }
  if (isParsedDocument(source)) {
    return { document: source, buffer: null };
  }
  if (typeof source === 'string') {
    try {
      const res = await fetch(source);
      if (!res.ok) {
        throw new Error(`Failed to fetch document (${res.status}): ${source}`);
      }
      return { buffer: await res.arrayBuffer(), document: null };
    } catch (err) {
      return {
        error: err instanceof Error ? err : new Error(String(err)),
      };
    }
  }
  try {
    return { buffer: await toArrayBuffer(source), document: null };
  } catch (err) {
    return {
      error: err instanceof Error ? err : new Error(String(err)),
    };
  }
}

function downloadBuffer(buffer: ArrayBuffer, filename: string): void {
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = /\.docx$/i.test(filename) ? filename : `${filename}.docx`;
  a.click();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ============================================================================
// COMPONENT
// ============================================================================

/**
 * Drop-in Word-style DOCX editor for React apps.
 *
 * Prefer this over wiring {@link DocxEditor} + collaboration hooks yourself.
 */
export const DocXSafeEditor = forwardRef<DocXSafeEditorHandle, DocXSafeEditorProps>(
  function DocXSafeEditor(props, ref) {
    const {
      source,
      height = '100%',
      className,
      style,
      collaboration: collabProps = null,
      onReady,
      onError,
      author = 'Anonymous',
      documentName = 'Document.docx',
      showRuler = true,
      showVariablePanel = false,
      defaultReviewSidebarOpen = false,
      ...editorProps
    } = props;

    const innerRef = useRef<DocxEditorRef>(null);
    const [document, setDocument] = useState<Document | null>(null);
    const [documentBuffer, setDocumentBuffer] = useState<ArrayBuffer | null>(null);
    const [loadError, setLoadError] = useState<string | null>(null);
    const [sourceKey, setSourceKey] = useState(0);

    // Resolve source → buffer / document whenever it changes
    useEffect(() => {
      let cancelled = false;
      setLoadError(null);
      void (async () => {
        const result = await resolveSource(source);
        if (cancelled) return;
        if (result.error) {
          setLoadError(result.error.message);
          setDocument(null);
          setDocumentBuffer(null);
          onError?.(result.error);
          return;
        }
        setDocument(result.document ?? null);
        setDocumentBuffer(result.buffer ?? null);
        setSourceKey((k) => k + 1);
      })();
      return () => {
        cancelled = true;
      };
    }, [source, onError]);

    // Collaboration shorthand
    const collabUser: CollaborationUser | null = useMemo(() => {
      if (!collabProps) return null;
      const id = collabProps.user.id?.trim() || randomUserId();
      return {
        id,
        name: collabProps.user.name.trim() || 'Guest',
        color: collabProps.user.color || colorForUser(id),
      };
    }, [collabProps]);

    const collab = useCollaboration(
      collabProps && collabUser
        ? {
            roomId: collabProps.roomId,
            websocketUrl: collabProps.websocketUrl,
            user: collabUser,
            token: collabProps.token,
          }
        : null
    );

    const download = useCallback(
      async (filename?: string) => {
        const buffer = await innerRef.current?.save();
        if (!buffer) return;
        downloadBuffer(buffer, filename ?? documentName);
      },
      [documentName]
    );

    const api = useMemo<DocXSafeEditorHandle>(
      () => ({
        getAgent: () => innerRef.current?.getAgent() ?? null,
        getDocument: () => innerRef.current?.getDocument() ?? null,
        getEditorRef: () => innerRef.current?.getEditorRef() ?? null,
        save: async () => (await innerRef.current?.save()) ?? null,
        setZoom: (zoom: number) => innerRef.current?.setZoom(zoom),
        getZoom: () => innerRef.current?.getZoom() ?? 1,
        focus: () => innerRef.current?.focus(),
        getCurrentPage: () => innerRef.current?.getCurrentPage() ?? 1,
        getTotalPages: () => innerRef.current?.getTotalPages() ?? 1,
        scrollToPage: (page: number) => innerRef.current?.scrollToPage(page),
        openPrintPreview: () => innerRef.current?.openPrintPreview(),
        print: () => innerRef.current?.print(),
        getMode: () => innerRef.current?.getMode() ?? 'editing',
        setMode: (mode) => innerRef.current?.setMode(mode),
        download,
      }),
      [download]
    );

    useImperativeHandle(ref, () => api, [api]);

    useEffect(() => {
      onReady?.(api);
      // Intentionally once when ready — host can keep the handle
      // eslint-disable-next-line react-hooks/exhaustive-deps
    }, []);

    const shellStyle: CSSProperties = {
      display: 'flex',
      flexDirection: 'column',
      height: typeof height === 'number' ? `${height}px` : height,
      minHeight: 320,
      width: '100%',
      overflow: 'hidden',
      background: '#f3f3f3',
      ...style,
    };

    if (loadError) {
      return (
        <div
          className={className}
          style={{
            ...shellStyle,
            alignItems: 'center',
            justifyContent: 'center',
            padding: 24,
            color: '#b91c1c',
            fontFamily: 'system-ui, sans-serif',
          }}
          role="alert"
        >
          {loadError}
        </div>
      );
    }

    return (
      <div className={className} style={shellStyle} data-testid="docxsafe-editor">
        <DocxEditor
          key={sourceKey}
          ref={innerRef}
          document={document}
          documentBuffer={documentBuffer}
          author={author}
          documentName={documentName}
          showRuler={showRuler}
          showVariablePanel={showVariablePanel}
          defaultReviewSidebarOpen={defaultReviewSidebarOpen}
          onError={onError}
          collaboration={
            collab.session ? { session: collab.session, user: collabUser ?? undefined } : null
          }
          style={{ flex: 1, minHeight: 0, height: '100%' }}
          {...editorProps}
        />
      </div>
    );
  }
);

DocXSafeEditor.displayName = 'DocXSafeEditor';

export default DocXSafeEditor;
