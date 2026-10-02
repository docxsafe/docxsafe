/**
 * DOCX Editor Demo
 *
 * Complete demo showing all features:
 * - Load sample or custom DOCX
 * - Full editing with toolbar
 * - Context menu AI (mock handler)
 * - Save/download
 * - Word tags (content controls) via Developer / Review
 */

import './styles.css';
import React, { useState, useRef, useCallback, useEffect, useMemo } from 'react';
import { createRoot } from 'react-dom/client';
import {
  DocxEditor,
  type DocxEditorRef,
  createEmptyDocument,
  type Document,
} from '../src/index';
import { useCollaboration, colorForUser } from '../src/collaboration';

// ============================================================================
// STYLES
// ============================================================================

const styles: Record<string, React.CSSProperties> = {
  container: {
    display: 'flex',
    flexDirection: 'column',
    height: '100vh',
    overflow: 'hidden',
    background: '#f8fafc',
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    padding: '12px 20px',
    background: '#fff',
    borderBottom: '1px solid #e2e8f0',
  },
  headerLeft: {
    display: 'flex',
    alignItems: 'center',
    gap: '16px',
  },
  headerRight: {
    display: 'flex',
    alignItems: 'center',
    gap: '8px',
  },
  title: {
    fontSize: '18px',
    fontWeight: 600,
    margin: 0,
    color: '#0f172a',
    letterSpacing: '-0.025em',
  },
  titleLink: {
    textDecoration: 'none',
    color: 'inherit',
  },
  fileName: {
    fontSize: '13px',
    color: '#64748b',
    padding: '4px 10px',
    background: '#f1f5f9',
    borderRadius: '6px',
  },
  status: {
    fontSize: '12px',
    color: '#64748b',
    padding: '4px 8px',
    background: '#f1f5f9',
    borderRadius: '4px',
  },
  main: {
    flex: 1,
    display: 'flex',
    overflow: 'hidden',
  },
};

// ============================================================================
// DEMO APP
// ============================================================================

function Demo() {
  const editorRef = useRef<DocxEditorRef>(null);
  // Expose the editor to e2e tests in development
  useEffect(() => {
    if (import.meta.env.DEV) {
      (window as unknown as { __docxEditorRef?: typeof editorRef }).__docxEditorRef = editorRef;
    }
  }, []);
  const [currentDocument] = useState<Document | null>(() => createEmptyDocument());
  // File > New / Open / Save live in the editor's ribbon (File tab)
  const [documentBuffer] = useState<ArrayBuffer | null>(null);
  const [fileName, setFileName] = useState<string>('Untitled.docx');
  const [status, setStatus] = useState<string>('');

  // Collaboration: File → Open a document, then turn Collaborate on.
  // Share the same room id in a second tab/window to co-edit.
  const [roomId, setRoomId] = useState('decidendi-demo-room');
  /** Room actually joined — only changes when Collaborate is toggled (or Rejoin). */
  const [activeRoom, setActiveRoom] = useState<string | null>(null);
  const [userName, setUserName] = useState(
    () => `User ${Math.floor(Math.random() * 90 + 10)}`
  );
  const collabUser = useMemo(
    () => ({ name: userName, color: colorForUser(userName) }),
    [userName]
  );
  const signaling = useMemo(() => {
    const proto = window.location.protocol === 'https:' ? 'wss' : 'ws';
    // Vite proxies `/signaling` → local `npm run signaling` (port 4444)
    return [
      `${proto}://${window.location.host}/signaling`,
      'ws://localhost:4444',
      'wss://y-webrtc-eu.fly.dev',
    ];
  }, []);

  const collab = useCollaboration(
    activeRoom
      ? {
          roomId: activeRoom,
          user: collabUser,
          signaling,
        }
      : null
  );

  const collabEnabled = activeRoom != null;
  const joinCollab = useCallback((enabled: boolean) => {
    if (enabled) {
      setActiveRoom(roomId.trim() || 'decidendi-demo-room');
    } else {
      setActiveRoom(null);
    }
  }, [roomId]);
  const rejoinRoom = useCallback(() => {
    setActiveRoom(roomId.trim() || 'decidendi-demo-room');
  }, [roomId]);

  const handleDocumentChange = useCallback((_doc: Document) => {
    // no-op (disable noisy logging)
  }, []);

  const handleError = useCallback((error: Error) => {
    console.error('Editor error:', error);
    setStatus(`Error: ${error.message}`);
  }, []);

  const handleFontsLoaded = useCallback(() => {
    console.log('Fonts loaded');
  }, []);

  return (
    <div style={styles.container}>
      <header style={styles.header}>
        <div style={styles.headerLeft}>
          <a
            href="https://github.com/decidendi/decidendi-editor"
            target="_blank"
            rel="noopener noreferrer"
            style={styles.titleLink}
          >
            <h1 style={styles.title}>Decidendi Editor</h1>
          </a>
          {fileName && <span style={styles.fileName}>{fileName}</span>}
        </div>
        <div style={styles.headerRight}>
          <input
            style={{ ...styles.fileName, width: 140 }}
            value={roomId}
            onChange={(e) => setRoomId(e.target.value)}
            onBlur={() => {
              if (collabEnabled && roomId.trim() && roomId.trim() !== activeRoom) {
                rejoinRoom();
              }
            }}
            onKeyDown={(e) => {
              if (e.key === 'Enter' && collabEnabled) rejoinRoom();
            }}
            placeholder="Room id"
            title="Set a shared room id, then enable Collaborate"
            data-testid="collab-room"
          />
          <input
            style={{ ...styles.fileName, width: 100 }}
            value={userName}
            onChange={(e) => setUserName(e.target.value)}
            placeholder="Your name"
            data-testid="collab-name"
          />
          <label
            style={{ ...styles.status, display: 'inline-flex', gap: 6, alignItems: 'center' }}
            title="Open a document first, then turn this on to share the room"
          >
            <input
              type="checkbox"
              checked={collabEnabled}
              onChange={(e) => joinCollab(e.target.checked)}
              data-testid="collab-toggle"
            />
            Collaborate
          </label>
          {status && <span style={styles.status}>{status}</span>}
        </div>
      </header>

      <main style={styles.main}>
        {/* Do not remount on collab toggle — that wiped File > Open content */}
        <DocxEditor
          ref={editorRef}
          document={documentBuffer ? undefined : currentDocument}
          documentBuffer={documentBuffer}
          onChange={handleDocumentChange}
          onError={handleError}
          onFontsLoaded={handleFontsLoaded}
          showToolbar={true}
          showRuler={true}
          showVariablePanel={true}
          showZoomControl={true}
          showPageNumbers={false}
          initialZoom={1.0}
          variablePanelPosition="right"
          author={userName || 'Demo User'}
          documentName={fileName}
          onNew={() => setFileName('Document1.docx')}
          onOpen={(file) => setFileName(file.name)}
          collaboration={
            collab.session
              ? { session: collab.session, user: collab.session.user }
              : undefined
          }
        />
      </main>
    </div>
  );
}

// ============================================================================
// MOUNT
// ============================================================================

const container = document.getElementById('app');
if (container) {
  const root = createRoot(container);
  root.render(<Demo />);
}
