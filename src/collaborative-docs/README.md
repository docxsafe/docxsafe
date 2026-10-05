# Collaborative Docs Library

Host-agnostic Yjs + **WebSocket** collaboration (not WebRTC).

## Quick start (this repo)

```bash
# Terminal 1 — collaboration server (persists to .collab-data/)
npm run collaboration:server

# Terminal 2 — editor demo
npm run dev
```

Open the demo in two **different Chrome profiles** (or browsers), set the same room id, enable Collaborate.

## Packages

| Import                                      | Role                             |
| ------------------------------------------- | -------------------------------- |
| `docxsafe-editor/collaborative-docs`        | Browser client API               |
| `docxsafe-editor/collaborative-docs/server` | Node server + storage            |
| `docxsafe-editor/collaboration`             | Editor bridge (DocxEditor props) |

## Host integration

```ts
import { createCollaborativeDocs } from 'docxsafe-editor/collaborative-docs';

const collab = createCollaborativeDocs({
  websocket: { url: 'wss://collab.example.com' },
  auth: {
    getToken: () => getCurrentUserToken(),
    getCurrentUser: () => getCurrentUser(),
  },
  authorization: {
    canRead: ({ userId, matterId, documentId }) => hostCanRead(...),
    canWrite: ({ userId, matterId, documentId }) => hostCanWrite(...),
  },
});

const session = await collab.openDocument({
  tenantId: 'firm-123',
  matterId: 'M123',
  documentId: 'D456',
});
```

See `examples/docxsafe/integration.example.ts`.

## Room IDs

```text
tenant:{tenantId}:matter:{matterId}:document:{documentId}
```

Room IDs are not credentials — the server authenticates the token and authorizes before join.
