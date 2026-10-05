# DocXSafe Editor

Open-source WYSIWYG DOCX editor for React. Open, review, edit, and save `.docx` files entirely in the browser — no server required.

**Features**

- **Word editor** — full-fidelity editing of text, tables, images, lists, headers/footers
- **Editing modes** — Editing, Suggesting (tracked changes) and Viewing (read-only), switchable from the toolbar or via the `mode` prop
- **Suggestions** — every edit in Suggesting mode is recorded as a Word revision (`w:ins` / `w:del`) with author and date; accept or reject one by one or all at once
- **Comments** — add, reply, resolve and delete threaded comments anchored to text; saved to `word/comments.xml` so Word sees them
- **Word tags** — insert and manage Word content controls (`w:sdt` with Tag and Title), plus docxtemplater `{variable}` tags via the template plugin
- **Print / Save as PDF** — prints the paginated document; choose "Save as PDF" in the print dialog

## Installation

```bash
npm install docxsafe-editor
```

## Quick Start

```tsx
import { useRef } from 'react';
import { DocXSafeEditor, type DocXSafeEditorRef } from 'docxsafe-editor';
import 'docxsafe-editor/styles.css';

function App() {
  const editorRef = useRef<DocXSafeEditorRef>(null);

  return (
    <div style={{ height: '100vh' }}>
      <DocXSafeEditor
        ref={editorRef}
        source="/contracts/msa.docx" // File | Blob | ArrayBuffer | URL | omit for blank
        author="Alice"
        onSave={async (buffer) => {
          await fetch('/api/documents/1', { method: 'PUT', body: buffer });
        }}
      />
    </div>
  );
}
```

`DocXSafeEditor` is the drop-in component for host apps. It wraps the full Word-style UI (ribbon, comments, tags, print) and can optionally take a `collaboration` prop for real-time co-editing.

For lower-level control, use `DocxEditor` directly (same package).

> **Next.js / SSR:** The editor requires the DOM. Use a dynamic import with `ssr: false` (or load it inside `useEffect`) so it never runs on the server.

### With collaboration

```tsx
<DocXSafeEditor
  source={file}
  author="Alice"
  collaboration={{
    roomId: 'acme-msa',
    websocketUrl: 'ws://localhost:1234',
    user: { name: 'Alice' },
  }}
/>
```

### Ref helpers

```tsx
const ref = useRef<DocXSafeEditorRef>(null);

await ref.current?.save(); // ArrayBuffer of the .docx
await ref.current?.download('signed.docx');
ref.current?.print();
ref.current?.focus();
```

## Advanced: `DocxEditor`

```tsx
import { useRef } from 'react';
import { DocxEditor, type DocxEditorRef } from 'docxsafe-editor';
import 'docxsafe-editor/styles.css';

function Editor({ file }: { file: ArrayBuffer }) {
  const editorRef = useRef<DocxEditorRef>(null);

  const handleSave = async () => {
    const buffer = await editorRef.current?.save();
    if (buffer) {
      await fetch('/api/documents/1', { method: 'PUT', body: buffer });
    }
  };

  return (
    <>
      <button onClick={handleSave}>Save</button>
      <DocxEditor ref={editorRef} documentBuffer={file} onChange={() => {}} />
    </>
  );
}
```

## Props (`DocXSafeEditor`)

| Prop            | Type                                                | Default           | Description                              |
| --------------- | --------------------------------------------------- | ----------------- | ---------------------------------------- |
| `source`        | `File \| Blob \| ArrayBuffer \| string \| Document` | —                 | Document to open (URL string is fetched) |
| `author`        | `string`                                            | `'Anonymous'`     | Author for comments / track changes      |
| `documentName`  | `string`                                            | `'Document.docx'` | Name shown in File > Info / downloads    |
| `height`        | `string \| number`                                  | `'100%'`          | Outer shell height                       |
| `readOnly`      | `boolean`                                           | `false`           | Read-only preview                        |
| `collaboration` | `{ roomId, websocketUrl, user }`                    | —                 | Optional real-time co-editing            |
| `onSave`        | `(buffer: ArrayBuffer) => void`                     | —                 | Called on save                           |
| `onChange`      | `(doc: Document) => void`                           | —                 | Called on document change                |
| `onError`       | `(error: Error) => void`                            | —                 | Called on error                          |

All other `DocxEditor` props (zoom, modes, rulers, …) are also accepted.

## Read-Only Preview

Use `readOnly` for a preview-only viewer. This disables editing, caret, and selection UI.

```tsx
<DocXSafeEditor source={file} readOnly />
```

## Plugins

Extend the editor with the plugin system. Wrap `DocxEditor` in a `PluginHost` and pass plugins that can contribute ProseMirror plugins, side panels, document overlays, and custom CSS:

```tsx
import { DocxEditor, PluginHost, templatePlugin } from 'docxsafe-editor';

function Editor({ file }: { file: ArrayBuffer }) {
  return (
    <PluginHost plugins={[templatePlugin]}>
      <DocxEditor documentBuffer={file} />
    </PluginHost>
  );
}
```

| Plugin                                 | Description                                                                                  |
| -------------------------------------- | -------------------------------------------------------------------------------------------- |
| [Docxtemplater](src/plugins/template/) | Syntax highlighting and annotation panel for [docxtemplater](https://docxtemplater.com) tags |

See [docs/PLUGINS.md](docs/PLUGINS.md) for the full plugin API, including how to create custom plugins with panels, overlays, and ProseMirror integrations.

## Features

- Full WYSIWYG editing with Microsoft Word fidelity
- Text and paragraph formatting (bold, italic, fonts, colors, alignment, spacing)
- Tables, images, hyperlinks
- Extensible plugin architecture
- Undo/redo, find & replace, keyboard shortcuts
- Print preview
- Zero server dependencies

## Development

```bash
bun install
bun run dev
```

## License

MIT
