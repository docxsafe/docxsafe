# Decidendi Editor

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
npm install decidendi-editor
```

## Quick Start

```tsx
import { useRef } from 'react';
import { DocxEditor, type DocxEditorRef } from 'decidendi-editor';
import 'decidendi-editor/styles.css';

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

> **Next.js / SSR:** The editor requires the DOM. Use a dynamic import or lazy `useEffect` load to avoid server-side rendering issues.

## Props

| Prop                | Type                            | Default | Description                                 |
| ------------------- | ------------------------------- | ------- | ------------------------------------------- |
| `documentBuffer`    | `ArrayBuffer`                   | —       | `.docx` file contents to load               |
| `document`          | `Document`                      | —       | Pre-parsed document (alternative to buffer) |
| `readOnly`          | `boolean`                       | `false` | Read-only preview (no caret/selection)      |
| `showToolbar`       | `boolean`                       | `true`  | Show formatting toolbar                     |
| `showRuler`         | `boolean`                       | `false` | Show horizontal ruler                       |
| `showZoomControl`   | `boolean`                       | `true`  | Show zoom controls                          |
| `showVariablePanel` | `boolean`                       | `true`  | Show template variable panel                |
| `initialZoom`       | `number`                        | `1.0`   | Initial zoom level                          |
| `onChange`          | `(doc: Document) => void`       | —       | Called on document change                   |
| `onSave`            | `(buffer: ArrayBuffer) => void` | —       | Called on save                              |
| `onError`           | `(error: Error) => void`        | —       | Called on error                             |

## Ref Methods

```tsx
const ref = useRef<DocxEditorRef>(null);

await ref.current.save(); // Returns ArrayBuffer of the .docx
ref.current.getDocument(); // Current document object
ref.current.setZoom(1.5); // Set zoom to 150%
ref.current.focus(); // Focus the editor
ref.current.scrollToPage(3); // Scroll to page 3
ref.current.print(); // Print the document
```

## Read-Only Preview

Use `readOnly` for a preview-only viewer. This disables editing, caret, and selection UI.

```tsx
<DocxEditor documentBuffer={file} readOnly />
```

## Plugins

Extend the editor with the plugin system. Wrap `DocxEditor` in a `PluginHost` and pass plugins that can contribute ProseMirror plugins, side panels, document overlays, and custom CSS:

```tsx
import { DocxEditor, PluginHost, templatePlugin } from 'decidendi-editor';

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
