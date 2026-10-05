/**
 * Drop-in usage example for `DocXSafeEditor`.
 *
 * ```bash
 * npm install docxsafe-editor react react-dom
 * ```
 *
 * ```tsx
 * import { useRef } from 'react';
 * import { DocXSafeEditor, type DocXSafeEditorRef } from 'docxsafe-editor';
 * import 'docxsafe-editor/styles.css';
 *
 * export function MyApp() {
 *   const editorRef = useRef<DocXSafeEditorRef>(null);
 *
 *   return (
 *     <div style={{ height: '100vh' }}>
 *       <DocXSafeEditor
 *         ref={editorRef}
 *         // File | Blob | ArrayBuffer | URL string | omit for blank
 *         source="/docs/offer-letter.docx"
 *         author="Alice Wong"
 *         documentName="Offer Letter.docx"
 *         onSave={async (buffer) => {
 *           await fetch('/api/documents/123', { method: 'PUT', body: buffer });
 *         }}
 *         // Optional real-time collab (same roomId = same document)
 *         collaboration={{
 *           roomId: 'tenant:acme:matter:M1:document:D1',
 *           websocketUrl: 'wss://collab.example.com',
 *           user: { name: 'Alice Wong' },
 *         }}
 *       />
 *     </div>
 *   );
 * }
 *
 * // Later:
 * // const docx = await editorRef.current?.save();
 * // await editorRef.current?.download('signed.docx');
 * ```
 */

export {};
