/**
 * The editor webview's page. It only loads the editor bundle, which creates
 * its own root and asks the extension for the document.
 *
 * Content Security Policy: the bundle, its stylesheet, and its fonts load
 * from the extension (`cspSource`); the bundle runs no worker, fetches
 * nothing, and evaluates no strings. Images and fonts the document carries
 * arrive as `data:` and `blob:` URLs. Styles allow inline CSS because the
 * laid-out pages position text with `style` attributes; a nonce cannot admit
 * an attribute, and CSS runs no code.
 */

import { randomBytes } from "node:crypto";

export const createNonce = (): string => randomBytes(18).toString("base64");

export const editorContentSecurityPolicy = (nonce: string, cspSource: string): string =>
  [
    "default-src 'none'",
    `script-src 'nonce-${nonce}' ${cspSource}`,
    `style-src ${cspSource} 'unsafe-inline'`,
    `font-src ${cspSource} data: blob:`,
    `img-src ${cspSource} data: blob:`,
    "worker-src 'none'",
    "connect-src 'none'",
  ].join("; ");

const escapeAttribute = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");

/**
 * The smoke test's hand on the keyboard, in the page only under
 * `FOLIO_VSCODE_TEST=1`: `{ type: "folio-test-type", text }` types `text` at
 * the end of the first paragraph with text, through the editor's own input
 * handling (the browser's insert-text command on the focused editor).
 */
export const EDITOR_TEST_SCRIPT = `
window.addEventListener("message", (event) => {
  const data = event.data;
  if (typeof data !== "object" || data === null || data.type !== "folio-test-type") return;
  const editable =
    document.querySelector(".paged-editor__hidden-pm [contenteditable='true']") ??
    document.querySelector("[contenteditable='true']");
  if (editable === null) return;
  editable.focus();
  const block = [...editable.querySelectorAll("p")].find((node) => node.textContent.trim() !== "");
  if (block !== undefined) {
    const range = document.createRange();
    range.selectNodeContents(block);
    range.collapse(false);
    const selection = window.getSelection();
    selection.removeAllRanges();
    selection.addRange(range);
  }
  setTimeout(() => document.execCommand("insertText", false, String(data.text)), 100);
});
`;

export type EditorShellOptions = {
  readonly nonce: string;
  /** `webview.cspSource`. */
  readonly cspSource: string;
  /** `webview.asWebviewUri` of `dist/editor/editor.js` and `dist/editor/editor.css`. */
  readonly scriptUri: string;
  readonly styleUri: string;
  readonly fileName: string;
  /** Add {@link EDITOR_TEST_SCRIPT}; only the smoke test sets it. */
  readonly testMode?: boolean;
};

export const editorShellHtml = ({
  nonce,
  cspSource,
  scriptUri,
  styleUri,
  fileName,
  testMode = false,
}: EditorShellOptions): string =>
  [
    "<!doctype html>",
    '<html lang="en"><head><meta charset="utf-8">',
    `<meta http-equiv="Content-Security-Policy" content="${escapeAttribute(editorContentSecurityPolicy(nonce, cspSource))}">`,
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    `<title>${escapeAttribute(fileName)}</title>`,
    `<link rel="stylesheet" href="${escapeAttribute(styleUri)}">`,
    "</head><body>",
    ...(testMode
      ? [`<script nonce="${escapeAttribute(nonce)}">${EDITOR_TEST_SCRIPT}</script>`]
      : []),
    `<script nonce="${escapeAttribute(nonce)}" src="${escapeAttribute(scriptUri)}"></script>`,
    "</body></html>",
  ].join("\n");
