/**
 * The webviews' pages. The preview's is a toolbar, a status line, and the
 * pages; the rendered document arrives later as a message, and this page only
 * hosts it. The editor's (at the end) loads the editor bundle.
 *
 * The preview's Content Security Policy: nothing loads from the network; the one script is
 * the extension's own, admitted by nonce; fonts and images are the `data:`
 * URLs the renderer inlines. Styles allow inline CSS because the rendered
 * pages position every glyph with `style` attributes. A nonce cannot admit an
 * attribute, and CSS runs no code.
 */

import { randomBytes } from "node:crypto";

export const createNonce = (): string => randomBytes(18).toString("base64");

/** The policy the preview's page carries. */
export const contentSecurityPolicy = (nonce: string): string =>
  [
    "default-src 'none'",
    `script-src 'nonce-${nonce}'`,
    "style-src 'unsafe-inline'",
    "font-src data:",
    "img-src data:",
  ].join("; ");

const escapeAttribute = (value: string): string =>
  value
    .replaceAll("&", "&amp;")
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");

/**
 * The shell's own rules. `#folio-pages` and `body.folio-preview` outrank the
 * `html, body` rules the rendered document brings, so the editor's theme
 * frames the pages.
 */
const SHELL_STYLE = `
body.folio-preview {
  margin: 0;
  padding: 0;
  display: block;
  background: var(--vscode-editorWidget-background, var(--vscode-editor-background));
  color: var(--vscode-foreground);
  font-family: var(--vscode-font-family);
  font-size: var(--vscode-font-size);
}
#folio-bar {
  position: sticky;
  top: 0;
  z-index: 1;
  display: flex;
  gap: 12px;
  align-items: baseline;
  padding: 6px 12px;
  background: var(--vscode-editor-background);
  border-bottom: 1px solid var(--vscode-panel-border, transparent);
}
#folio-name { font-weight: 600; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
#folio-status { color: var(--vscode-descriptionForeground); white-space: nowrap; }
#folio-error { display: none; margin: 16px; padding: 12px 16px; border-left: 3px solid var(--vscode-editorError-foreground, #c00); background: var(--vscode-inputValidation-errorBackground, transparent); }
#folio-error.visible { display: block; }
#folio-error-message { margin: 0; white-space: pre-wrap; font-family: var(--vscode-editor-font-family, monospace); }
#folio-error-hint { margin: 8px 0 0; }
#folio-error-hint:empty { display: none; }
#folio-retry { margin-top: 12px; padding: 4px 12px; border: 0; color: var(--vscode-button-foreground); background: var(--vscode-button-background); cursor: pointer; }
#folio-retry:hover { background: var(--vscode-button-hoverBackground); }
#folio-pages { display: flex; flex-direction: column; align-items: center; gap: 16px; padding: 16px; overflow-x: auto; }
#folio-pages.stale { opacity: 0.6; }
#folio-pages .layout-page { flex: none; box-shadow: 0 1px 4px rgba(0, 0, 0, 0.25); }
`;

export type ShellOptions = {
  readonly nonce: string;
  /** `webview.asWebviewUri` of `dist/webview/main.js`. */
  readonly scriptUri: string;
  readonly fileName: string;
};

export const shellHtml = ({ nonce, scriptUri, fileName }: ShellOptions): string =>
  [
    "<!doctype html>",
    '<html lang="en"><head><meta charset="utf-8">',
    `<meta http-equiv="Content-Security-Policy" content="${escapeAttribute(contentSecurityPolicy(nonce))}">`,
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    `<title>${escapeAttribute(fileName)}</title>`,
    `<style>${SHELL_STYLE}</style>`,
    '<style id="folio-document-style"></style>',
    '</head><body class="folio-preview">',
    '<header id="folio-bar">',
    `<span id="folio-name">${escapeAttribute(fileName)}</span>`,
    '<span id="folio-status" role="status" aria-live="polite">Rendering…</span>',
    "</header>",
    '<section id="folio-error" role="alert">',
    '<pre id="folio-error-message"></pre>',
    '<p id="folio-error-hint"></p>',
    '<button id="folio-retry" type="button">Try again</button>',
    "</section>",
    '<main id="folio-pages"></main>',
    `<script nonce="${escapeAttribute(nonce)}" src="${escapeAttribute(scriptUri)}"></script>`,
    "</body></html>",
  ].join("\n");

/**
 * The editor webview's policy. The editor bundle, its stylesheet, and its
 * fonts load from the extension (`cspSource`); the bundle runs no worker,
 * fetches nothing, and evaluates no strings. Images and fonts the document
 * carries arrive as `data:` and `blob:` URLs.
 */
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

export type EditorShellOptions = {
  readonly nonce: string;
  /** `webview.cspSource`. */
  readonly cspSource: string;
  /** `webview.asWebviewUri` of `dist/editor/editor.js` and `dist/editor/editor.css`. */
  readonly scriptUri: string;
  readonly styleUri: string;
  readonly fileName: string;
};

/** The editor webview's page: the bundle creates its own root and asks for the document. */
export const editorShellHtml = ({
  nonce,
  cspSource,
  scriptUri,
  styleUri,
  fileName,
}: EditorShellOptions): string =>
  [
    "<!doctype html>",
    '<html lang="en"><head><meta charset="utf-8">',
    `<meta http-equiv="Content-Security-Policy" content="${escapeAttribute(editorContentSecurityPolicy(nonce, cspSource))}">`,
    '<meta name="viewport" content="width=device-width, initial-scale=1">',
    `<title>${escapeAttribute(fileName)}</title>`,
    `<link rel="stylesheet" href="${escapeAttribute(styleUri)}">`,
    "</head><body>",
    `<script nonce="${escapeAttribute(nonce)}" src="${escapeAttribute(scriptUri)}"></script>`,
    "</body></html>",
  ].join("\n");
