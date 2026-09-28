/**
 * A page standing in for a VS Code webview: the webview's Content Security
 * Policy, a fake extension on the other side of `postMessage`, a subset of a
 * VS Code theme's variables, and the built bundle served from the webview's
 * resource origin.
 */

import type { Page } from "@playwright/test";
import { readFile } from "node:fs/promises";
import path from "node:path";

import type { EditorMessage, HostMessage } from "../src/protocol";

export const ORIGIN = "https://webview.test";
const NONCE = "folio-test-nonce";
const DIST = path.resolve(import.meta.dirname, "../dist/vscode");

/** What the extension's shell sets, with the webview's resource origin as `<cspSource>`. */
const CONTENT_SECURITY_POLICY = [
  "default-src 'none'",
  `script-src 'nonce-${NONCE}' ${ORIGIN}`,
  `style-src ${ORIGIN} 'unsafe-inline'`,
  `font-src ${ORIGIN} data: blob:`,
  `img-src ${ORIGIN} data: blob:`,
  "worker-src 'none'",
  "connect-src 'none'",
].join("; ");

/**
 * Stands in for VS Code: `acquireVsCodeApi` records what the webview posts,
 * the page records policy violations, and `Worker` is counted so a spawn
 * attempt shows even if it would throw.
 */
const FAKE_EXTENSION = `
window.__folio = { sent: [], violations: [], workers: 0 };
window.acquireVsCodeApi = () => ({
  postMessage: (message) => window.__folio.sent.push(message),
  getState: () => undefined,
  setState: () => undefined,
});
document.addEventListener("securitypolicyviolation", (event) => {
  window.__folio.violations.push(event.violatedDirective + " " + event.blockedURI);
});
const NativeWorker = window.Worker;
window.Worker = function (...args) {
  window.__folio.workers += 1;
  return new NativeWorker(...args);
};
`;

export type WebviewTheme = "light" | "dark";

/** A subset of the variables VS Code defines on a theme's `html`. */
const THEME_VARIABLES = {
  light: `
html {
  --vscode-font-family: system-ui, sans-serif;
  --vscode-font-size: 13px;
  --vscode-editor-background: #ffffff;
  --vscode-editor-foreground: #3b3b3b;
  --vscode-editorWidget-background: #f8f8f8;
  --vscode-button-background: #005fb8;
  --vscode-button-foreground: #ffffff;
  --vscode-descriptionForeground: #3b3b3b;
  --vscode-focusBorder: #005fb8;
}
`,
  dark: `
html {
  --vscode-font-family: system-ui, sans-serif;
  --vscode-font-size: 13px;
  --vscode-editor-background: #1f1f1f;
  --vscode-editor-foreground: #cccccc;
  --vscode-editorWidget-background: #202020;
  --vscode-button-background: #0078d4;
  --vscode-button-foreground: #ffffff;
  --vscode-descriptionForeground: #9d9d9d;
  --vscode-focusBorder: #0078d4;
  --vscode-panel-border: #2b2b2b;
  --vscode-widget-border: #313131;
  --vscode-toolbar-hoverBackground: #5a5d5e50;
}
`,
} as const satisfies Record<WebviewTheme, string>;

const shell = (theme: WebviewTheme) => `<!doctype html>
<html lang="en"><head><meta charset="utf-8">
<meta http-equiv="Content-Security-Policy" content="${CONTENT_SECURITY_POLICY}">
<link rel="stylesheet" href="${ORIGIN}/editor.css">
<style>${THEME_VARIABLES[theme]}</style>
<script nonce="${NONCE}">${FAKE_EXTENSION}</script>
</head><body class="vscode-${theme}">
<script nonce="${NONCE}" src="${ORIGIN}/editor.js"></script>
</body></html>`;

const CONTENT_TYPES: Record<string, string> = {
  ".css": "text/css",
  ".js": "text/javascript",
  ".woff2": "font/woff2",
};

export type FolioProbe = { sent: EditorMessage[]; violations: string[]; workers: number };

declare global {
  var __folio: FolioProbe;
}

/** Serve the shell and the built bundle from the webview origin, then open the shell. */
export const openWebview = async (page: Page, theme: WebviewTheme = "light"): Promise<void> => {
  await page.route(`${ORIGIN}/**`, async (route) => {
    const { pathname } = new URL(route.request().url());
    if (pathname === "/index.html") {
      await route.fulfill({ contentType: "text/html", body: shell(theme) });
      return;
    }
    const file = path.join(DIST, path.normalize(pathname));
    if (!file.startsWith(`${DIST}${path.sep}`)) {
      await route.fulfill({ status: 404 });
      return;
    }
    await route.fulfill({
      contentType: CONTENT_TYPES[path.extname(file)] ?? "application/octet-stream",
      body: await readFile(file),
    });
  });
  await page.goto(`${ORIGIN}/index.html`);
};

/** Wait until the webview has posted a message with every field of `expected`. */
export const waitForSent = (page: Page, expected: Record<string, string | number>) =>
  page.waitForFunction(
    (fields) =>
      globalThis.__folio.sent.some((message) =>
        Object.entries(fields).every(([key, value]) => Reflect.get(message, key) === value),
      ),
    expected,
  );

/** Post `load` or `reload`; the document bytes arrive as a `Uint8Array`, as from the extension. */
export const postDocument = (page: Page, message: Extract<HostMessage, { document: unknown }>) => {
  const {
    document: { bytes, ...document },
    ...fields
  } = message;
  return page.evaluate(
    ({ fields: rest, document: file, bytes: array }) =>
      window.postMessage({ ...rest, document: { ...file, bytes: new Uint8Array(array) } }, "*"),
    { fields, document, bytes: [...bytes] },
  );
};
