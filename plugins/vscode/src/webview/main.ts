/**
 * The preview webview's script. It shows the pages the extension sends, keeps
 * the reader's scroll position across re-renders, and reports errors.
 *
 * The rendered document is parsed inertly with DOMParser. Only its `<style>`
 * text and its body's page elements are kept; the page's policy blocks any
 * script or remote load that could be hiding in them regardless.
 */

import { isHostMessage, pageCountLabel, type HostMessage, type WebviewMessage } from "../protocol";

type VsCodeApi = {
  postMessage: (message: WebviewMessage) => void;
  getState: () => unknown;
  setState: (state: unknown) => void;
};

declare const acquireVsCodeApi: () => VsCodeApi;

const vscode = acquireVsCodeApi();

const element = (id: string): HTMLElement => {
  const found = document.getElementById(id);
  if (found === null) throw new Error(`The preview page has no #${id}.`);
  return found;
};

const documentStyle = element("folio-document-style");
const name = element("folio-name");
const status = element("folio-status");
const errorBox = element("folio-error");
const errorMessage = element("folio-error-message");
const errorHint = element("folio-error-hint");
const retry = element("folio-retry");
const pages = element("folio-pages");

/** Elements a rendered page never needs; dropped before the pages are adopted. */
const INERT_SELECTOR = "script, iframe, object, embed, link, meta, base, form";

const scrollState = (): number => {
  const state = vscode.getState();
  if (typeof state === "object" && state !== null && "scrollY" in state) {
    const { scrollY } = state;
    return typeof scrollY === "number" ? scrollY : 0;
  }
  return 0;
};

let restoreScroll = scrollState();
let saveQueued = false;

window.addEventListener("scroll", () => {
  if (saveQueued) return;
  saveQueued = true;
  requestAnimationFrame(() => {
    saveQueued = false;
    vscode.setState({ scrollY: window.scrollY });
  });
});

const showDocument = (html: string, pageCount: number) => {
  const parsed = new DOMParser().parseFromString(html, "text/html");
  for (const node of parsed.querySelectorAll(INERT_SELECTOR)) node.remove();
  const css = [...parsed.querySelectorAll("style")].map((style) => style.textContent).join("\n");
  // A re-created webview restores the saved position once; a re-render keeps the current one.
  const scrollY = restoreScroll > 0 ? restoreScroll : window.scrollY;
  restoreScroll = 0;
  documentStyle.textContent = css;
  pages.replaceChildren(...[...parsed.body.childNodes].map((node) => document.adoptNode(node)));
  pages.classList.remove("stale");
  errorBox.classList.remove("visible");
  status.textContent = pageCountLabel(pageCount);
  if (scrollY > 0) window.scrollTo(0, scrollY);
};

const showError = (message: string, hint: string | undefined) => {
  errorMessage.textContent = message;
  errorHint.textContent = hint ?? "";
  errorBox.classList.add("visible");
  pages.classList.add("stale");
  status.textContent = "Cannot render";
};

const handle = (message: HostMessage) => {
  name.textContent = message.fileName;
  switch (message.type) {
    case "loading":
      status.textContent = "Rendering…";
      errorBox.classList.remove("visible");
      pages.classList.add("stale");
      return;
    case "document":
      showDocument(message.html, message.pageCount);
      return;
    case "error":
      showError(message.message, message.hint);
      return;
    default: {
      const unreachable: never = message;
      return unreachable;
    }
  }
};

window.addEventListener("message", (event: MessageEvent<unknown>) => {
  if (isHostMessage(event.data)) handle(event.data);
});

retry.addEventListener("click", () => {
  vscode.postMessage({ type: "retry" });
});

vscode.postMessage({ type: "ready" });
