/**
 * Entry for a VS Code custom-editor webview. It speaks the protocol in
 * `./protocol` over `postMessage`: it says `ready`, mounts the editor on the
 * extension's `load`, and from then on turns host messages into editor
 * commands and editor notifications into messages. The extension owns the
 * file, the undo stack and saving.
 */

import type { FolioEditorHandle, FolioEditorHost } from "./host";
import { mountFolioEditor } from "./mount";
import { isHostMessage } from "./protocol";
import type { EditorMessage, HostMessage } from "./protocol";

type VsCodeApi = { postMessage: (message: EditorMessage) => void };

declare const acquireVsCodeApi: () => VsCodeApi;

const vscode = acquireVsCodeApi();
const post = (message: EditorMessage) => vscode.postMessage(message);

/**
 * VS Code marks the theme kind on `body`; folio switches its dark palette with
 * a `dark` class on an ancestor. The `--vscode-*` colors themselves reach the
 * editor through the token mapping in `vscode.css`.
 */
const syncThemeKind = () => {
  const { classList } = document.body;
  const dark =
    classList.contains("vscode-dark") ||
    (classList.contains("vscode-high-contrast") &&
      !classList.contains("vscode-high-contrast-light"));
  document.documentElement.classList.toggle("dark", dark);
};

syncThemeKind();
new MutationObserver(syncThemeKind).observe(document.body, {
  attributes: true,
  attributeFilter: ["class"],
});

const root = document.createElement("div");
root.id = "folio-editor";
document.body.append(root);

let editor: FolioEditorHandle | null = null;

const errorMessage = (error: unknown): string =>
  error instanceof Error ? error.message : "folio could not serialize the document.";

const serialize = async (requestId: number) => {
  if (editor === null) {
    post({ type: "serializeFailed", requestId, message: "The editor has not loaded a document." });
    return;
  }
  try {
    const { bytes, fileVersion, strategy } = await editor.serialize();
    post({ type: "serialized", requestId, bytes, fileVersion, strategy });
  } catch (error) {
    post({ type: "serializeFailed", requestId, message: errorMessage(error) });
  }
};

const mount = ({ document, author, mode, locale }: Extract<HostMessage, { type: "load" }>) => {
  const host: FolioEditorHost = {
    init: { document, author, mode, locale },
    onLoaded: (fileVersion) => post({ type: "loaded", fileVersion }),
    onLoadFailed: (message) => post({ type: "loadFailed", message }),
    onEdit: () => post({ type: "edit" }),
    onDirtyChange: (dirty) => post({ type: "dirty", dirty }),
    onModeChange: (next) => post({ type: "modeChanged", mode: next }),
    onError: (message) => post({ type: "error", message }),
  };
  editor = mountFolioEditor(root, host);
};

const handle = (message: HostMessage) => {
  switch (message.type) {
    case "load":
      // A second `load` (after one that failed, say) replaces the document.
      if (editor === null) mount(message);
      else editor.reload(message.document);
      return;
    case "reload":
      editor?.reload(message.document);
      return;
    case "serialize":
      void serialize(message.requestId);
      return;
    case "undo":
      editor?.undo();
      return;
    case "redo":
      editor?.redo();
      return;
    case "setMode":
      editor?.setMode(message.mode);
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

post({ type: "ready" });
