/**
 * The read-only `.docx` preview: a custom editor whose webview shows the
 * document's pages as folio lays them out. It re-renders when the file changes
 * on disk, so an edit an agent makes through the MCP server shows up.
 */

import * as vscode from "vscode";

import { debounce } from "./debounce";
import { isWebviewMessage, type HostMessage } from "./protocol";
import { renderDocument } from "./render";
import type { CliRuntime } from "./runtime";
import { createNonce, shellHtml } from "./shell";

export const PREVIEW_VIEW_TYPE = "folio.docxPreview";

/** How long the file must be quiet before a change re-renders it. */
const RERENDER_DELAY_MS = 300;

const describe = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

const baseName = (uri: vscode.Uri): string => uri.path.split("/").at(-1) ?? uri.path;

/** One open preview: its webview, its watcher, and the render in flight. */
class PreviewSession implements vscode.Disposable {
  private readonly disposables: vscode.Disposable[] = [];
  private readonly fileName: string;
  private readonly rerender = debounce(() => {
    void this.render();
  }, RERENDER_DELAY_MS);
  private inFlight: AbortController | undefined;
  /** What the webview shows, re-sent when a hidden webview comes back. */
  private latest: HostMessage;
  private disposed = false;
  private readonly uri: vscode.Uri;
  private readonly panel: vscode.WebviewPanel;
  private readonly runtime: CliRuntime;

  constructor(
    uri: vscode.Uri,
    panel: vscode.WebviewPanel,
    runtime: CliRuntime,
    webviewRoot: vscode.Uri,
  ) {
    this.uri = uri;
    this.panel = panel;
    this.runtime = runtime;
    this.fileName = baseName(uri);
    this.latest = { type: "loading", fileName: this.fileName };
    const { webview } = panel;
    webview.options = { enableScripts: true, localResourceRoots: [webviewRoot] };
    webview.html = shellHtml({
      nonce: createNonce(),
      scriptUri: webview.asWebviewUri(vscode.Uri.joinPath(webviewRoot, "main.js")).toString(),
      fileName: this.fileName,
    });

    // Watch the folder rather than the file: a folio write replaces the file
    // by rename. Events for other files in the folder are ignored.
    const watcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(vscode.Uri.joinPath(uri, ".."), "*"),
    );
    const onChange = (changed: vscode.Uri) => {
      if (changed.toString() === uri.toString()) this.rerender.schedule();
    };
    this.disposables.push(
      watcher,
      watcher.onDidChange(onChange),
      watcher.onDidCreate(onChange),
      watcher.onDidDelete(onChange),
      webview.onDidReceiveMessage((message: unknown) => {
        if (!isWebviewMessage(message)) return;
        if (message.type === "ready") void webview.postMessage(this.latest);
        if (message.type === "retry") void this.render();
      }),
    );
    void this.render();
  }

  private post(message: HostMessage): void {
    this.latest = message;
    void this.panel.webview.postMessage(message);
  }

  /** Render the file as it is now; a newer render cancels this one. */
  private async render(): Promise<void> {
    this.inFlight?.abort();
    const controller = new AbortController();
    this.inFlight = controller;
    // Keep showing the last pages while the new ones render.
    if (this.latest.type !== "document") this.post({ type: "loading", fileName: this.fileName });

    let bytes: Uint8Array;
    try {
      bytes = await vscode.workspace.fs.readFile(this.uri);
    } catch (error) {
      if (controller.signal.aborted || this.disposed) return;
      this.post({
        type: "error",
        fileName: this.fileName,
        message: `Cannot read ${this.fileName}: ${describe(error)}`,
      });
      return;
    }
    if (controller.signal.aborted || this.disposed) return;

    const outcome = await renderDocument({
      runtime: this.runtime,
      bytes,
      fileName: this.fileName,
      signal: controller.signal,
    }).catch((error: unknown) => ({ type: "error" as const, message: describe(error) }));
    if (controller.signal.aborted || this.disposed || outcome.type === "cancelled") return;
    this.inFlight = undefined;
    if (outcome.type === "document") {
      this.post({
        type: "document",
        fileName: this.fileName,
        html: outcome.html,
        pageCount: outcome.pageCount,
      });
      return;
    }
    this.post({
      type: "error",
      fileName: this.fileName,
      message: outcome.message,
      ...("hint" in outcome && outcome.hint !== undefined && { hint: outcome.hint }),
    });
  }

  dispose(): void {
    if (this.disposed) return;
    this.disposed = true;
    this.inFlight?.abort();
    this.rerender.dispose();
    for (const disposable of this.disposables.splice(0)) disposable.dispose();
  }
}

class DocxPreviewProvider implements vscode.CustomReadonlyEditorProvider {
  private readonly runtime: CliRuntime;
  private readonly webviewRoot: vscode.Uri;

  constructor(runtime: CliRuntime, webviewRoot: vscode.Uri) {
    this.runtime = runtime;
    this.webviewRoot = webviewRoot;
  }

  openCustomDocument(uri: vscode.Uri): vscode.CustomDocument {
    return { uri, dispose: () => undefined };
  }

  resolveCustomEditor(document: vscode.CustomDocument, panel: vscode.WebviewPanel): void {
    const session = new PreviewSession(document.uri, panel, this.runtime, this.webviewRoot);
    panel.onDidDispose(() => session.dispose());
  }
}

/** The file behind the active tab, whichever editor shows it. */
const activeTabUri = (): vscode.Uri | undefined => {
  const input: unknown = vscode.window.tabGroups.activeTabGroup.activeTab?.input;
  if (input instanceof vscode.TabInputText || input instanceof vscode.TabInputCustom) {
    return input.uri;
  }
  return undefined;
};

export const registerPreview = (
  context: vscode.ExtensionContext,
  runtime: CliRuntime,
): vscode.Disposable =>
  vscode.Disposable.from(
    vscode.window.registerCustomEditorProvider(
      PREVIEW_VIEW_TYPE,
      new DocxPreviewProvider(
        runtime,
        vscode.Uri.joinPath(context.extensionUri, "dist", "webview"),
      ),
      { supportsMultipleEditorsPerDocument: true },
    ),
    vscode.commands.registerCommand("folio.openPreview", async (target?: vscode.Uri) => {
      const uri = target ?? activeTabUri();
      if (uri === undefined) {
        void vscode.window.showInformationMessage("Folio: select a .docx file to preview.");
        return;
      }
      await vscode.commands.executeCommand("vscode.openWith", uri, PREVIEW_VIEW_TYPE);
    }),
  );
