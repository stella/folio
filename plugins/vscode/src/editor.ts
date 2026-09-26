/**
 * The `.docx` editor, the default for `.docx` files: a custom editor whose
 * webview runs folio's editor (`packages/editor-web`, copied to
 * `dist/editor`). The workbench owns the undo stack, dirty state, and hot
 * exit; each save goes through `folio save`, which checks the version the
 * edits were made from, backs up the previous file, and journals the save.
 * While the document has unsaved edits the extension host holds the editor
 * lease, so a folio write (an agent's tool call) asks it to save first
 * instead of failing. See `session.ts` for the rules.
 */

import { readFile } from "node:fs/promises";
import path from "node:path";
import * as vscode from "vscode";

import { decodeBackup, type Backup } from "./backup";
import { debounce } from "./debounce";
import type { FolioEditingMode, FolioEditorSaveStrategy } from "./editor-protocol";
import {
  CONFIRM_REWRITE_SETTING,
  initialMode,
  osUserName,
  readGitUserName,
  resolveEditorAuthor,
  TRACK_CHANGES_SETTING,
} from "./editor-settings";
import { readCommitForVersion, updatedNotice } from "./journal";
import { EditorLease } from "./lease";
import type { CliRuntime } from "./runtime";
import { fileVersionOf, saveWithCli } from "./save";
import { DocxSession, SaveCancelled, type RewriteWarning, type SessionUi } from "./session";
import { createNonce, editorShellHtml } from "./shell";

export const EDITOR_VIEW_TYPE = "folio.docxEditor";

const AUTHOR_SETTING = "folio.author";

/** How long the file must be quiet before a change on disk is looked at. */
const DISK_CHANGE_DELAY_MS = 300;

const baseName = (uri: vscode.Uri): string => uri.path.split("/").at(-1) ?? uri.path;

const describe = (error: unknown): string =>
  error instanceof Error ? error.message : String(error);

/** `Report.docx` becomes `Report copy.docx`, beside it. */
const copyUri = (uri: vscode.Uri): vscode.Uri => {
  const name = baseName(uri);
  const extension = path.extname(name);
  const stem = name.slice(0, name.length - extension.length);
  return vscode.Uri.joinPath(uri, "..", `${stem} copy${extension}`);
};

/** What a document opened with, for its session. */
type Opened = {
  readonly documentPath: string | null;
  readonly author: string;
  readonly bytes: Uint8Array;
  readonly baseline: string;
  readonly restored: boolean;
  readonly mode: FolioEditingMode;
};

/** An open `.docx`: its session, and the watcher that tells it the file changed. */
class FolioDocxDocument implements vscode.CustomDocument {
  readonly uri: vscode.Uri;
  readonly fileName: string;
  readonly session: DocxSession;
  panel: vscode.WebviewPanel | undefined;
  /** What the webview sent, kept only for the smoke test (`FOLIO_VSCODE_TEST`). */
  readonly received: unknown[] = [];
  readonly disposables: vscode.Disposable[] = [];

  /** `createSession` may keep the document: its callbacks run only once it is built. */
  constructor(uri: vscode.Uri, createSession: (document: FolioDocxDocument) => DocxSession) {
    this.uri = uri;
    this.fileName = baseName(uri);
    this.session = createSession(this);
  }

  /**
   * Watch the folder rather than the file: a folio write replaces the file by
   * rename. Events for other files in the folder are ignored.
   */
  watch(): void {
    const check = debounce(() => {
      void this.session.diskChanged();
    }, DISK_CHANGE_DELAY_MS);
    const watcher = vscode.workspace.createFileSystemWatcher(
      new vscode.RelativePattern(vscode.Uri.joinPath(this.uri, ".."), "*"),
    );
    const onChange = (changed: vscode.Uri) => {
      if (changed.toString() === this.uri.toString()) check.schedule();
    };
    this.disposables.push(watcher, watcher.onDidChange(onChange), watcher.onDidCreate(onChange), {
      dispose: () => check.dispose(),
    });
  }

  dispose(): void {
    for (const disposable of this.disposables.splice(0)) disposable.dispose();
    void this.session.dispose();
  }
}

class DocxEditorProvider implements vscode.CustomEditorProvider<FolioDocxDocument> {
  private readonly changed = new vscode.EventEmitter<
    vscode.CustomDocumentEditEvent<FolioDocxDocument>
  >();
  readonly onDidChangeCustomDocument = this.changed.event;
  private readonly runtime: CliRuntime;
  private readonly editorRoot: vscode.Uri;
  /** Shared by every document: the rewrite warning shows once per session. */
  private readonly rewrite: RewriteWarning = {
    enabled: () => vscode.workspace.getConfiguration().get<boolean>(CONFIRM_REWRITE_SETTING, true),
    disable: async () => {
      await vscode.workspace
        .getConfiguration()
        .update(CONFIRM_REWRITE_SETTING, false, vscode.ConfigurationTarget.Global);
    },
    confirmed: false,
  };

  private readonly testMode: boolean;
  /** The open documents, by URI. */
  private readonly documents = new Map<string, FolioDocxDocument>();
  /** Documents "Open Read-Only" is opening, by URI. */
  private readonly readOnlyRequests = new Set<string>();

  constructor(runtime: CliRuntime, editorRoot: vscode.Uri, testMode: boolean) {
    this.runtime = runtime;
    this.editorRoot = editorRoot;
    this.testMode = testMode;
  }

  dispose(): void {
    this.changed.dispose();
  }

  documentFor(uri: vscode.Uri): FolioDocxDocument | undefined {
    return this.documents.get(uri.toString());
  }

  /** Open `uri` in the editor in viewing mode, or switch its open editor to it. */
  async openReadOnly(uri: vscode.Uri): Promise<void> {
    const open = this.documentFor(uri);
    if (open !== undefined) {
      open.session.setMode("viewing");
      open.panel?.reveal();
      return;
    }
    const key = uri.toString();
    this.readOnlyRequests.add(key);
    try {
      await vscode.commands.executeCommand("vscode.openWith", uri, EDITOR_VIEW_TYPE);
    } finally {
      this.readOnlyRequests.delete(key);
    }
  }

  async openCustomDocument(
    uri: vscode.Uri,
    openContext: vscode.CustomDocumentOpenContext,
  ): Promise<FolioDocxDocument> {
    // The CLI saves files on disk; any other file system opens read-only.
    const documentPath = uri.scheme === "file" ? uri.fsPath : null;
    const readOnly = this.readOnlyRequests.has(uri.toString());
    const restored =
      openContext.backupId === undefined ? null : await this.readBackup(openContext.backupId);
    const bytes = restored?.bytes ?? (await vscode.workspace.fs.readFile(uri));
    const configuration = vscode.workspace.getConfiguration();
    const author = resolveEditorAuthor({
      setting: configuration.get<string>(AUTHOR_SETTING),
      gitUserName: await readGitUserName(
        documentPath === null ? undefined : path.dirname(documentPath),
      ),
      osUserName: osUserName(),
    });
    const opened: Opened = {
      documentPath,
      author,
      bytes,
      baseline: restored?.baseline ?? fileVersionOf(bytes),
      restored: restored !== null,
      mode: readOnly
        ? "viewing"
        : initialMode(configuration.get(TRACK_CHANGES_SETTING), documentPath !== null),
    };
    const document = new FolioDocxDocument(uri, (created) => this.createSession(created, opened));
    document.watch();
    const key = uri.toString();
    this.documents.set(key, document);
    document.disposables.push({
      dispose: () => {
        if (this.documents.get(key) === document) this.documents.delete(key);
      },
    });
    return document;
  }

  private createSession(document: FolioDocxDocument, opened: Opened): DocxSession {
    const { uri, fileName } = document;
    const { documentPath, author, bytes, baseline, restored, mode } = opened;
    const lease =
      documentPath === null
        ? null
        : new EditorLease(documentPath, {
            onFlushRequest: (request) => void document.session.flush(request.owner),
            onLost: (message) => document.session.leaseLost(message),
          });
    return new DocxSession({
      fileName,
      documentPath,
      initial: { bytes, baseline, restored },
      author,
      locale: vscode.env.language,
      mode,
      disk: {
        read: async () => await vscode.workspace.fs.readFile(uri),
        versionOf: async (filePath) => {
          try {
            return fileVersionOf(await readFile(filePath));
          } catch {
            return null;
          }
        },
        write: async (filePath, data) => {
          await vscode.workspace.fs.writeFile(vscode.Uri.file(filePath), data);
        },
      },
      saveWithCli: (request) => saveWithCli(this.runtime, request),
      describeChange: async (version) =>
        updatedNotice(
          fileName,
          documentPath === null ? null : await readCommitForVersion(documentPath, version),
        ),
      lease,
      ui: this.sessionUi(document),
      rewrite: this.rewrite,
      onEdit: () =>
        this.changed.fire({
          document,
          label: "Edit",
          undo: () => document.session.undo(),
          redo: () => document.session.redo(),
        }),
    });
  }

  /** The backup's document, or `null` (the file on disk opens) when it cannot be read. */
  private async readBackup(backupId: string): Promise<Backup | null> {
    try {
      const backup = decodeBackup(await vscode.workspace.fs.readFile(vscode.Uri.parse(backupId)));
      if (backup !== null) return backup;
    } catch {
      // Reported below.
    }
    void vscode.window.showWarningMessage(
      "Folio could not restore the unsaved edits of this document; it opened as saved on disk.",
    );
    return null;
  }

  private sessionUi(document: FolioDocxDocument): SessionUi {
    const { fileName } = document;
    return {
      confirmRewrite: async () => {
        const choice = await vscode.window.showWarningMessage(
          `This save rewrites the whole document package; a backup of the previous version is kept in .folio/backups/${fileName}/.`,
          {
            modal: true,
            detail: "Always saves without asking again (the folio.editor.confirmRewrite setting).",
          },
          "Save",
          "Save a Copy…",
          "Always",
        );
        if (choice === "Save") return "save";
        if (choice === "Save a Copy…") return "saveCopy";
        if (choice === "Always") return "always";
        return undefined;
      },
      resolveStale: async () => {
        const choice = await vscode.window.showWarningMessage(
          `${fileName} changed on disk since you opened it.`,
          {
            modal: true,
            detail:
              "Overwrite saves your version over it; the version on disk is kept in .folio/backups. Reload discards your unsaved edits.",
          },
          "Overwrite",
          "Reload",
        );
        if (choice === "Overwrite") return "overwrite";
        if (choice === "Reload") return "reload";
        return undefined;
      },
      resolveConflict: async () => {
        const choice = await vscode.window.showWarningMessage(
          `${fileName} changed on disk while you have unsaved edits.`,
          "Reload Theirs",
          "Keep Mine",
          "Save Mine as Copy…",
        );
        if (choice === "Reload Theirs") return "reload";
        if (choice === "Keep Mine") return "keep";
        if (choice === "Save Mine as Copy…") return "saveCopy";
        return undefined;
      },
      saveCopy: async (bytes, strategy) => {
        const target = await vscode.window.showSaveDialog({
          defaultUri: copyUri(document.uri),
          filters: { "DOCX document": ["docx"] },
          saveLabel: "Save Copy",
        });
        if (target === undefined) return false;
        await this.writeCopy(document, target, { bytes, strategy });
        return true;
      },
      saveThroughWorkbench: async () => (await vscode.workspace.save(document.uri)) !== undefined,
      revertThroughWorkbench: async () => {
        // Revert acts on the active editor.
        document.panel?.reveal(undefined, false);
        await vscode.commands.executeCommand("workbench.action.files.revert");
      },
      notifyUpdated: (message) => {
        vscode.window.setStatusBarMessage(`$(sync) ${message}`, 8000);
      },
      showLoadFailed: (message) => {
        void vscode.window.showErrorMessage(`Folio could not open ${fileName}: ${message}`);
      },
      showInfo: (message) => {
        void vscode.window.showInformationMessage(`Folio: ${message}`);
      },
      showError: (message) => {
        void vscode.window.showErrorMessage(`Folio: ${message}`);
      },
      showLeaseLost: (message) => {
        void vscode.window.showWarningMessage(
          `Folio: ${fileName} is no longer held for your unsaved edits (${message}). A folio write may change it before you save.`,
        );
      },
    };
  }

  /** Save As and Save a Copy: through `folio save -o` on disk, directly elsewhere. */
  private async writeCopy(
    document: FolioDocxDocument,
    target: vscode.Uri,
    serialized?: { readonly bytes: Uint8Array; readonly strategy: FolioEditorSaveStrategy },
  ): Promise<void> {
    if (target.scheme === "file") {
      await document.session.saveCopyTo(target.fsPath, serialized);
      return;
    }
    const { bytes } = serialized ?? (await document.session.serialize());
    await vscode.workspace.fs.writeFile(target, bytes);
  }

  resolveCustomEditor(document: FolioDocxDocument, panel: vscode.WebviewPanel): void {
    const { webview } = panel;
    webview.options = { enableScripts: true, localResourceRoots: [this.editorRoot] };
    webview.html = editorShellHtml({
      nonce: createNonce(),
      cspSource: webview.cspSource,
      scriptUri: webview.asWebviewUri(vscode.Uri.joinPath(this.editorRoot, "editor.js")).toString(),
      styleUri: webview.asWebviewUri(vscode.Uri.joinPath(this.editorRoot, "editor.css")).toString(),
      fileName: document.fileName,
      testMode: this.testMode,
    });
    document.panel = panel;
    document.session.attach((message) => void webview.postMessage(message));
    const received = webview.onDidReceiveMessage((message: unknown) => {
      if (this.testMode) document.received.push(message);
      document.session.handleMessage(message);
    });
    panel.onDidDispose(() => {
      received.dispose();
      if (document.panel === panel) {
        document.panel = undefined;
        document.session.attach(undefined);
      }
    });
  }

  async saveCustomDocument(document: FolioDocxDocument): Promise<void> {
    try {
      await document.session.save();
    } catch (error) {
      if (error instanceof SaveCancelled) throw new vscode.CancellationError();
      throw new Error(`Folio could not save ${document.fileName}: ${describe(error)}`);
    }
  }

  async saveCustomDocumentAs(document: FolioDocxDocument, destination: vscode.Uri): Promise<void> {
    if (destination.toString() === document.uri.toString()) {
      await this.saveCustomDocument(document);
      return;
    }
    await this.writeCopy(document, destination);
  }

  async revertCustomDocument(document: FolioDocxDocument): Promise<void> {
    await document.session.revert();
  }

  async backupCustomDocument(
    document: FolioDocxDocument,
    context: vscode.CustomDocumentBackupContext,
  ): Promise<vscode.CustomDocumentBackup> {
    const data = await document.session.backup();
    await vscode.workspace.fs.writeFile(context.destination, data);
    return {
      id: context.destination.toString(),
      delete: () => {
        void vscode.workspace.fs.delete(context.destination).then(undefined, () => undefined);
      },
    };
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

/**
 * What the smoke test (`test/smoke`) reaches through the extension's exports.
 * `activate` returns it only when `FOLIO_VSCODE_TEST=1`, which is also what
 * puts the test script in the webview.
 */
export type EditorTestHooks = {
  /** The messages the webview of `uri` sent, oldest first. */
  readonly received: (uri: vscode.Uri) => readonly unknown[];
  /** Post `message` to the webview of `uri`; `false` when it has none. */
  readonly post: (uri: vscode.Uri, message: unknown) => Promise<boolean>;
};

export type RegisteredEditor = {
  readonly disposable: vscode.Disposable;
  readonly testHooks: EditorTestHooks;
};

export const registerEditor = (
  context: vscode.ExtensionContext,
  runtime: CliRuntime,
  testMode: boolean,
): RegisteredEditor => {
  const provider = new DocxEditorProvider(
    runtime,
    vscode.Uri.joinPath(context.extensionUri, "dist", "editor"),
    testMode,
  );
  const disposable = vscode.Disposable.from(
    provider,
    vscode.window.registerCustomEditorProvider(EDITOR_VIEW_TYPE, provider, {
      supportsMultipleEditorsPerDocument: false,
      webviewOptions: { retainContextWhenHidden: true },
    }),
    vscode.commands.registerCommand("folio.openReadOnly", async (target?: vscode.Uri) => {
      const uri = target ?? activeTabUri();
      if (uri === undefined) {
        void vscode.window.showInformationMessage("Folio: select a .docx file to open.");
        return;
      }
      await provider.openReadOnly(uri);
    }),
  );
  const testHooks: EditorTestHooks = {
    received: (uri) => provider.documentFor(uri)?.received ?? [],
    post: async (uri, message) => {
      const panel = provider.documentFor(uri)?.panel;
      return panel === undefined ? false : await panel.webview.postMessage(message);
    },
  };
  return { disposable, testHooks };
};
