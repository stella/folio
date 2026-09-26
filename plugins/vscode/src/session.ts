/**
 * One `.docx` open in the editor: the bytes the webview edits, the version
 * of the file on disk those edits were made from (the baseline), the undo
 * steps the workbench holds, and the lease. Saving, reverting, backing up,
 * answering a flush request, and reacting to the file changing on disk all
 * go through here. Nothing here imports `vscode`: the workbench, dialogs, and
 * the CLI are passed in, so it is unit-tested.
 */

import { encodeBackup } from "./backup";
import {
  isEditorMessage,
  type EditorMessage,
  type FolioEditingMode,
  type FolioEditorDocument,
  type FolioEditorSaveStrategy,
  type HostMessage,
} from "./editor-protocol";
import type { EditorLease } from "./lease";
import { fileVersionOf, type CliSaveOutcome, type CliSaveRequest } from "./save";

/** The user declined a save (a dialog's Cancel); the document stays dirty. */
export class SaveCancelled extends Error {
  constructor() {
    super("Save cancelled.");
    this.name = "SaveCancelled";
  }
}

export type RewriteChoice = "save" | "saveCopy" | "always";
export type StaleChoice = "overwrite" | "reload";
export type ConflictChoice = "reload" | "keep" | "saveCopy";

/** Dialogs and workbench actions; `undefined` from a dialog is its Cancel. */
export type SessionUi = {
  /** The one-time warning before a save that rewrites the whole package. */
  readonly confirmRewrite: () => Promise<RewriteChoice | undefined>;
  /** The file changed on disk since the edits' baseline; the save was refused. */
  readonly resolveStale: () => Promise<StaleChoice | undefined>;
  /** The file changed on disk while the document has unsaved edits. */
  readonly resolveConflict: () => Promise<ConflictChoice | undefined>;
  /** Ask where to save a copy of `bytes` and write it; `false` when cancelled. */
  readonly saveCopy: (bytes: Uint8Array, strategy: FolioEditorSaveStrategy) => Promise<boolean>;
  /** Save through the workbench, so its dirty state follows; `false` when it did not save. */
  readonly saveThroughWorkbench: () => Promise<boolean>;
  /** Revert through the workbench, so its dirty state follows. */
  readonly revertThroughWorkbench: () => Promise<void>;
  readonly notifyUpdated: (message: string) => void;
  readonly showLoadFailed: (message: string) => void;
  readonly showError: (message: string) => void;
  readonly showLeaseLost: (message: string) => void;
};

/**
 * The warning before the first full-package rewrite: once per extension
 * session, for every document, unless `folio.editor.confirmRewrite` is off.
 */
export type RewriteWarning = {
  readonly enabled: () => boolean;
  /** Turn the setting off ("Always"). */
  readonly disable: () => Promise<void>;
  confirmed: boolean;
};

export const needsRewriteConfirmation = (
  strategy: FolioEditorSaveStrategy,
  warning: RewriteWarning,
): boolean => strategy.type === "full-repack" && warning.enabled() && !warning.confirmed;

export type SessionLease = Pick<EditorLease, "token" | "ensure" | "release" | "forget" | "dispose">;

export type SessionDisk = {
  /** The document's bytes on disk now. */
  readonly read: () => Promise<Uint8Array>;
  /** The fileVersion of another file, or `null` when it does not exist. */
  readonly versionOf: (filePath: string) => Promise<string | null>;
  /** Write another file directly (Save As when the document itself is gone). */
  readonly write: (filePath: string, bytes: Uint8Array) => Promise<void>;
};

export type SessionOptions = {
  readonly fileName: string;
  /** The document's path on disk; `null` for a file system the CLI cannot reach (read-only). */
  readonly documentPath: string | null;
  readonly initial: {
    readonly bytes: Uint8Array;
    /** The version of the file on disk `bytes` were edited from. */
    readonly baseline: string;
    /** Restored from a hot-exit backup: dirty until saved or reverted. */
    readonly restored: boolean;
  };
  readonly author: string;
  readonly locale: string;
  readonly mode: FolioEditingMode;
  readonly disk: SessionDisk;
  readonly saveWithCli: (request: CliSaveRequest) => Promise<CliSaveOutcome>;
  /** The notice for a reload after someone else changed the file to `version`. */
  readonly describeChange: (version: string) => Promise<string>;
  readonly lease: SessionLease | null;
  readonly ui: SessionUi;
  readonly rewrite: RewriteWarning;
  /** The user made one undo step; the workbench gets one edit. */
  readonly onEdit: () => void;
  readonly serializeTimeoutMs?: number;
};

type Serialized = { readonly bytes: Uint8Array; readonly strategy: FolioEditorSaveStrategy };

type PendingSerialization = {
  readonly resolve: (serialized: Serialized) => void;
  readonly reject: (error: Error) => void;
  readonly timer: ReturnType<typeof setTimeout>;
};

const DEFAULT_SERIALIZE_TIMEOUT_MS = 60_000;

/** Marks a restored document dirty: no undo step has this id. */
const RESTORED = -1;
/** The saved point of an empty undo stack. */
const EMPTY = 0;

export class DocxSession {
  private readonly options: SessionOptions;
  private post: ((message: HostMessage) => void) | undefined;
  /** The bytes the webview loads: as opened, as last saved, or as reloaded. */
  private bytes: Uint8Array;
  /** The version on disk the edits were made from; every save expects it. */
  private baselineVersion: string;
  private mode: FolioEditingMode;
  /** Bumped by each load; a serialization of an older load is stale. */
  private generation = 0;
  /** The undo steps applied, and those undone and redoable, by id. */
  private applied: number[] = [];
  private undone: number[] = [];
  private nextEditId = 1;
  /** The top undo step when the document was last saved, reverted or reloaded. */
  private savedAt: number;
  /** Counts every edit; a save compares it to see whether the user kept typing. */
  private editCount = 0;
  private nextRequestId = 1;
  private readonly pending = new Map<number, PendingSerialization>();
  private saving: Promise<void> | undefined;
  private diskCheckDue = false;
  private flushing = false;
  private conflictOpen = false;
  private disposed = false;

  constructor(options: SessionOptions) {
    this.options = options;
    this.bytes = options.initial.bytes;
    this.baselineVersion = options.initial.baseline;
    this.mode = options.mode;
    this.savedAt = options.initial.restored ? RESTORED : EMPTY;
    if (options.initial.restored) void options.lease?.ensure();
  }

  get baseline(): string {
    return this.baselineVersion;
  }

  /** The `fileVersion` the webview's current load carries. */
  private get loadedVersion(): string {
    return `load-${String(this.generation)}`;
  }

  private get top(): number {
    return this.applied.at(-1) ?? EMPTY;
  }

  /** Whether the document differs from what was last saved, reverted or reloaded. */
  get dirty(): boolean {
    return this.top !== this.savedAt;
  }

  /** Connect the webview; `undefined` when it closed. */
  attach(post: ((message: HostMessage) => void) | undefined): void {
    this.post = post;
    if (post === undefined) this.rejectPending("The editor closed.");
  }

  /** A copy of `bytes`: a posted message is transferred, and the session keeps its own. */
  private webviewDocument(bytes: Uint8Array): FolioEditorDocument {
    return {
      bytes: bytes.slice(),
      fileVersion: this.loadedVersion,
      fileName: this.options.fileName,
    };
  }

  private load(): void {
    this.generation += 1;
    this.post?.({
      type: "load",
      document: this.webviewDocument(this.bytes),
      author: this.options.author,
      mode: this.mode,
      locale: this.options.locale,
    });
  }

  handleMessage(message: unknown): void {
    if (!isEditorMessage(message)) return;
    this.handle(message);
  }

  private handle(message: EditorMessage): void {
    switch (message.type) {
      case "ready":
        this.load();
        return;
      case "loaded":
      case "dirty":
        return;
      case "loadFailed":
        this.options.ui.showLoadFailed(message.message);
        return;
      case "edit":
        this.recordEdit();
        return;
      case "modeChanged":
        this.mode = message.mode;
        return;
      case "serialized": {
        const pending = this.takePending(message.requestId);
        if (pending === undefined) return;
        if (message.fileVersion !== this.loadedVersion) {
          pending.reject(
            new Error("The document was reloaded while it was being saved. Save again."),
          );
          return;
        }
        pending.resolve({ bytes: message.bytes, strategy: message.strategy });
        return;
      }
      case "serializeFailed":
        this.takePending(message.requestId)?.reject(new Error(message.message));
        return;
      case "error":
        this.options.ui.showError(message.message);
        return;
      default: {
        const unreachable: never = message;
        return unreachable;
      }
    }
  }

  private recordEdit(): void {
    this.applied.push(this.nextEditId);
    this.nextEditId += 1;
    this.undone = [];
    this.editCount += 1;
    this.options.onEdit();
    void this.options.lease?.ensure();
  }

  /** The workbench undoes one of this document's edits. */
  undo(): void {
    const id = this.applied.pop();
    if (id !== undefined) this.undone.push(id);
    this.editCount += 1;
    this.post?.({ type: "undo" });
  }

  /** The workbench redoes one of this document's edits. */
  redo(): void {
    const id = this.undone.pop();
    if (id !== undefined) this.applied.push(id);
    this.editCount += 1;
    this.post?.({ type: "redo" });
    if (this.dirty) void this.options.lease?.ensure();
  }

  private takePending(requestId: number): PendingSerialization | undefined {
    const pending = this.pending.get(requestId);
    if (pending === undefined) return undefined;
    this.pending.delete(requestId);
    clearTimeout(pending.timer);
    return pending;
  }

  private rejectPending(reason: string): void {
    for (const requestId of [...this.pending.keys()]) {
      this.takePending(requestId)?.reject(new Error(reason));
    }
  }

  /** The webview's document as `.docx` bytes. */
  serialize(): Promise<Serialized> {
    const { post } = this;
    if (post === undefined) return Promise.reject(new Error("The editor is not open."));
    const requestId = this.nextRequestId;
    this.nextRequestId += 1;
    const timeoutMs = this.options.serializeTimeoutMs ?? DEFAULT_SERIALIZE_TIMEOUT_MS;
    return new Promise<Serialized>((resolve, reject) => {
      const timer = setTimeout(() => {
        this.takePending(requestId)?.reject(
          new Error(
            `The editor did not produce the document within ${String(timeoutMs / 1000)} seconds.`,
          ),
        );
      }, timeoutMs);
      this.pending.set(requestId, { resolve, reject, timer });
      post({ type: "serialize", requestId });
    });
  }

  /** Save in place. Throws {@link SaveCancelled} when the user cancels a dialog. */
  async save(): Promise<void> {
    const running = this.runSave();
    this.saving = running;
    try {
      await running;
    } finally {
      this.saving = undefined;
      if (this.diskCheckDue) {
        this.diskCheckDue = false;
        void this.diskChanged();
      }
    }
  }

  private async runSave(): Promise<void> {
    const { documentPath } = this.options;
    if (documentPath === null) {
      throw new Error(`${this.options.fileName} is read-only here. Use Save As to keep a copy.`);
    }
    const editCount = this.editCount;
    const savedTop = this.top;
    const { bytes, strategy } = await this.serialize();
    if (needsRewriteConfirmation(strategy, this.options.rewrite)) {
      const choice = await this.options.ui.confirmRewrite();
      if (choice === undefined) throw new SaveCancelled();
      if (choice === "saveCopy") {
        await this.options.ui.saveCopy(bytes, strategy);
        throw new SaveCancelled();
      }
      this.options.rewrite.confirmed = true;
      if (choice === "always") await this.options.rewrite.disable();
    }

    const { lease } = this.options;
    let expectedVersion = this.baselineVersion;
    let leaseToken = lease?.token;
    for (;;) {
      const outcome = await this.options.saveWithCli({
        documentPath,
        bytes,
        expectedVersion,
        author: this.options.author,
        strategy,
        ...(leaseToken !== undefined && { leaseToken }),
      });
      if (outcome.type === "saved") {
        this.baselineVersion = outcome.fileVersion;
        this.bytes = bytes;
        this.savedAt = savedTop;
        // Keep the lease while the user typed on during the save; a flush
        // lets it go regardless, so the waiting write can land.
        if (this.flushing || this.editCount === editCount) await lease?.release();
        return;
      }
      if (outcome.code === "stale_version") {
        const choice = await this.options.ui.resolveStale();
        if (choice === "overwrite") {
          expectedVersion = fileVersionOf(await this.options.disk.read());
          continue;
        }
        if (choice === "reload") {
          await this.reloadFromDisk();
          return;
        }
        throw new SaveCancelled();
      }
      if (outcome.code === "locked" && leaseToken !== undefined) {
        // The lease expired or was taken over: save the way any writer would.
        lease?.forget();
        leaseToken = undefined;
        continue;
      }
      throw new Error(
        outcome.hint === undefined ? outcome.message : `${outcome.message} ${outcome.hint}`,
      );
    }
  }

  /**
   * Save the current document to another file on disk (Save As, Save a
   * Copy). The document itself, its baseline and its lease are unchanged.
   */
  async saveCopyTo(
    destinationPath: string,
    serialized?: { readonly bytes: Uint8Array; readonly strategy: FolioEditorSaveStrategy },
  ): Promise<void> {
    const { bytes, strategy } = serialized ?? (await this.serialize());
    const { documentPath } = this.options;
    const source = documentPath === null ? null : await this.options.disk.read().catch(() => null);
    if (documentPath === null || source === null) {
      // The CLI saves from a document on disk; without one, write the copy directly.
      await this.options.disk.write(destinationPath, bytes);
      return;
    }
    const outcome = await this.options.saveWithCli({
      documentPath,
      bytes,
      // A copy leaves the document alone, so any version of it will do.
      expectedVersion: fileVersionOf(source),
      author: this.options.author,
      strategy,
      destination: {
        path: destinationPath,
        expectedVersion: await this.options.disk.versionOf(destinationPath),
      },
    });
    if (outcome.type === "error") {
      throw new Error(
        outcome.hint === undefined ? outcome.message : `${outcome.message} ${outcome.hint}`,
      );
    }
  }

  /** Show the file as it is on disk now, dropping unsaved edits. */
  async reloadFromDisk(read?: Uint8Array): Promise<void> {
    const bytes = read ?? (await this.options.disk.read());
    this.bytes = bytes;
    this.baselineVersion = fileVersionOf(bytes);
    this.savedAt = this.top;
    this.generation += 1;
    this.post?.({
      type: "reload",
      document: this.webviewDocument(bytes),
    });
    await this.options.lease?.release();
  }

  /** The workbench's revert. */
  revert(): Promise<void> {
    return this.reloadFromDisk();
  }

  /** The hot-exit backup: the unsaved document behind its baseline. */
  async backup(): Promise<Uint8Array> {
    const { bytes } = await this.serialize();
    return encodeBackup({ baseline: this.baselineVersion, bytes });
  }

  /**
   * The file on disk changed (the watcher saw it). Our own save is not a
   * change; a clean document reloads and says who changed it; a dirty one
   * asks what to keep.
   */
  async diskChanged(): Promise<void> {
    if (this.disposed) return;
    if (this.saving !== undefined) {
      // The save's own rename; look again once its receipt has the new version.
      this.diskCheckDue = true;
      return;
    }
    let bytes: Uint8Array;
    try {
      bytes = await this.options.disk.read();
    } catch {
      // Deleted or moved: the workbench shows that on the tab.
      return;
    }
    const version = fileVersionOf(bytes);
    if (version === this.baselineVersion || this.disposed) return;
    if (!this.dirty) {
      await this.reloadFromDisk(bytes);
      this.options.ui.notifyUpdated(await this.options.describeChange(version));
      return;
    }
    if (this.conflictOpen) return;
    this.conflictOpen = true;
    let choice: ConflictChoice | undefined;
    try {
      choice = await this.options.ui.resolveConflict();
    } finally {
      this.conflictOpen = false;
    }
    switch (choice) {
      case "reload":
        await this.options.ui.revertThroughWorkbench();
        return;
      case "keep":
        // The next save replaces their version (backed up) without asking again.
        this.baselineVersion = version;
        void this.options.lease?.ensure();
        return;
      case "saveCopy": {
        const serialized = await this.serialize();
        if (await this.options.ui.saveCopy(serialized.bytes, serialized.strategy)) {
          await this.options.ui.revertThroughWorkbench();
        }
        return;
      }
      case undefined:
        return;
      default: {
        const unreachable: never = choice;
        return unreachable;
      }
    }
  }

  /**
   * A write asks for the lease: save the unsaved edits, then let it go so
   * the write lands on the saved version. The reload that shows the write
   * follows from the watcher.
   */
  async flush(): Promise<void> {
    if (this.flushing || this.disposed) return;
    this.flushing = true;
    try {
      if (await this.options.ui.saveThroughWorkbench()) await this.options.lease?.release();
    } finally {
      this.flushing = false;
    }
  }

  leaseLost(message: string): void {
    this.options.ui.showLeaseLost(message);
  }

  async dispose(): Promise<void> {
    this.disposed = true;
    this.rejectPending("The document closed.");
    await this.options.lease?.dispose();
  }
}
