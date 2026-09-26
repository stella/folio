/**
 * The contract between the editor bundle and the page that embeds it. The
 * host (a VS Code webview's extension side, a `folio serve --edit` page) owns
 * the file: it hands the editor bytes, keeps the undo stack the user sees, and
 * asks for bytes back when it saves. The editor owns only the live document.
 */

import type { FolioEditingMode } from "./modes";

/** A `.docx` file as the host read it. */
export type FolioEditorDocument = {
  readonly bytes: Uint8Array;
  /**
   * The host's version of these bytes (an mtime, an etag, a revision), opaque
   * to the editor. It comes back with every serialization so the host can tell
   * which file state the saved bytes were edited from.
   */
  readonly fileVersion: string;
  /** Shown in the editor's chrome; never used to reach the file. */
  readonly fileName: string;
};

/** What the editor opens with. */
export type FolioEditorInit = {
  readonly document: FolioEditorDocument;
  /** Recorded on tracked changes and comments. */
  readonly author: string;
  readonly mode: FolioEditingMode;
  /** UI language; the editor falls back to English for one it has no catalog for. */
  readonly locale: string;
};

/**
 * Why a save rewrote the whole package instead of patching the edited
 * paragraphs into the previous bytes: `structuralChange` (paragraphs added or
 * removed, or styles, headers, footers or section properties changed),
 * `untrackedChange` (an edit the change tracker cannot key to a paragraph), or
 * `noBodyView` (the body editor never started, so there was no change tracker
 * to patch from).
 */
export type FolioEditorRepackReason = "structuralChange" | "untrackedChange" | "noBodyView";

/**
 * How a serialization wrote the package, as far as the editor can tell before
 * it runs. `selective-first` patched the edited paragraphs into the previous
 * bytes unless the patch declined, in which case it fell back to a full
 * repack; folio-react does not report which of the two happened.
 */
export type FolioEditorSaveStrategy =
  | { readonly type: "selective-first" }
  | { readonly type: "full-repack"; readonly reason: FolioEditorRepackReason };

export type FolioEditorSerialization = {
  readonly bytes: Uint8Array;
  /** The `fileVersion` of the document these bytes were edited from. */
  readonly fileVersion: string;
  readonly strategy: FolioEditorSaveStrategy;
};

/** What the editor opens with, and the notifications it sends its host. */
export type FolioEditorHost = {
  readonly init: FolioEditorInit;
  /** The document is on screen, or a reload of it finished. */
  readonly onLoaded: (fileVersion: string) => void;
  /** The document could not be opened. */
  readonly onLoadFailed: (message: string) => void;
  /**
   * The user started a new undo step. Each call is one entry the host's undo
   * stack should gain; its undo and redo come back through
   * {@link FolioEditorHandle.undo} / {@link FolioEditorHandle.redo}.
   */
  readonly onEdit: () => void;
  /** Whether the editor holds edits that no serialization has captured yet. */
  readonly onDirtyChange: (dirty: boolean) => void;
  /** The user switched between editing and suggesting in the editor's toolbar. */
  readonly onModeChange: (mode: FolioEditingMode) => void;
  /** An error after the document loaded; a failed serialization rejects instead. */
  readonly onError: (message: string) => void;
};

/** Commands from the host to a mounted editor. */
export type FolioEditorHandle = {
  /** The current document as `.docx` bytes. Rejects when serialization fails. */
  readonly serialize: () => Promise<FolioEditorSerialization>;
  /** Undo one step. `false` when the editor had nothing to undo. */
  readonly undo: () => boolean;
  /** Redo one step. `false` when the editor had nothing to redo. */
  readonly redo: () => boolean;
  /** Replace the document, dropping the editor's own undo history. */
  readonly reload: (document: FolioEditorDocument) => void;
  readonly setMode: (mode: FolioEditingMode) => void;
  readonly unmount: () => void;
};
