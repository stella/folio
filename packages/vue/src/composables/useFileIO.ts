/**
 * File I/O composable for DocxEditor — owns the hidden `.docx` and image
 * file-picker refs (Insert > Image inserts directly via the shared
 * `insertImageFromFile` flow, no dialog), the document-name change emit path,
 * the `.docx` download flow, and the load/save bridge over `useDocxEditor`'s
 * `loadBuffer` / `loadParsedDocument` / `saveBlob`. Re-emits `ready` after a
 * tick so host listeners that read comments/tracked-changes on the event see
 * the freshly-extracted arrays, not stale data.
 *
 * NB: our fork's `insertImageFromFile` takes an `onInserted` callback (not an
 * options bag with `onError`) and no-ops on validation failure, so the image
 * insert path has no error emit. The upstream's `readDocxFileFromInput` core
 * util is absent from our fork, so the trivial input → { buffer, name } read is
 * inlined here.
 */

import { ref } from "vue";
import type { EditorView } from "prosemirror-view";
import { insertImageFromFile } from "@stll/folio-core/prosemirror/commands/image";
import type { Document } from "@stll/folio-core/types/document";

type DocxBufferInput = ArrayBuffer | Uint8Array | Blob | File;

interface ReadDocxResult {
  buffer: ArrayBuffer;
  name: string;
}

/** Read the first picked file from a file-input change event; resets the
 *  input so the same file can be re-selected. */
async function readDocxFileFromInput(event: Event): Promise<ReadDocxResult | null> {
  const input = event.target;
  if (!(input instanceof HTMLInputElement)) return null;
  const file = input.files?.[0];
  input.value = ""; // allow re-selecting the same file
  if (!file) return null;
  return { buffer: await file.arrayBuffer(), name: file.name };
}

export interface UseFileIOOptions {
  /** From useDocxEditor — loads a .docx buffer into the editor. */
  loadBuffer: (buffer: DocxBufferInput) => Promise<void>;
  /**
   * Host override for File > Open / Cmd+O. When provided, the picked file is
   * passed here instead of running the built-in local load — lets consumers
   * route Open through their own import pipeline (e.g. a CRDT backend).
   */
  onOpen?: (file: File) => void | Promise<void>;
  /** From useDocxEditor — loads an already-parsed Document model. */
  loadParsedDocument: (doc: Document) => void;
  /** From useDocxEditor — returns the current Document, or null. */
  getDocument: () => Document | null;
  /** From useDocxEditor — serializes the current state to a .docx Blob. */
  saveBlob: () => Promise<Blob | null>;
  /** Fired after load+extract so the next tick sees comments/tracked-changes. */
  extractCommentsAndChanges: () => void;
  /** SFC's emit function — re-emits ready / rename / update:document / error. */
  emit: (event: string, ...args: unknown[]) => void;
  /** Accessor — read freshly inside the handler so prop updates are honored. */
  documentName: () => string | undefined;
  onDocumentNameChange?: (name: string) => void;
  /** Active editor view to insert images into — the header/footer being edited, else the body. */
  getActiveView: () => EditorView | null;
  /** Vue's `nextTick` — passed in so the composable doesn't require its own import wiring. */
  nextTick: () => Promise<void>;
}

export function useFileIO(opts: UseFileIOOptions) {
  const docxInputRef = ref<HTMLInputElement | null>(null);
  const imageInputRef = ref<HTMLInputElement | null>(null);

  // Insert > Image: open the OS picker (via `imageInputRef.click()` in the menu)
  // and insert directly through the shared core flow — no dialog. Mirrors React.
  function handleImageFileChange(event: Event) {
    const input = event.target;
    if (!(input instanceof HTMLInputElement)) return;
    const file = input.files?.[0];
    const view = opts.getActiveView();
    if (file && view) {
      void insertImageFromFile(view, file, () => {});
    }
    input.value = ""; // allow re-selecting the same file
  }

  async function emitReadyAfterSidebarStateRefresh() {
    // Extract comments BEFORE emitting `ready` so host listeners that read
    // comments / tracked changes on the event see the new doc, not stale arrays.
    await opts.nextTick();
    opts.extractCommentsAndChanges();
    opts.emit("ready");
  }

  async function handleDocxFileChange(event: Event) {
    // Host-supplied open handler takes over: hand it the picked file and skip
    // the built-in local load. Mirrors React's useFileIO onOpen branch.
    if (opts.onOpen) {
      const input = event.target;
      const file = input instanceof HTMLInputElement ? input.files?.[0] : undefined;
      if (input instanceof HTMLInputElement) input.value = ""; // allow re-selecting
      if (!file) return;
      try {
        await opts.onOpen(file);
      } catch (err) {
        opts.emit("error", err instanceof Error ? err : new Error("Failed to open document"));
      }
      return;
    }

    try {
      const result = await readDocxFileFromInput(event);
      if (!result) return;
      await opts.loadBuffer(result.buffer);
      opts.emit("update:document", opts.getDocument());
      opts.emit("rename", result.name);
      await emitReadyAfterSidebarStateRefresh();
    } catch (err) {
      opts.emit("error", err instanceof Error ? err : new Error("Failed to open document"));
    }
  }

  function handleDocumentNameChange(name: string) {
    opts.onDocumentNameChange?.(name);
    opts.emit("rename", name);
  }

  /**
   * File > Save in the menu bar should produce a downloadable .docx, not
   * just stash the Blob and forget. Falls back to "document.docx" when
   * the host doesn't supply a `documentName` prop.
   */
  async function downloadCurrentDocument() {
    const blob = await opts.saveBlob();
    if (!blob) return;
    const baseName = (opts.documentName() ?? "").trim() || "document";
    const url = URL.createObjectURL(blob);
    const a = window.document.createElement("a");
    a.href = url;
    a.download = `${baseName.replace(/\.docx$/iu, "")}.docx`;
    // The anchor never enters the DOM tree — `.click()` works without
    // appending in modern browsers, and skipping the append/remove dance
    // avoids a layout flash on tall pages.
    a.click();
    // Defer revoke so Safari has time to start the download.
    setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  async function loadDocumentBuffer(buffer: DocxBufferInput) {
    await opts.loadBuffer(buffer);
    opts.emit("update:document", opts.getDocument());
    await emitReadyAfterSidebarStateRefresh();
  }

  function loadDocument(doc: Document) {
    opts.loadParsedDocument(doc);
    opts.emit("update:document", doc);
    void emitReadyAfterSidebarStateRefresh();
  }

  async function save(): Promise<ArrayBuffer | null> {
    const blob = await opts.saveBlob();
    return blob ? blob.arrayBuffer() : null;
  }

  return {
    docxInputRef,
    imageInputRef,
    handleImageFileChange,
    handleDocxFileChange,
    handleDocumentNameChange,
    downloadCurrentDocument,
    emitReadyAfterSidebarStateRefresh,
    loadDocumentBuffer,
    loadDocument,
    save,
  };
}
