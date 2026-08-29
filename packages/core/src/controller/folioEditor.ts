// The composition half of the headless editor controller (seam-architecture
// Seam 6: `FolioEditor`). Unifies the hidden-editor imperative API, layout
// access, and the event emitter into one framework-agnostic object that
// framework adapters and desktop/headless hosts drive. Kept in `core/` with no
// React or `paged-editor/*` dependency so it stays portable across hosts.

import { panic } from "better-result";
import type { EditorState } from "prosemirror-state";

import type { Layout } from "../layout-engine/types";
import type { Document } from "../types/document";
import type { DocxInput } from "../utils/docxInput";
import type { FolioEditorEmitter } from "./folioEditorEvents";
import type { HiddenEditorApi } from "./hiddenEditorApi";
import type { LayoutRunOptions } from "./layoutScheduler";

export type FolioEditorDeps = {
  // The live hidden-editor API; null until the view exists (the adapter holds
  // it in a ref), so it's read fresh on every call.
  getEditorApi: () => HiddenEditorApi | null;
  getLayout: () => Layout | null;
  runLayout: (state: EditorState, options: LayoutRunOptions) => void;
  /** Required only when the document I/O methods are used. */
  getDocumentIO?: () => FolioEditorDocumentIO;
  emitter: FolioEditorEmitter;
};

export const FOLIO_DOCX_SERIALIZATION_MODE = Object.freeze({
  full: "full",
  preferSelective: "prefer-selective",
} as const);

export type FolioDocxSerializationMode =
  (typeof FOLIO_DOCX_SERIALIZATION_MODE)[keyof typeof FOLIO_DOCX_SERIALIZATION_MODE];

export type FolioGetDocxOptions = {
  mode?: FolioDocxSerializationMode;
};

/**
 * Document lifecycle operations supplied by a host integration.
 *
 * The controller owns the stable contract while each adapter supplies its
 * current persistence implementation. This lets document I/O migrate without
 * coupling core to React, Vue, or host callbacks.
 */
export type FolioEditorDocumentIO = {
  getDocx: (options?: FolioGetDocxOptions) => Promise<ArrayBuffer | null>;
  loadDocument: (document: Document) => void;
  loadDocx: (input: DocxInput) => Promise<void>;
};

// The headless controller surface (Seam 6). The HiddenEditorApi methods plus
// layout access and event subscription.
export type FolioEditor = HiddenEditorApi & {
  getDocx: FolioEditorDocumentIO["getDocx"];
  loadDocument: FolioEditorDocumentIO["loadDocument"];
  loadDocx: FolioEditorDocumentIO["loadDocx"];
  getLayout: () => Layout | null;
  /** Re-run layout for the current editor state (no-op if there is no view). */
  relayout: () => void;
  on: FolioEditorEmitter["on"];
};

export const createFolioEditor = (deps: FolioEditorDeps): FolioEditor => {
  const getDocumentIO = () =>
    deps.getDocumentIO?.() ?? panic("FolioEditor document I/O requires a getDocumentIO dependency");

  return {
    getDocx: (options) => getDocumentIO().getDocx(options),

    loadDocument: (document) => getDocumentIO().loadDocument(document),

    loadDocx: (input) => getDocumentIO().loadDocx(input),

    ensureView: () => deps.getEditorApi()?.ensureView(),

    isViewRequested: () => deps.getEditorApi()?.isViewRequested() ?? false,

    getState: () => deps.getEditorApi()?.getState() ?? null,

    getView: () => deps.getEditorApi()?.getView() ?? null,

    getDocument: () => deps.getEditorApi()?.getDocument() ?? null,

    focus: () => deps.getEditorApi()?.focus(),

    blur: () => deps.getEditorApi()?.blur(),

    isFocused: () => deps.getEditorApi()?.isFocused() ?? false,

    dispatch: (tr) => deps.getEditorApi()?.dispatch(tr),

    executeCommand: (command) => deps.getEditorApi()?.executeCommand(command) ?? false,

    undo: () => deps.getEditorApi()?.undo() ?? false,

    redo: () => deps.getEditorApi()?.redo() ?? false,

    canUndo: () => deps.getEditorApi()?.canUndo() ?? false,

    canRedo: () => deps.getEditorApi()?.canRedo() ?? false,

    setSelection: (anchor, head) => deps.getEditorApi()?.setSelection(anchor, head),

    setNodeSelection: (pos) => deps.getEditorApi()?.setNodeSelection(pos),

    setCellSelection: (anchorCellPos, headCellPos) =>
      deps.getEditorApi()?.setCellSelection(anchorCellPos, headCellPos),

    scrollToSelection: () => deps.getEditorApi()?.scrollToSelection(),

    getLayout: () => deps.getLayout(),

    relayout: () => {
      const state = deps.getEditorApi()?.getState();
      if (state) {
        deps.runLayout(state, { reason: "manual" });
      }
    },

    // Wrap rather than tear off the method, so the contract holds for any emitter
    // implementation (e.g. a class-based one whose `on` relies on `this`).
    on: (event, listener) => deps.emitter.on(event, listener),
  };
};
