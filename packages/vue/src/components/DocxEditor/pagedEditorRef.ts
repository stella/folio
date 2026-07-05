import type { EditorState, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import type { Layout } from "@stll/folio-core/layout-engine";
import type { ScrollToParaIdOptions } from "@stll/folio-core/paged-layout/paragraphFlash";
import type { Document } from "@stll/folio-core/types/document";

/**
 * Imperative handle exposed by the Vue paged editor sub-component.
 *
 * Structurally mirrors the React package's `PagedEditorRef`
 * (`packages/react/src/paged-editor/PagedEditor.tsx`). The Vue PagedEditor
 * component itself is not ported yet; this type defines the contract so
 * `DocxEditorRef.getEditorRef()` can return it.
 *
 * TODO(vue): re-home to the Vue PagedEditor component when it lands.
 */
export type PagedEditorRef = {
  /** The headless editor controller (imperative API + events; Seam 6). */
  getEditor: () => FolioEditor;
  /** Get the current document. */
  getDocument: () => Document | null;
  /** Get the ProseMirror EditorState. */
  getState: () => EditorState | null;
  /** Get the ProseMirror EditorView. */
  getView: () => EditorView | null;
  /**
   * Look up the persistent hidden HF EditorView by `rId`. Returns null when the
   * slot isn't mounted (no HF for that rId, or the hidden host hasn't mounted).
   */
  getHfView: (rId: string) => EditorView | null;
  /**
   * Force-create the hidden editor view if it has been deferred. Use from
   * surfaces that need a live view before any user interaction.
   */
  ensureView: (options?: { focus?: boolean }) => void;
  /** Focus the editor. */
  focus: () => void;
  /** Blur the editor. */
  blur: () => void;
  /** Check if focused. */
  isFocused: () => boolean;
  /** Dispatch a transaction. */
  dispatch: (tr: Transaction) => void;
  /** Undo. */
  undo: () => boolean;
  /** Redo. */
  redo: () => boolean;
  /** Check whether undo is available. */
  canUndo: () => boolean;
  /** Check whether redo is available. */
  canRedo: () => boolean;
  /** Set selection by PM position. */
  setSelection: (anchor: number, head?: number) => void;
  /** Get current layout. */
  getLayout: () => Layout | null;
  /** Force re-layout. */
  relayout: () => void;
  /** Scroll the visible pages to bring a PM position into view. */
  scrollToPosition: (pmPos: number) => void;
  /** Scroll the visible pages to bring a page into view. */
  scrollToPage: (pageNumber: number) => void;
  /**
   * Scroll the paginated view to the paragraph with the given Word `w14:paraId`.
   * Returns whether a matching paragraph exists in the ProseMirror document.
   */
  scrollToParaId: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
  /**
   * Resolve the page number (1-indexed) that contains the given PM position, or
   * null if no layout is available yet. Works for unrendered pages too via the
   * page shell map.
   */
  getPageNumberForPmPos: (pmPos: number) => number | null;
};
