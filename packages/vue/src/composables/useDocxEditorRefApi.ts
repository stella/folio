/**
 * Ref-API assembler for the Vue `DocxEditor` shell.
 *
 * Takes the foundational primitives from `useDocxEditor` (the headless
 * `FolioEditor` controller plus reactive `editorView` / `layout` / `zoom`) and
 * returns the `DocxEditorRef`-shaped `exposed` object that `DocxEditor.vue`
 * hands to `defineExpose`. This mirrors the React package's imperative editor
 * ref: `satisfies DocxEditorRef` enforces the surface at composable-build time,
 * so any drift from the React contract is a typecheck error here rather than a
 * runtime surprise at the call site.
 *
 * PORT-BLOCKED: several ref methods depend on adapter surfaces the fork has not
 * ported yet (comment/AI-edit/content-control composables, the Vue PagedEditor,
 * a `scrollVisiblePositionIntoView` from a pages-pointer composable). Those are
 * stubbed to a sensible default (null / false / empty) with a per-method note.
 * The assembled object still `satisfies DocxEditorRef`.
 */

import type { Ref } from "vue";

import type { EditorView } from "prosemirror-view";

import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import type { Layout } from "@stll/folio-core/layout-engine";
import { findPageIndexContainingPmPos } from "@stll/folio-core/layout-engine";
import {
  findParagraphFragmentsByParaId,
  flashParagraphElements,
} from "@stll/folio-core/paged-layout/paragraphFlash";
import type { ScrollToParaIdOptions } from "@stll/folio-core/paged-layout/paragraphFlash";
import type { Document } from "@stll/folio-core/types/document";
import type { DocxInput } from "@stll/folio-core/utils/docxInput";

import type { DocxEditorRef } from "../components/DocxEditor/types";

export type UseDocxEditorRefApiOptions = {
  /** Headless controller handle (imperative API + events; Seam 6). */
  editor: FolioEditor;
  /** Off-screen ProseMirror EditorView, or null before mount. */
  editorView: Ref<EditorView | null>;
  /** Computed page layout, or null before first paint. */
  layout: Ref<Layout | null>;
  /** Painted pages container (the `HTMLDivElement` the painter writes into). */
  pagesRef: Ref<HTMLElement | null>;
  /** Scrolling viewport wrapping the pages container. */
  pagesViewportRef: Ref<HTMLElement | null>;
  /** Current zoom factor (1 = 100%). */
  zoom: Ref<number>;
  // Action handles from useDocxEditor.
  focus: () => void;
  getDocument: () => Document | null;
  setZoom: (zoom: number) => void;
  /** useDocxEditor.save returns a Blob; the ref surface exposes an ArrayBuffer. */
  save: () => Promise<Blob | null>;
  loadDocument: (doc: Document) => void;
  loadDocumentBuffer: (buffer: DocxInput) => Promise<void>;
  /** Optional host hook for print. */
  onPrint?: (() => void) | undefined;
  /** Optional host hook fired with the serialized `.docx` bytes after `save()`. */
  onSave?: ((buffer: ArrayBuffer) => void) | undefined;
}

export function useDocxEditorRefApi(opts: UseDocxEditorRefApiOptions): {
  exposed: DocxEditorRef;
} {
  function print(): void {
    opts.onPrint?.();
    window.print();
  }

  async function save(): Promise<ArrayBuffer | null> {
    const blob = await opts.save();
    if (!blob) {
      return null;
    }
    const buffer = await blob.arrayBuffer();
    // Mirror React's handleSave: notify the host with the serialized bytes.
    opts.onSave?.(buffer);
    return buffer;
  }

  function scrollToParaId(paraId: string, options?: ScrollToParaIdOptions): boolean {
    const root = opts.pagesRef.value;
    if (!root) {
      return false;
    }
    const fragments = findParagraphFragmentsByParaId(root, paraId);
    const first = fragments.at(0);
    if (!first) {
      return false;
    }
    first.scrollIntoView({ block: "center", behavior: "smooth" });
    if (options?.highlight) {
      flashParagraphElements(fragments, options.highlight);
    }
    return true;
  }

  function getZoom(): number {
    return opts.zoom.value;
  }

  function getTotalPages(): number {
    return opts.layout.value?.pages.length ?? 0;
  }

  function getCurrentPage(): number {
    const currentLayout = opts.layout.value;
    const view = opts.editorView.value;
    if (!currentLayout || !view) {
      return 0;
    }
    const pageIndex = findPageIndexContainingPmPos(currentLayout, view.state.selection.from);
    return pageIndex == null ? 0 : pageIndex + 1;
  }

  function scrollToPage(pageNumber: number): void {
    if (!Number.isInteger(pageNumber) || pageNumber < 1) {
      return;
    }
    const viewport = opts.pagesViewportRef.value;
    const pageEl = opts.pagesRef.value?.querySelector<HTMLElement>(
      `[data-page-number="${pageNumber}"]`,
    );
    if (!viewport || !pageEl) {
      return;
    }
    const viewportRect = viewport.getBoundingClientRect();
    const pageRect = pageEl.getBoundingClientRect();
    viewport.scrollTo({
      top: pageRect.top - viewportRect.top + viewport.scrollTop - 24,
      behavior: "smooth",
    });
  }

  const exposed = {
    getDocument: opts.getDocument,
    // PORT-BLOCKED: hasPendingChanges needs the PM-serialized-vs-live tracking
    // the React shell keeps; useDocxEditor does not surface it yet.
    hasPendingChanges: () => false,
    // PORT-BLOCKED: getEditorRef needs the Vue PagedEditor component, not ported.
    getEditorRef: () => null,
    getEditor: () => opts.editor,
    // options.selective is ignored until selective save is wired (PORT-BLOCKED).
    save,
    setZoom: opts.setZoom,
    getZoom,
    focus: opts.focus,
    getCurrentPage,
    getTotalPages,
    scrollToPage,
    scrollToParaId,
    openPrintPreview: () => print(),
    print,
    loadDocument: opts.loadDocument,
    loadDocumentBuffer: opts.loadDocumentBuffer,
    // The fork's controller ensureView() takes no focus argument yet; the
    // { focus } option is accepted for React parity but ignored (PORT-BLOCKED).
    ensureEditorView: () => opts.editor.ensureView(),
    // PORT-BLOCKED: AI-edit surface (snapshot/apply/accept/reject/scroll) needs
    // the folio-core ai-edits bridge wired through a composable; not ported.
    createAIEditSnapshot: () => null,
    applyAIEditOperations: () => ({ applied: [], skipped: [] }),
    acceptAIEditOperation: () => false,
    rejectAIEditOperation: () => false,
    undo: () => opts.editor.undo(),
    redo: () => opts.editor.redo(),
    scrollToAIEditOperation: () => false,
    scrollToBlock: () => false,
    // PORT-BLOCKED: content-control ref methods need transaction wrappers over
    // folio-core/content-controls; not ported.
    getContentControls: () => [],
    scrollToContentControl: () => false,
    setContentControlContent: () => false,
    setContentControlValue: () => false,
    removeContentControl: () => false,
  } satisfies DocxEditorRef;

  return { exposed };
}
