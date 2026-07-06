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
 * PORT-BLOCKED: a few ref methods still depend on adapter surfaces the fork has
 * not ported yet (`hasPendingChanges` needs PM-serialized-vs-live tracking;
 * `getEditorRef` needs the Vue PagedEditor component). Those stay stubbed to a
 * sensible default (null / false) with a per-method note. The AI-edit and
 * content-control methods are wired over folio-core directly, mirroring the
 * React adapter. The assembled object `satisfies DocxEditorRef`.
 */

import type { Ref } from "vue";

import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { applyFolioAIEditOperations, createFolioAIEditSnapshot } from "@stll/folio-core/ai-edits";
import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import type { Layout } from "@stll/folio-core/layout-engine";
import { findPageIndexContainingPmPos } from "@stll/folio-core/layout-engine";
import {
  findParagraphFragmentsByParaId,
  flashParagraphElements,
} from "@stll/folio-core/paged-layout/paragraphFlash";
import type { ScrollToParaIdOptions } from "@stll/folio-core/paged-layout/paragraphFlash";
import {
  acceptAIEditRevision,
  findAIEditRevisionRange,
  rejectAIEditRevision,
} from "@stll/folio-core/prosemirror/commands/comments";
import {
  blockSdtAttrsToSdtProperties,
  findBlockSdtMatch,
  findBlockSdtMatches,
  removeContentControlTr,
  setContentControlContentTr,
  setContentControlValueTr,
} from "@stll/folio-core/prosemirror/commands/contentControls";
import { setContentControlContentBlocksTr } from "@stll/folio-core/prosemirror/commands/contentControlsBlockFill";
import type { Document } from "@stll/folio-core/types/document";
import type { DocxInput } from "@stll/folio-core/utils/docxInput";

import type { DocxEditorRef } from "../components/DocxEditor/types";
import { clampRangeToDocSize, resolveFolioAIBlockRange } from "../utils/aiEditRange";

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
  /**
   * Scroll the pages viewport so a PM document position is visible (from
   * usePagesPointer). Backs the AI-edit / content-control scroll methods, the
   * Vue analogue of React's `PagedEditorRef.scrollToPosition`.
   */
  scrollVisiblePositionIntoView: (pmPos: number) => void;
  /** Editor author, used as the default operation author for AI edits. */
  author: () => string;
  /**
   * Mint a comment for an AI-edit operation that carries comment text, append
   * it to the thread list, and return its id (mirrors React's `createCommentId`
   * closure over the comment manager). Wired from useCommentManagement.
   */
  createAIEditComment: (text: string) => number;
  // Action handles from useDocxEditor.
  focus: () => void;
  getDocument: () => Document | null;
  setZoom: (zoom: number) => void;
  /** useDocxEditor.save returns a Blob; the ref surface exposes an ArrayBuffer. */
  save: (options?: { selective?: boolean }) => Promise<Blob | null>;
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

  async function save(options?: { selective?: boolean }): Promise<ArrayBuffer | null> {
    const blob = await opts.save(options);
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
    // Threads `options.selective` into useDocxEditor.save (selective-save gate).
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
    createAIEditSnapshot: () => {
      const view = opts.editorView.value;
      return view ? createFolioAIEditSnapshot(view.state.doc) : null;
    },
    applyAIEditOperations: ({
      snapshot,
      operations,
      mode = "tracked-changes",
      author: operationAuthor = opts.author(),
    }) => {
      const view = opts.editorView.value;
      if (!view) {
        return {
          applied: [],
          skipped: operations.map((operation) => ({
            id: operation.id,
            reason: "unsupportedBlock",
          })),
        };
      }
      return applyFolioAIEditOperations({
        view,
        snapshot,
        operations,
        mode,
        author: operationAuthor,
        createCommentId: opts.createAIEditComment,
      });
    },
    acceptAIEditOperation: (revisionIds) => {
      const view = opts.editorView.value;
      if (!view) {
        return false;
      }
      return acceptAIEditRevision(revisionIds)(view.state, view.dispatch);
    },
    rejectAIEditOperation: (revisionIds) => {
      const view = opts.editorView.value;
      if (!view) {
        return false;
      }
      return rejectAIEditRevision(revisionIds)(view.state, view.dispatch);
    },
    undo: () => opts.editor.undo(),
    redo: () => opts.editor.redo(),
    scrollToAIEditOperation: (revisionIds) => {
      const view = opts.editorView.value;
      if (!view) {
        return false;
      }
      const range = findAIEditRevisionRange(view.state, revisionIds);
      if (!range) {
        return false;
      }
      // `TextSelection.between` clamps to the nearest valid inline position;
      // `create` throws when from/to land on a block boundary, which is exactly
      // what happens for marks that wrap an entire paragraph.
      // `clampRangeToDocSize` guards against a revision range whose endpoints
      // fell past the doc end after concurrent edits.
      const { from, to } = clampRangeToDocSize(view.state.doc.content.size, range);
      const $from = view.state.doc.resolve(from);
      const $to = view.state.doc.resolve(to);
      view.dispatch(view.state.tr.setSelection(TextSelection.between($from, $to)));
      requestAnimationFrame(() => opts.scrollVisiblePositionIntoView(from));
      return true;
    },
    scrollToBlock: (blockId, snapshot) => {
      const view = opts.editorView.value;
      if (!view) {
        return false;
      }
      // ParaId-backed ids resolve against the live document so queued
      // suggestions still navigate correctly after earlier accepts insert or
      // delete paragraphs above them. `seq-*` fallback ids keep using the
      // snapshot the AI saw.
      const range = resolveFolioAIBlockRange({ blockId, doc: view.state.doc, snapshot });
      if (range === null) {
        return false;
      }
      const { from, to } = range;
      const $from = view.state.doc.resolve(from);
      const $to = view.state.doc.resolve(to);
      view.dispatch(view.state.tr.setSelection(TextSelection.between($from, $to)));
      requestAnimationFrame(() => opts.scrollVisiblePositionIntoView(from));
      return true;
    },
    getContentControls: (filter = {}) => {
      const view = opts.editorView.value;
      if (!view) {
        return [];
      }
      return findBlockSdtMatches(view.state.doc, filter).map((match) => ({
        properties: blockSdtAttrsToSdtProperties(match.node),
        path: match.path,
        pmPos: match.pos,
      }));
    },
    scrollToContentControl: (filter) => {
      const view = opts.editorView.value;
      if (!view) {
        return false;
      }
      const match = findBlockSdtMatch(view.state.doc, filter);
      if (!match) {
        return false;
      }
      // Place selection just inside the SDT (after its opening token).
      const inside = match.pos + 1;
      const $pos = view.state.doc.resolve(inside);
      view.dispatch(view.state.tr.setSelection(TextSelection.near($pos)));
      requestAnimationFrame(() => opts.scrollVisiblePositionIntoView(inside));
      return true;
    },
    setContentControlContent: (filter, input, options = {}) => {
      const view = opts.editorView.value;
      if (!view) {
        return false;
      }
      const tr =
        typeof input === "string"
          ? setContentControlContentTr(view.state, filter, input, options)
          : setContentControlContentBlocksTr(view.state, filter, input, options);
      if (!tr) {
        return false;
      }
      view.dispatch(tr);
      return true;
    },
    setContentControlValue: (filter, input, options = {}) => {
      const view = opts.editorView.value;
      if (!view) {
        return false;
      }
      const tr = setContentControlValueTr(view.state, filter, input, options);
      if (!tr) {
        return false;
      }
      view.dispatch(tr);
      return true;
    },
    removeContentControl: (filter, options = {}) => {
      const view = opts.editorView.value;
      if (!view) {
        return false;
      }
      const tr = removeContentControlTr(view.state, filter, options);
      if (!tr) {
        return false;
      }
      view.dispatch(tr);
      return true;
    },
  } satisfies DocxEditorRef;

  return { exposed };
}
