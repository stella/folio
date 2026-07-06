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
 * `getEditorRef()` synthesizes a `PagedEditorRef`-shaped handle from the same
 * primitives (the Vue package has no ported `PagedEditor` component to source
 * one from). Most of its surface maps straight onto the headless `FolioEditor`
 * controller (`opts.editor`), which already implements the hidden-editor
 * imperative API 1:1 with React's `PagedEditorRef`; the scroll/page methods
 * reuse this file's own `scrollToPage` / `scrollToParaId` /
 * `scrollVisiblePositionIntoView`. `getHfView` stays a documented no-op:
 * `useDocxEditor`'s layout pipeline always passes `hfPMs: null`, so there is no
 * persistent hidden header/footer `EditorView` to look up by `rId` yet
 * (PORT-BLOCKED, see that method's note). The AI-edit and content-control
 * methods are wired over folio-core directly, mirroring the React adapter.
 * `hasPendingChanges` is wired from the doc-dirty flag (`useDocxEditor`) OR-ed
 * with the comment-list dirty flag (`useCommentManagement`); see its
 * per-method note. The assembled object `satisfies DocxEditorRef`.
 */

import type { Ref } from "vue";

import { TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
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
import type { PagedEditorRef } from "../components/DocxEditor/pagedEditorRef";
import { clampRangeToDocSize, resolveFolioAIBlockRange } from "../utils/aiEditRange";

export type UseDocxEditorRefApiOptions = {
  /** Headless controller handle (imperative API + events; Seam 6). */
  editor: FolioEditor;
  /** Off-screen ProseMirror EditorView, or null before mount. */
  editorView: Ref<EditorView | null>;
  /**
   * Doc-dirty flag from useDocxEditor: live PM edits not yet serialized. Backs
   * the doc side of `hasPendingChanges` (React's ParagraphChangeTracker signals).
   */
  isDirty: Readonly<Ref<boolean>>;
  /**
   * Comment-list dirty flag from useCommentManagement: comment mutations not yet
   * serialized. Backs the comment side of `hasPendingChanges` (React's
   * `commentsDirtyRef`).
   */
  commentsDirty: Readonly<Ref<boolean>>;
  /**
   * Clear {@link commentsDirty} after a successful save. Called from the `save`
   * wrapper below — the single public save entry point — so the comment flag is
   * reset alongside useDocxEditor's doc-dirty flag, mirroring React's handleSave.
   */
  clearCommentsDirty: () => void;
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
};

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
    // Save succeeded: clear the comment-list dirty flag (useDocxEditor already
    // cleared the doc-dirty flag when it produced the bytes). Mirrors React's
    // handleSave resetting commentsDirtyRef.
    opts.clearCommentsDirty();
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

  // Synthesized from the same primitives useDocxEditor/useDocxEditorRefApi
  // already hold — see the module docblock. Every method below reads through
  // `opts`/`layout.value` at call time, so the assembled handle itself is
  // static; cache it instead of re-allocating the object and its 20+ closures
  // on every getEditorRef() call. Always returns a live handle (the Vue shell
  // has no "editor not yet mounted" state distinct from a null editorView,
  // which every method below already guards against).
  let pagedEditorRef: PagedEditorRef | null = null;

  function getEditorRef(): PagedEditorRef {
    pagedEditorRef ??= {
      getEditor: () => opts.editor,
      getDocument: () => opts.editor.getDocument(),
      getState: () => opts.editor.getState(),
      getView: () => opts.editor.getView(),
      // PORT-BLOCKED: no persistent hidden header/footer EditorView to look up
      // by rId yet. useDocxEditor's layout pipeline always calls
      // runLayoutPipelineCompute with `hfPMs: null` (see HfPmsHandle there),
      // so header/footer content is rendered straight from the document model
      // rather than a live PM view. Returns null until that lands.
      getHfView: () => null,
      // The fork's controller ensureView() takes no focus argument yet; the
      // { focus } option is accepted for React parity but ignored, matching
      // ensureEditorView above (PORT-BLOCKED).
      ensureView: () => opts.editor.ensureView(),
      focus: () => opts.editor.focus(),
      blur: () => opts.editor.blur(),
      isFocused: () => opts.editor.isFocused(),
      dispatch: (tr: Transaction) => opts.editor.dispatch(tr),
      undo: () => opts.editor.undo(),
      redo: () => opts.editor.redo(),
      canUndo: () => opts.editor.canUndo(),
      canRedo: () => opts.editor.canRedo(),
      setSelection: (anchor: number, head?: number) => opts.editor.setSelection(anchor, head),
      getLayout: () => opts.editor.getLayout(),
      relayout: () => opts.editor.relayout(),
      scrollToPosition: (pmPos: number) => opts.scrollVisiblePositionIntoView(pmPos),
      scrollToPage,
      scrollToParaId,
      getPageNumberForPmPos: (pmPos: number) => {
        const currentLayout = opts.layout.value;
        if (!currentLayout) {
          return null;
        }
        const pageIndex = findPageIndexContainingPmPos(currentLayout, pmPos);
        return pageIndex == null ? null : pageIndex + 1;
      },
    };
    return pagedEditorRef;
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
    // React derives this live from the ParagraphChangeTracker plugin
    // (getChangedParagraphIds/hasStructuralChanges/hasUntrackedChanges) OR-ed
    // with commentsDirtyRef, clearing the tracker on save. Vue keeps the tracker
    // uncleared across saves (it is the selective-save baseline), so instead of
    // reading the tracker we track two flags with the same set/clear points as
    // React: `isDirty` (any doc-changing transaction; cleared on save + load) and
    // `commentsDirty` (any comment mutation; cleared on save + load). Same net
    // semantics: pending edits after typing/comment changes, clean after save.
    hasPendingChanges: () => {
      if (!opts.editorView.value) {
        return false;
      }
      return opts.isDirty.value || opts.commentsDirty.value;
    },
    getEditorRef,
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
