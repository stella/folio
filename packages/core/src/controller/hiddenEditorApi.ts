import type { CanonicalSaveSnapshot } from "../types/canonicalSave";
/**
 * Hidden-editor imperative API
 *
 * Framework-agnostic method surface for the off-screen ProseMirror editor.
 * Operates on an injected view accessor plus the document context and
 * destruction guard, so the same imperative handle can be driven from the
 * React component or a headless controller without depending on React.
 */

import type { DocumentOp, OpStory, FormattingPatch } from "@stll/docx-core/ops";
import type { Node as PMNode } from "prosemirror-model";

import { undo, redo } from "prosemirror-history";
import type { Command, EditorState, Transaction } from "prosemirror-state";
import { NodeSelection, Selection, TextSelection } from "prosemirror-state";
import { CellSelection } from "prosemirror-tables";
import type { EditorView } from "prosemirror-view";

import { cloneDocumentWithParagraphPropertySources } from "../docx/documentClone";
import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import type { Document, SectionProperties } from "../types/document";
import type { CanonicalSectionPropertiesResult } from "../types/canonicalSections";
import type { Comment } from "../types/content";
import type {
  CanonicalWatermarkRequest,
  CanonicalWatermarkResult,
} from "../types/canonicalWatermark";
import type {
  FolioDocumentOperationResult,
  FolioDocumentOperationUndoHandle,
  FolioDocumentOperationUndoResult,
} from "../document-operations";
import type { CanonicalPublicOperationOptions } from "./canonicalPublicOperations";
import type { CanonicalCommentRequest, CanonicalCommentResult } from "../types/canonicalComments";
import type { CanonicalSessionMode } from "./canonicalSession";
import type { createCanonicalInputBoundary } from "./canonicalInput";

type CanonicalStoryHistoryOptions = {
  view: EditorView;
  story: OpStory;
  direction: "undo" | "redo";
};
type CanonicalStoryTextOptions = {
  view: EditorView;
  story: OpStory;
  intent: Parameters<Parameters<typeof createCanonicalInputBoundary>[0]["replace"]>[0];
};

export type HiddenEditorApi = {
  /** Request the off-screen EditorView (idempotent; creates it when possible). */
  ensureView: () => void;
  /** Whether view creation has been requested. */
  isViewRequested: () => boolean;
  /** Get the ProseMirror EditorState */
  getState: () => EditorState | null;
  /** Get the ProseMirror EditorView */
  getView: () => EditorView | null;
  /** Get the current Document from PM state */
  getDocument: () => Document | null;
  /** Canonical snapshot, or null in the default session. */
  getCanonicalComments: () => Comment[] | null;
  getCanonicalDocument: () => Document | null;
  captureCanonicalSave: () => CanonicalSaveSnapshot | null;
  /** A save cannot acknowledge newer edits or an unfinished composition. */
  isCanonicalSaveCurrent: (version: number) => boolean;
  /** Set the explicit mode for an active canonical session. */
  setCanonicalMode: (mode: CanonicalSessionMode) => boolean;
  /** Resolve canonical revisions as one batch and journal its exact inverse. */
  resolveCanonicalRevisions: (
    revisionIds: readonly number[],
    resolution: "accept" | "reject",
  ) => boolean;
  updateCanonicalInputLifecycle: (
    action: "beginComposition" | "endComposition" | "breakUndoGroup",
  ) => boolean;
  applyCanonicalStoryHistory: (options: CanonicalStoryHistoryOptions) => boolean;
  applyCanonicalOperations: (ops: readonly DocumentOp[]) => boolean;
  applyCanonicalComment: (request: CanonicalCommentRequest) => CanonicalCommentResult | null;
  /** Apply a final-section patch; null is reserved for default sessions. */
  applyCanonicalSectionProperties: (
    patch: FormattingPatch<SectionProperties>,
  ) => CanonicalSectionPropertiesResult | null;
  applyCanonicalWatermark: (change: CanonicalWatermarkRequest) => CanonicalWatermarkResult | null;
  getCanonicalStorySelection: (story: OpStory) => { anchor: number; head: number } | null;
  getCanonicalStoryProjection: (story: OpStory) => PMNode | null;
  replaceCanonicalStoryText: (options: CanonicalStoryTextOptions) => boolean;
  applyCanonicalDocumentOperations: (
    options: CanonicalPublicOperationOptions,
  ) => FolioDocumentOperationResult | null;
  undoCanonicalDocumentOperations: (
    handle: FolioDocumentOperationUndoHandle,
  ) => FolioDocumentOperationUndoResult | null;
  /** Focus the hidden editor */
  focus: () => void;
  /** Blur the hidden editor */
  blur: () => void;
  /** Check if focused */
  isFocused: () => boolean;
  /** Dispatch a transaction */
  dispatch: (tr: Transaction) => void;
  /** Execute a ProseMirror command */
  executeCommand: (command: Command) => boolean;
  /** Undo */
  undo: () => boolean;
  /** Redo */
  redo: () => boolean;
  /** Check if undo is available */
  canUndo: () => boolean;
  /** Check if redo is available */
  canRedo: () => boolean;
  /** Set selection by PM position */
  setSelection: (anchor: number, head?: number) => void;
  /** Set node selection at a PM position (for images, etc.) */
  setNodeSelection: (pos: number) => void;
  /** Set cell selection between two positions inside table cells */
  setCellSelection: (anchorCellPos: number, headCellPos: number) => void;
  /** Scroll the PM view to selection (no-op since hidden) */
  scrollToSelection: () => void;
};

export type HiddenEditorApiDeps = {
  getView: () => EditorView | null;
  getDocumentContext: () => Document | null;
  getCanonicalComments?: () => Comment[] | null;
  getCanonicalDocument?: () => Document | null;
  captureCanonicalSave?: HiddenEditorApi["captureCanonicalSave"];
  isCanonicalSaveCurrent?: HiddenEditorApi["isCanonicalSaveCurrent"];
  setCanonicalMode?: HiddenEditorApi["setCanonicalMode"];
  resolveCanonicalRevisions?: HiddenEditorApi["resolveCanonicalRevisions"];
  executeCanonicalCommand?: (command: Command) => boolean | undefined;
  getCanonicalHistory?: () => Pick<HiddenEditorApi, "undo" | "redo" | "canUndo" | "canRedo"> | null;
  canonicalOperations?: Pick<
    HiddenEditorApi,
    | "updateCanonicalInputLifecycle"
    | "applyCanonicalOperations"
    | "applyCanonicalComment"
    | "applyCanonicalSectionProperties"
    | "applyCanonicalWatermark"
    | "getCanonicalStoryProjection"
    | "replaceCanonicalStoryText"
    | "applyCanonicalStoryHistory"
    | "getCanonicalStorySelection"
    | "applyCanonicalDocumentOperations"
    | "undoCanonicalDocumentOperations"
  >;
  isDestroying: () => boolean;
  ensureView: () => void;
  isViewRequested: () => boolean;
};

/**
 * Convert PM state to Document
 */
const stateToDocument = (state: EditorState, originalDoc: Document | null): Document | null => {
  if (!originalDoc) {
    return null;
  }

  // fromProseDoc preserves the base document structure when provided
  // canonical-gap: pm-save-projection
  return fromProseDoc(state.doc, originalDoc);
};

export const createHiddenEditorApi = (deps: HiddenEditorApiDeps): HiddenEditorApi => {
  // Skip dispatching while the view is being destroyed. Every dispatching
  // method routes through this so the framework-agnostic API is self-sufficient
  // rather than relying on the host view's dispatchTransaction backstop.
  const dispatchTr = (view: EditorView, tr: Transaction): void => {
    if (!deps.isDestroying()) {
      view.dispatch(tr);
    }
  };

  const setSelection = (anchor: number, head?: number): void => {
    const view = deps.getView();
    if (!view) {
      return;
    }
    const { state } = view;
    const docEnd = state.doc.content.size;
    const clampedAnchor = Math.max(0, Math.min(anchor, docEnd));
    const clampedHead = head === undefined ? clampedAnchor : Math.max(0, Math.min(head, docEnd));
    const $anchor = state.doc.resolve(clampedAnchor);
    const $head = state.doc.resolve(clampedHead);
    const selection =
      head === undefined ? Selection.near($anchor) : TextSelection.between($anchor, $head);
    dispatchTr(view, state.tr.setSelection(selection));
  };

  return {
    ensureView: deps.ensureView,

    isViewRequested: deps.isViewRequested,

    getState: () => deps.getView()?.state ?? null,

    getView: () => deps.getView() ?? null,

    setCanonicalMode: (mode) => !deps.isDestroying() && (deps.setCanonicalMode?.(mode) ?? false),

    resolveCanonicalRevisions: (revisionIds, resolution) =>
      !deps.isDestroying() && (deps.resolveCanonicalRevisions?.(revisionIds, resolution) ?? false),

    getCanonicalComments: () => deps.getCanonicalComments?.() ?? null,
    captureCanonicalSave: () => deps.captureCanonicalSave?.() ?? null,
    isCanonicalSaveCurrent: (version) => deps.isCanonicalSaveCurrent?.(version) ?? false,

    getCanonicalDocument: () => {
      const canonical = deps.getCanonicalDocument?.();
      return canonical ? cloneDocumentWithParagraphPropertySources(canonical) : null;
    },

    getDocument: () => {
      const canonical = deps.getCanonicalDocument?.();
      if (canonical) return cloneDocumentWithParagraphPropertySources(canonical);
      const view = deps.getView();
      if (!view) {
        return null;
      }
      return stateToDocument(view.state, deps.getDocumentContext());
    },

    updateCanonicalInputLifecycle: (action) =>
      deps.canonicalOperations?.updateCanonicalInputLifecycle(action) ?? false,
    applyCanonicalStoryHistory: (options) =>
      deps.canonicalOperations?.applyCanonicalStoryHistory(options) ?? false,
    applyCanonicalComment: (request) =>
      deps.canonicalOperations?.applyCanonicalComment(request) ?? null,
    applyCanonicalSectionProperties: (patch) =>
      deps.canonicalOperations?.applyCanonicalSectionProperties(patch) ?? null,
    applyCanonicalWatermark: (change) =>
      deps.canonicalOperations?.applyCanonicalWatermark(change) ?? null,
    applyCanonicalOperations: (ops) =>
      deps.canonicalOperations?.applyCanonicalOperations(ops) ?? false,
    getCanonicalStorySelection: (story) =>
      deps.canonicalOperations?.getCanonicalStorySelection(story) ?? null,
    getCanonicalStoryProjection: (story) =>
      deps.canonicalOperations?.getCanonicalStoryProjection(story) ?? null,
    replaceCanonicalStoryText: (options) =>
      deps.canonicalOperations?.replaceCanonicalStoryText(options) ?? false,

    applyCanonicalDocumentOperations: (options) =>
      deps.canonicalOperations?.applyCanonicalDocumentOperations(options) ?? null,
    undoCanonicalDocumentOperations: (handle) =>
      deps.canonicalOperations?.undoCanonicalDocumentOperations(handle) ?? null,

    focus: () => {
      const view = deps.getView();
      if (!view || view.hasFocus()) {
        return;
      }
      view.focus();
    },

    blur: () => {
      const view = deps.getView();
      const dom = view?.dom;
      if (view?.hasFocus() && dom instanceof HTMLElement) {
        dom.blur();
      }
    },

    isFocused: () => deps.getView()?.hasFocus() ?? false,

    dispatch: (tr: Transaction) => {
      const view = deps.getView();
      if (view) {
        dispatchTr(view, tr);
      }
    },

    executeCommand: (command: Command) => {
      const view = deps.getView();
      if (!view) {
        return false;
      }
      const canonical = deps.executeCanonicalCommand?.(command);
      if (canonical !== undefined) return canonical;
      return command(view.state, (tr) => dispatchTr(view, tr), view);
    },

    undo: () => {
      const canonical = deps.getCanonicalHistory?.();
      if (canonical) return canonical.undo();
      const view = deps.getView();
      if (!view) {
        return false;
      }
      return undo(view.state, (tr) => dispatchTr(view, tr));
    },

    redo: () => {
      const canonical = deps.getCanonicalHistory?.();
      if (canonical) return canonical.redo();
      const view = deps.getView();
      if (!view) {
        return false;
      }
      return redo(view.state, (tr) => dispatchTr(view, tr));
    },

    canUndo: () => {
      const canonical = deps.getCanonicalHistory?.();
      if (canonical) return canonical.canUndo();
      const view = deps.getView();
      if (!view) {
        return false;
      }
      return undo(view.state);
    },

    canRedo: () => {
      const canonical = deps.getCanonicalHistory?.();
      if (canonical) return canonical.canRedo();
      const view = deps.getView();
      if (!view) {
        return false;
      }
      return redo(view.state);
    },

    setSelection,

    setNodeSelection: (pos: number) => {
      const view = deps.getView();
      if (!view) {
        return;
      }
      const { state } = view;
      try {
        const selection = NodeSelection.create(state.doc, pos);
        dispatchTr(view, state.tr.setSelection(selection));
      } catch {
        // Fallback to text selection if NodeSelection fails
        setSelection(pos);
      }
    },

    setCellSelection: (anchorCellPos: number, headCellPos: number) => {
      const view = deps.getView();
      if (!view) {
        return;
      }
      const { state } = view;
      try {
        const cellSel = CellSelection.create(state.doc, anchorCellPos, headCellPos);
        dispatchTr(view, state.tr.setSelection(cellSel));
      } catch {
        // Fallback to text selection if positions aren't valid for CellSelection
        setSelection(anchorCellPos, headCellPos);
      }
    },

    scrollToSelection: () => {
      // No-op for hidden editor - visual scrolling handled by PagedEditor
    },
  };
};
