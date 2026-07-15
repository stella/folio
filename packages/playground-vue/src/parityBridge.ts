import { insertTableInView, setAnonymizationTermsMeta } from "@stll/folio-vue";
import type { DocxEditorRef } from "@stll/folio-vue";

/**
 * Cross-adapter parity bridge (Vue). Kept identical to the React playground's
 * `buildParityBridge`; every method is built only on the `DocxEditorRef` members
 * classified `paired` in scripts/parity/parity.contract.json, so the same test
 * body drives both adapters. See tests/parity/parity-fixture.ts.
 */
export type FolioParityBridge = {
  /** Total laid-out pages (0 before the first layout). */
  getTotalPages: () => number;
  /** Force-create the deferred editor view (no focus steal). */
  ensureView: () => void;
  /** Whether the live ProseMirror view exists yet. */
  hasView: () => boolean;
  /** Concatenated document text (block separators collapse to empty). */
  getDocumentText: () => string;
  /** Insert text at the current selection. Returns false with no live view. */
  insertText: (text: string) => boolean;
  /** Bold the first word of the document. Returns whether the mark applied. */
  boldFirstWord: () => boolean;
  /** Select the first word through the shared editor controller. */
  selectFirstWord: () => boolean;
  /** Count painted range-selection rectangles. */
  countSelectionRects: () => number;
  /** Insert a rows×cols table at the selection (core helper). Returns success. */
  insertTable: (rows: number, cols: number) => boolean;
  /** Count `table` nodes in the live document (0 with no live view). */
  countTables: () => number;
  /** Comment-mark the first word via the shared core schema. Returns success. */
  commentFirstWord: () => boolean;
  /** Count painted `[data-comment-id]` anchors in the pages (shared painter attr). */
  countCommentAnchors: () => number;
  /** Block count of the AI-edit snapshot over the live doc (0 with no live view). */
  aiSnapshotBlockCount: () => number;
  /** Reveal the first stable snapshot block and report target/current pages. */
  navigateToFirstBlock: () => { shown: boolean; targetPage: number; currentPage: number };
  /** Plain text of the current live editor selection. */
  getSelectedText: () => string;
  /** Apply and undo one direct document-operation batch; true only when content restores. */
  applyAndUndoDocumentOperation: () => boolean;
  /** Push an anonymization term matching the first word. Returns whether one was pushed. */
  anonymizeFirstWord: () => boolean;
  /** Count painted anonymization highlight rects in the overlay. */
  countAnonymizationRects: () => number;
  /** Serialize to DOCX and return the byte length (0 on failure). */
  save: () => Promise<number>;
  /** Whether the live editor has edits not yet serialized by save(). */
  hasPendingChanges: () => boolean;
  /**
   * Insert text through `getEditorRef().dispatch` (the nested `PagedEditorRef`
   * handle), not the raw ProseMirror `view.dispatch` the other insert methods
   * use. Exercises the Vue-synthesized ref's `dispatch` method end-to-end.
   */
  insertTextViaPagedEditorRef: (text: string) => boolean;
  /**
   * Page number (1-indexed) containing the current selection anchor, resolved
   * through `getEditorRef().getPageNumberForPmPos`. 0 with no live view/layout.
   */
  getPageNumberForSelection: () => number;
};

export function buildParityBridge(getRef: () => DocxEditorRef | null): FolioParityBridge {
  const liveView = () => getRef()?.getEditor()?.getView() ?? null;
  return {
    getTotalPages: () => getRef()?.getTotalPages() ?? 0,
    ensureView: () => getRef()?.ensureEditorView({ focus: false }),
    hasView: () => liveView() !== null,
    getDocumentText: () => getRef()?.getEditor()?.getState()?.doc.textContent ?? "",
    insertText: (text) => {
      const view = liveView();
      if (!view) {
        return false;
      }
      const { state } = view;
      view.dispatch(state.tr.insertText(text, state.selection.from, state.selection.to));
      return true;
    },
    boldFirstWord: () => {
      const view = liveView();
      if (!view) {
        return false;
      }
      const boldType = view.state.schema.marks["bold"];
      if (!boldType) {
        return false;
      }
      let range: { from: number; to: number } | null = null;
      view.state.doc.descendants((node, pos) => {
        if (range) {
          return false;
        }
        if (!node.isText || !node.text) {
          return true;
        }
        const leading = node.text.length - node.text.trimStart().length;
        const word = node.text.slice(leading).split(/\s+/)[0];
        if (!word) {
          return true;
        }
        range = { from: pos + leading, to: pos + leading + word.length };
        return false;
      });
      if (!range) {
        return false;
      }
      const { from, to } = range;
      view.dispatch(view.state.tr.addMark(from, to, boldType.create()));
      return view.state.doc.rangeHasMark(from, to, boldType);
    },
    selectFirstWord: () => {
      const editor = getRef()?.getEditor();
      const view = editor?.getView();
      if (!editor || !view) {
        return false;
      }
      let range: { from: number; to: number } | null = null;
      view.state.doc.descendants((node, pos) => {
        if (range || !node.isText || !node.text) {
          return range === null;
        }
        const leading = node.text.length - node.text.trimStart().length;
        const word = node.text.slice(leading).split(/\s+/u).at(0);
        if (!word) {
          return true;
        }
        range = { from: pos + leading, to: pos + leading + word.length };
        return false;
      });
      if (!range) {
        return false;
      }
      const { from, to } = range;
      editor.setSelection(from, to);
      return true;
    },
    countSelectionRects: () => document.querySelectorAll("[data-folio-selection-rect]").length,
    insertTable: (rows, cols) => {
      const view = liveView();
      if (!view) {
        return false;
      }
      return insertTableInView(view, rows, cols);
    },
    countTables: () => {
      const view = liveView();
      if (!view) {
        return 0;
      }
      let count = 0;
      view.state.doc.descendants((node) => {
        if (node.type.name === "table") {
          count += 1;
        }
      });
      return count;
    },
    commentFirstWord: () => {
      const view = liveView();
      if (!view) {
        return false;
      }
      const commentType = view.state.schema.marks["comment"];
      if (!commentType) {
        return false;
      }
      let range: { from: number; to: number } | null = null;
      view.state.doc.descendants((node, pos) => {
        if (range) {
          return false;
        }
        if (!node.isText || !node.text) {
          return true;
        }
        const leading = node.text.length - node.text.trimStart().length;
        const word = node.text.slice(leading).split(/\s+/)[0];
        if (!word) {
          return true;
        }
        range = { from: pos + leading, to: pos + leading + word.length };
        return false;
      });
      if (!range) {
        return false;
      }
      const { from, to } = range;
      view.dispatch(view.state.tr.addMark(from, to, commentType.create({ commentId: 424242 })));
      return view.state.doc.rangeHasMark(from, to, commentType);
    },
    countCommentAnchors: () =>
      document.querySelectorAll(".paged-editor__pages [data-comment-id]").length,
    aiSnapshotBlockCount: () => getRef()?.createAIEditSnapshot()?.blocks.length ?? 0,
    navigateToFirstBlock: () => {
      const ref = getRef();
      const snapshot = ref?.createAIEditSnapshot();
      const firstBlock = snapshot?.blocks.at(0);
      if (!ref || !snapshot || !firstBlock) {
        return { shown: false, targetPage: 0, currentPage: 0 };
      }
      const target = { type: "block", story: "main", blockId: firstBlock.id } as const;
      const targetPage = ref.getTargetPage(target, snapshot) ?? 0;
      const shown = ref.showInDocument(target, snapshot);
      return { shown, targetPage, currentPage: ref.getCurrentPage() };
    },
    getSelectedText: () => getRef()?.getSelectionText() ?? "",
    applyAndUndoDocumentOperation: () => {
      const ref = getRef();
      const firstSnapshot = ref?.createAIEditSnapshot();
      const firstBlock = firstSnapshot?.blocks.at(0);
      const before = liveView()?.state.doc.textContent;
      if (!ref || !firstSnapshot || !firstBlock || before === undefined) {
        return false;
      }
      const first = ref.applyDocumentOperations({
        snapshot: firstSnapshot,
        batch: {
          version: 1,
          mode: "direct",
          operations: [
            {
              id: "parity-undo-first",
              type: "insertAfterBlock",
              blockId: firstBlock.id,
              text: "First temporary undo paragraph.",
            },
          ],
        },
      });
      const secondSnapshot = ref.createAIEditSnapshot();
      const secondBlock = secondSnapshot?.blocks.at(0);
      if (!first.undoHandle || !secondSnapshot || !secondBlock) {
        return false;
      }
      const second = ref.applyDocumentOperations({
        snapshot: secondSnapshot,
        batch: {
          version: 1,
          mode: "direct",
          operations: [
            {
              id: "parity-undo-second",
              type: "insertAfterBlock",
              blockId: secondBlock.id,
              text: "Second temporary undo paragraph.",
            },
          ],
        },
      });
      const view = liveView();
      if (!second.undoHandle || !view || view.state.doc.textContent === before) {
        return false;
      }

      view.dispatch(view.state.tr.setMeta("folioParitySelectionOnly", true));
      const secondUndo = ref.undoDocumentOperations(second.undoHandle);
      const firstUndo = ref.undoDocumentOperations(first.undoHandle);
      return (
        secondUndo.status === "undone" &&
        firstUndo.status === "undone" &&
        liveView()?.state.doc.textContent === before
      );
    },
    anonymizeFirstWord: () => {
      const view = liveView();
      if (!view) {
        return false;
      }
      let word: string | null = null;
      view.state.doc.descendants((node) => {
        if (word) {
          return false;
        }
        if (!node.isText || !node.text) {
          return true;
        }
        const candidate = node.text.trimStart().split(/\s+/)[0];
        if (candidate) {
          word = candidate;
          return false;
        }
        return true;
      });
      if (!word) {
        return false;
      }
      const { key, payload } = setAnonymizationTermsMeta([{ canonical: word, label: "person" }]);
      view.dispatch(view.state.tr.setMeta(key, payload));
      return true;
    },
    countAnonymizationRects: () =>
      document.querySelectorAll("[data-folio-anonymization-overlay] .folio-anonymization-term")
        .length,
    save: async () => {
      const buffer = await (getRef()?.save() ?? Promise.resolve(null));
      return buffer?.byteLength ?? 0;
    },
    hasPendingChanges: () => getRef()?.hasPendingChanges() ?? false,
    insertTextViaPagedEditorRef: (text) => {
      const pagedRef = getRef()?.getEditorRef();
      const view = pagedRef?.getView();
      if (!pagedRef || !view) {
        return false;
      }
      const { state } = view;
      pagedRef.dispatch(state.tr.insertText(text, state.selection.from, state.selection.to));
      return true;
    },
    getPageNumberForSelection: () => {
      const pagedRef = getRef()?.getEditorRef();
      const view = pagedRef?.getView();
      if (!pagedRef || !view) {
        return 0;
      }
      return pagedRef.getPageNumberForPmPos(view.state.selection.from) ?? 0;
    },
  };
}
