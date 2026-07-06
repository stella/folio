import { insertTableInView } from "@stll/folio-vue";
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
  /** Insert a rows×cols table at the selection (core helper). Returns success. */
  insertTable: (rows: number, cols: number) => boolean;
  /** Count `table` nodes in the live document (0 with no live view). */
  countTables: () => number;
  /** Comment-mark the first word via the shared core schema. Returns success. */
  commentFirstWord: () => boolean;
  /** Count painted `[data-comment-id]` anchors in the pages (shared painter attr). */
  countCommentAnchors: () => number;
  /** Serialize to DOCX and return the byte length (0 on failure). */
  save: () => Promise<number>;
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
    save: async () => {
      const buffer = await (getRef()?.save() ?? Promise.resolve(null));
      return buffer?.byteLength ?? 0;
    },
  };
}
