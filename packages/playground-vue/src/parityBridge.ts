import {
  appendAutocompleteToken,
  clearAutocompleteSuggestion,
  finishAutocompleteSuggestion,
  getDocumentWatermark,
  insertTableInView,
  setAnonymizationTermsMeta,
  startAutocompleteSuggestion,
} from "@stll/folio-vue";
import type { DocxEditorRef } from "@stll/folio-vue";

import type { FolioParityBridge } from "../../../scripts/parity/bridge-contract";

export type { FolioParityBridge } from "../../../scripts/parity/bridge-contract";

export function buildParityBridge(
  getRef: () => DocxEditorRef | null,
  getClipboardCallbackCount: (kind: "copy" | "cut" | "paste") => number,
): FolioParityBridge {
  const autocompleteRequestId = "parity-autocomplete";
  const liveView = () => getRef()?.getEditor()?.getView() ?? null;
  return {
    runGeneratedFlow: async (source, batches) => {
      const ref = getRef();
      if (!ref) throw new Error("Generated flow requires an editor ref");
      await ref.loadDocumentBuffer(new Uint8Array(source));
      ref.ensureEditorView({ focus: false });
      const results = batches.map((batch) => {
        const snapshot = ref.createAIEditSnapshot();
        if (!snapshot) throw new Error("Generated flow requires a snapshot");
        return ref.applyDocumentOperations({ snapshot, batch });
      });
      const saved = await ref.save();
      if (!saved) throw new Error("Generated flow did not save");
      return { bytes: [...new Uint8Array(saved)], results };
    },
    getTotalPages: () => getRef()?.getTotalPages() ?? 0,
    ensureView: () => getRef()?.ensureEditorView({ focus: false }),
    hasView: () => liveView() !== null,
    getDocumentText: () => getRef()?.getEditor()?.getState()?.doc.textContent ?? "",
    getTextWatermark: () => {
      const document = getRef()?.getDocument();
      if (!document) {
        return null;
      }
      const watermark = getDocumentWatermark(document);
      return watermark?.kind === "text" ? watermark.text : null;
    },
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
    setupContentControls: () => {
      const view = liveView();
      if (!view) {
        return false;
      }
      const dropdown = view.state.schema.node(
        "blockSdt",
        {
          sdtType: "dropdown",
          tag: "state",
          listItems: JSON.stringify([
            { displayText: "California", value: "ca" },
            { displayText: "New York", value: "ny" },
          ]),
        },
        [view.state.schema.node("paragraph", {}, [view.state.schema.text("California")])],
      );
      const date = view.state.schema.node("blockSdt", { sdtType: "date", tag: "effective" }, [
        view.state.schema.node("paragraph", {}, [view.state.schema.text("2026-01-15")]),
      ]);
      view.dispatch(view.state.tr.replaceWith(0, view.state.doc.content.size, [dropdown, date]));
      return true;
    },
    dispatchClipboardEvent: (kind) => {
      const view = liveView();
      if (!view) {
        return 0;
      }
      view.dom.dispatchEvent(new ClipboardEvent(kind, { bubbles: true, cancelable: true }));
      return getClipboardCallbackCount(kind);
    },
    getCurrentTableProperties: () => {
      const view = liveView();
      if (!view) {
        return null;
      }
      const { $from } = view.state.selection;
      for (let depth = $from.depth; depth >= 0; depth--) {
        const node = $from.node(depth);
        if (node.type.name !== "table") {
          continue;
        }
        const width = node.attrs["width"];
        const widthType = node.attrs["widthType"];
        const justification = node.attrs["justification"];
        return {
          width: typeof width === "number" ? width : null,
          widthType: typeof widthType === "string" ? widthType : null,
          justification: typeof justification === "string" ? justification : null,
        };
      }
      return null;
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
    aiSnapshotBlockCount: () => getRef()?.createAIEditSnapshot()?.blocks.length ?? 0,
    pendingSuggestionPersistence: () => {
      const ref = getRef();
      const snapshot = ref?.createAIEditSnapshot();
      const block = snapshot?.blocks.at(0);
      if (!ref || !snapshot || !block) {
        return { exported: 0, restaged: 0, stale: 0, active: 0, version: null };
      }
      const result = ref.applyAIEditOperations({
        snapshot,
        mode: "suggested",
        operations: [
          {
            id: "parity-pending",
            type: "insertAfterBlock",
            blockId: block.id,
            text: "Pending proposal.",
          },
        ],
      });
      const id = result.applied.at(0)?.suggestionId;
      const parsed: unknown = JSON.parse(JSON.stringify(ref.exportPendingSuggestions()));
      const records = Array.isArray(parsed) ? parsed : [];
      if (id) ref.rejectSuggestion(id);
      const loaded = ref.loadPendingSuggestions(records);
      const version = records.at(0);
      return {
        exported: records.length,
        restaged: loaded.filter((entry) => entry.status === "restaged").length,
        stale: loaded.filter((entry) => entry.status === "stale").length,
        active: ref.getSuggestions().length,
        version:
          typeof version === "object" &&
          version !== null &&
          "version" in version &&
          typeof version.version === "number"
            ? version.version
            : null,
      };
    },
    readBlockGeometry: () => {
      const ref = getRef();
      const snapshot = ref?.createAIEditSnapshot();
      if (!ref || !snapshot) {
        return { rects: [], missingIsNull: true, hasScrollRoot: false };
      }
      const blockIds = snapshot.blocks.map(({ id }) => id);
      const rects = ref.getBlockRects(blockIds);
      return {
        rects: blockIds.flatMap((snapshotBlockId) => {
          const rect = rects.get(snapshotBlockId);
          return rect
            ? [
                {
                  snapshotBlockId,
                  blockId: rect.blockId,
                  page: rect.page,
                  top: rect.top,
                  height: rect.height,
                },
              ]
            : [];
        }),
        missingIsNull: ref.getBlockRect("missing-block-id") === null,
        hasScrollRoot: ref.getScrollRoot() !== null,
      };
    },
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
    startAutocomplete: (text) => {
      const view = liveView();
      if (!view) {
        return false;
      }
      view.dispatch(
        startAutocompleteSuggestion(
          view.state.tr,
          view.state.selection.head,
          autocompleteRequestId,
        ),
      );
      view.dispatch(appendAutocompleteToken(view.state.tr, autocompleteRequestId, text));
      return true;
    },
    finishAutocomplete: () => {
      const view = liveView();
      if (!view) {
        return false;
      }
      view.dispatch(finishAutocompleteSuggestion(view.state.tr, autocompleteRequestId));
      return true;
    },
    clearAutocomplete: () => {
      const view = liveView();
      if (!view) {
        return false;
      }
      view.dispatch(clearAutocompleteSuggestion(view.state.tr));
      return true;
    },
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
    typeThenReloadDocument: (marker) => {
      const ref = getRef();
      const view = liveView();
      const before = ref?.getDocument();
      if (!ref || !view || !before) {
        return false;
      }
      view.dispatch(view.state.tr.insertText(marker, view.state.selection.from));
      ref.loadDocument(before);
      return true;
    },
  };
}
