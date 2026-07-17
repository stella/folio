import { useCallback, useEffect, useRef, useState } from "react";
import type { ChangeEvent } from "react";
import { IntlProvider } from "use-intl";

import {
  DocxEditor,
  appendAutocompleteToken,
  clearAutocompleteSuggestion,
  createEmptyDocument,
  createStellaStyleDocumentPreset,
  finishAutocompleteSuggestion,
  getDocumentWatermark,
  insertImageFromFile,
  insertPageBreakInView,
  insertTableInView,
  insertTableOfContentsInView,
  setAnonymizationTermsMeta,
  startAutocompleteSuggestion,
} from "@stll/folio-react";
import type { Document as FolioDocument, DocxEditorRef, EditorMode } from "@stll/folio-react";
import { FOLIO_LOCALES, getFolioMessages } from "@stll/folio-react/messages";

import { CollaborationApp } from "./CollaborationApp";

const ZOOM_INITIAL = 1;
const DEFAULT_LOCALE = "en";
// Only Arabic in the bundled set needs RTL; flip the shell so the editor chrome
// (built on logical CSS properties) mirrors.
const RTL_LOCALES = new Set<string>(["ar", "he"]);

const createStellaStyleDocument = (): FolioDocument =>
  createEmptyDocument({ preset: createStellaStyleDocumentPreset() });

const languageLabel = (locale: string): string => {
  const name = new Intl.DisplayNames([locale], { type: "language" }).of(
    new Intl.Locale(locale).language,
  );
  return name ? `${name} (${locale})` : locale;
};

const isCollaborationDemo = (): boolean =>
  new URLSearchParams(window.location.search).has("collaboration");

declare global {
  // Test hook: visual + interaction specs read live editor state through this.
  var __folioPlayground:
    | {
        getEditorRef: () => DocxEditorRef | null;
      }
    | undefined;
  // Cross-adapter E2E bridge: identical surface in the React and Vue
  // playgrounds so the `tests/parity` specs drive both editors through one API.
  var __folioParity: FolioParityBridge | undefined;
}

/**
 * Cross-adapter parity bridge. Every method is built on the `DocxEditorRef`
 * members classified `paired` in `scripts/parity/parity.contract.json`, so the
 * exact same implementation works against the React and Vue adapters. The Vue
 * playground exposes a byte-for-byte equivalent object.
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
  /** Text watermark content, or null when none/a non-text watermark is active. */
  getTextWatermark: () => string | null;
  /** Insert text at the current selection. Returns false with no live view. */
  insertText: (text: string) => boolean;
  /** Bold the first word of the document. Returns whether the mark applied. */
  boldFirstWord: () => boolean;
  /** Select the first word through the shared editor controller. */
  selectFirstWord: () => boolean;
  /** Count painted range-selection rectangles. */
  countSelectionRects: () => number;
  /** Replace the live document with dropdown and date content controls. */
  setupContentControls: () => boolean;
  /** Dispatch a clipboard DOM event and return the matching host callback count. */
  dispatchClipboardEvent: (kind: "copy" | "cut" | "paste") => number;
  /** Table properties at the live selection, or null outside a table. */
  getCurrentTableProperties: () => {
    width: number | null;
    widthType: string | null;
    justification: string | null;
  } | null;
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
  /** Start a streaming autocomplete suggestion at the current selection. */
  startAutocomplete: (text: string) => boolean;
  /** Mark the active autocomplete suggestion as complete. */
  finishAutocomplete: () => boolean;
  /** Dismiss the active autocomplete suggestion. */
  clearAutocomplete: () => boolean;
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

function buildParityBridge(
  getRef: () => DocxEditorRef | null,
  getClipboardCallbackCount: (kind: "copy" | "cut" | "paste") => number,
): FolioParityBridge {
  const autocompleteRequestId = "parity-autocomplete";
  const liveView = () => getRef()?.getEditor()?.getView() ?? null;
  return {
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
  };
}

function createLargeDocument(paragraphCount: number): FolioDocument {
  const document = createEmptyDocument();
  const paragraphs: FolioDocument["package"]["document"]["content"] = [];

  for (let i = 0; i < paragraphCount; i += 1) {
    paragraphs.push({
      type: "paragraph",
      content: [
        {
          type: "run",
          content: [
            {
              type: "text",
              text: `Performance paragraph ${i + 1}: This legal drafting fixture provides enough body text to exercise paged layout measurement.`,
            },
          ],
          formatting: {
            fontSize: 22,
            fontFamily: {
              ascii: "Arial",
              hAnsi: "Arial",
            },
          },
        },
      ],
      formatting: {
        lineSpacing: 276,
      },
    });
  }

  document.package.document.content = paragraphs;
  return document;
}

export function App() {
  if (isCollaborationDemo()) {
    return <CollaborationApp />;
  }

  const editorRef = useRef<DocxEditorRef>(null);
  const clipboardCallbackCountsRef = useRef({ copy: 0, cut: 0, paste: 0 });
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [currentDocument, setCurrentDocument] = useState<FolioDocument | null>(null);
  const [documentBuffer, setDocumentBuffer] = useState<ArrayBuffer | null>(null);
  const [fileName, setFileName] = useState("Untitled.docx");
  const [status, setStatus] = useState("");
  const [editorMode, setEditorMode] = useState<EditorMode>("editing");
  const [locale, setLocale] = useState<string>(DEFAULT_LOCALE);
  const query = new URLSearchParams(window.location.search);
  const showMarginGuides = query.has("marginGuides");
  const marginGuideColor = query.get("marginGuideColor") ?? undefined;

  // Load fixture from ?file= query param (visual + interaction tests) or
  // generate a body from ?paragraphs= (performance tests).
  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const fixtureFile = params.get("file");
    const paragraphCount = Number(params.get("paragraphs"));
    if (fixtureFile) {
      void (async () => {
        try {
          setStatus("Loading fixture...");
          const response = await fetch(`/fixtures/${fixtureFile}`);
          if (!response.ok) {
            setStatus(`Fixture not found: ${fixtureFile}`);
            return;
          }
          const buffer = await response.arrayBuffer();
          setCurrentDocument(null);
          setDocumentBuffer(buffer);
          setFileName(fixtureFile);
          setStatus("");
        } catch {
          setStatus("Error loading fixture");
        }
      })();
      return;
    }
    if (Number.isInteger(paragraphCount) && paragraphCount > 0) {
      setCurrentDocument(createLargeDocument(paragraphCount));
      setFileName(`Generated ${paragraphCount} paragraphs.docx`);
      return;
    }
    setCurrentDocument(createStellaStyleDocument());
    setFileName("Untitled.docx");
  }, []);

  const handleNewDocument = useCallback(() => {
    setCurrentDocument(createStellaStyleDocument());
    setDocumentBuffer(null);
    setFileName("Untitled.docx");
    setStatus("");
  }, []);

  const handleFileSelect = useCallback(async (event: ChangeEvent<HTMLInputElement>) => {
    const input = event.currentTarget;
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    try {
      setStatus("Loading...");
      const buffer = await file.arrayBuffer();
      setCurrentDocument(null);
      setDocumentBuffer(buffer);
      setFileName(file.name);
      setStatus(`Loaded ${file.name}`);
    } catch {
      setStatus("Error loading file");
    } finally {
      input.value = "";
    }
  }, []);

  const handleOpenDocument = useCallback(() => {
    const input = fileInputRef.current;
    if (!input) {
      setStatus("File picker unavailable");
      return;
    }
    input.value = "";
    input.click();
  }, []);

  const handleSave = useCallback(async () => {
    if (!editorRef.current) {
      return;
    }
    try {
      setStatus("Saving...");
      const buffer = await editorRef.current.save();
      if (buffer) {
        const blob = new Blob([buffer], {
          type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = fileName || "document.docx";
        document.body.append(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);
        setStatus("Saved!");
        setTimeout(() => setStatus(""), 2000);
      }
    } catch {
      setStatus("Save failed");
    }
  }, [fileName]);

  const handleError = useCallback((error: Error) => {
    setStatus(`Error: ${error.message}`);
  }, []);

  const handleInsertImage = useCallback(() => {
    const view = editorRef.current?.getEditorRef()?.getView();
    if (!view) {
      return;
    }
    const input = document.createElement("input");
    input.type = "file";
    input.accept = "image/*";
    input.addEventListener("change", () => {
      const file = input.files?.[0];
      if (file) {
        insertImageFromFile(view, file, () => view.focus()).catch(handleError);
      }
    });
    input.click();
  }, [handleError]);

  const handleInsertTable = useCallback((rows: number, columns: number) => {
    const view = editorRef.current?.getEditorRef()?.getView();
    if (view) {
      insertTableInView(view, rows, columns);
    }
  }, []);

  const handleInsertPageBreak = useCallback(() => {
    const view = editorRef.current?.getEditorRef()?.getView();
    if (view) {
      insertPageBreakInView(view);
    }
  }, []);

  const handleInsertTOC = useCallback(() => {
    const view = editorRef.current?.getEditorRef()?.getView();
    if (view) {
      insertTableOfContentsInView(view);
    }
  }, []);

  const toggleDarkMode = useCallback(() => {
    document.documentElement.classList.toggle("dark");
  }, []);

  const trackChangesOn = editorMode === "suggesting";

  useEffect(() => {
    globalThis.__folioPlayground = {
      getEditorRef: () => editorRef.current,
    };
    globalThis.__folioParity = buildParityBridge(
      () => editorRef.current,
      (kind) => clipboardCallbackCountsRef.current[kind],
    );
    return () => {
      globalThis.__folioPlayground = undefined;
      globalThis.__folioParity = undefined;
    };
  }, []);

  return (
    <IntlProvider
      locale={locale}
      messages={getFolioMessages(locale)}
      timeZone={Intl.DateTimeFormat().resolvedOptions().timeZone}
    >
      <div className="pg-shell" dir={RTL_LOCALES.has(locale) ? "rtl" : "ltr"}>
        <main className="pg-editor-area">
          <DocxEditor
            ref={editorRef}
            document={documentBuffer ? null : currentDocument}
            documentBuffer={documentBuffer}
            author="Folio User"
            onError={handleError}
            showToolbar={true}
            showRuler={true}
            showMarginGuides={showMarginGuides}
            {...(marginGuideColor !== undefined ? { marginGuideColor } : {})}
            initialZoom={ZOOM_INITIAL}
            mode={editorMode}
            onModeChange={setEditorMode}
            onInsertImage={handleInsertImage}
            onInsertTable={handleInsertTable}
            showTableInsert={true}
            onInsertPageBreak={handleInsertPageBreak}
            onInsertTOC={handleInsertTOC}
            onCopy={() => clipboardCallbackCountsRef.current.copy++}
            onCut={() => clipboardCallbackCountsRef.current.cut++}
            onPaste={() => clipboardCallbackCountsRef.current.paste++}
          />
        </main>

        <div className="pg-controls" data-testid="playground-controls">
          <button
            type="button"
            className="pg-button"
            aria-pressed={trackChangesOn}
            onClick={() => setEditorMode(trackChangesOn ? "editing" : "suggesting")}
          >
            {trackChangesOn ? "Tracking" : "Track Changes"}
          </button>
          <button
            type="button"
            className="pg-button"
            aria-pressed={editorMode === "viewing"}
            onClick={() => setEditorMode(editorMode === "viewing" ? "editing" : "viewing")}
          >
            View Only
          </button>

          <span className="pg-sep" aria-hidden="true" />

          <button type="button" className="pg-button" onClick={handleOpenDocument}>
            Open
          </button>
          <input
            ref={fileInputRef}
            aria-label="Open .docx file"
            id="file-input"
            type="file"
            accept=".docx"
            onChange={(e) => void handleFileSelect(e)}
            className="pg-visually-hidden"
          />
          <button type="button" className="pg-button" onClick={handleNewDocument}>
            New
          </button>
          <button type="button" className="pg-button" onClick={() => void handleSave()}>
            Save
          </button>

          <span className="pg-sep" aria-hidden="true" />

          <button
            type="button"
            className="pg-button"
            onClick={toggleDarkMode}
            aria-label="Toggle theme"
            title="Toggle theme"
          >
            ◐
          </button>

          <span className="pg-sep" aria-hidden="true" />

          <select
            data-testid="language-select"
            aria-label="Editor language"
            className="pg-button"
            value={locale}
            onChange={(event) => setLocale(event.currentTarget.value)}
          >
            {FOLIO_LOCALES.map((loc) => (
              <option key={loc} value={loc}>
                {languageLabel(loc)}
              </option>
            ))}
          </select>

          {status && <span className="pg-status">{status}</span>}
          <span className="pg-filename">{fileName}</span>
        </div>
      </div>
    </IntlProvider>
  );
}
