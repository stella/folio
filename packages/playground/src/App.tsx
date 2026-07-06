import { useCallback, useEffect, useRef, useState } from "react";
import type { ChangeEvent } from "react";
import { IntlProvider } from "use-intl";

import {
  DocxEditor,
  createEmptyDocument,
  insertImageFromFile,
  insertPageBreakInView,
  insertTableInView,
  insertTableOfContentsInView,
} from "@stll/folio-react";
import type { Document as FolioDocument, DocxEditorRef, EditorMode } from "@stll/folio-react";
import { FOLIO_LOCALES, getFolioMessages } from "@stll/folio-react/messages";

import { CollaborationApp } from "./CollaborationApp";

const ZOOM_INITIAL = 1;
const DEFAULT_LOCALE = "en";
// Only Arabic in the bundled set needs RTL; flip the shell so the editor chrome
// (built on logical CSS properties) mirrors.
const RTL_LOCALES = new Set<string>(["ar", "he"]);

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
  /** Block count of the AI-edit snapshot over the live doc (0 with no live view). */
  aiSnapshotBlockCount: () => number;
  /** Serialize to DOCX and return the byte length (0 on failure). */
  save: () => Promise<number>;
};

function buildParityBridge(getRef: () => DocxEditorRef | null): FolioParityBridge {
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
    aiSnapshotBlockCount: () => getRef()?.createAIEditSnapshot()?.blocks.length ?? 0,
    save: async () => {
      const buffer = await (getRef()?.save() ?? Promise.resolve(null));
      return buffer?.byteLength ?? 0;
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
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [currentDocument, setCurrentDocument] = useState<FolioDocument | null>(null);
  const [documentBuffer, setDocumentBuffer] = useState<ArrayBuffer | null>(null);
  const [fileName, setFileName] = useState("Untitled.docx");
  const [status, setStatus] = useState("");
  const [editorMode, setEditorMode] = useState<EditorMode>("editing");
  const [locale, setLocale] = useState<string>(DEFAULT_LOCALE);

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
    setCurrentDocument(createEmptyDocument());
    setFileName("Untitled.docx");
  }, []);

  const handleNewDocument = useCallback(() => {
    setCurrentDocument(createEmptyDocument());
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
    globalThis.__folioParity = buildParityBridge(() => editorRef.current);
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
            initialZoom={ZOOM_INITIAL}
            mode={editorMode}
            onModeChange={setEditorMode}
            onInsertImage={handleInsertImage}
            onInsertTable={handleInsertTable}
            showTableInsert={true}
            onInsertPageBreak={handleInsertPageBreak}
            onInsertTOC={handleInsertTOC}
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
