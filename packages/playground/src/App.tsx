import { useCallback, useEffect, useRef, useState } from "react";
import type { ChangeEvent } from "react";

import { DocxEditor, createEmptyDocument } from "@stll/folio-react";
import type { Document as FolioDocument, DocxEditorRef, EditorMode } from "@stll/folio-react";

const ZOOM_MIN = 0.25;
const ZOOM_MAX = 2;
const ZOOM_STEP = 0.1;
const ZOOM_INITIAL = 1;

function clampZoom(zoom: number): number {
  return Math.min(Math.max(zoom, ZOOM_MIN), ZOOM_MAX);
}

declare global {
  // Test hook: visual + interaction specs read live editor state through this.
  var __folioPlayground:
    | {
        getEditorRef: () => DocxEditorRef | null;
      }
    | undefined;
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
  const editorRef = useRef<DocxEditorRef>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [currentDocument, setCurrentDocument] = useState<FolioDocument | null>(null);
  const [documentBuffer, setDocumentBuffer] = useState<ArrayBuffer | null>(null);
  const [fileName, setFileName] = useState("Untitled.docx");
  const [status, setStatus] = useState("");
  const [editorMode, setEditorMode] = useState<EditorMode>("editing");
  const [zoom, setZoom] = useState(ZOOM_INITIAL);

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

  const applyZoom = useCallback((next: number) => {
    const clamped = clampZoom(next);
    setZoom(clamped);
    editorRef.current?.setZoom(clamped);
  }, []);
  const handleZoomIn = useCallback(() => applyZoom(zoom + ZOOM_STEP), [applyZoom, zoom]);
  const handleZoomOut = useCallback(() => applyZoom(zoom - ZOOM_STEP), [applyZoom, zoom]);
  const handleZoomReset = useCallback(() => applyZoom(ZOOM_INITIAL), [applyZoom]);

  const toggleDarkMode = useCallback(() => {
    document.documentElement.classList.toggle("dark");
  }, []);

  const trackChangesOn = editorMode === "suggesting";

  useEffect(() => {
    globalThis.__folioPlayground = {
      getEditorRef: () => editorRef.current,
    };
    return () => {
      globalThis.__folioPlayground = undefined;
    };
  }, []);

  return (
    <div className="pg-shell">
      <main className="pg-editor-area">
        <DocxEditor
          ref={editorRef}
          document={documentBuffer ? null : currentDocument}
          documentBuffer={documentBuffer}
          author="Folio User"
          onError={handleError}
          showToolbar={true}
          initialZoom={ZOOM_INITIAL}
          mode={editorMode}
          onModeChange={setEditorMode}
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
          onClick={handleZoomOut}
          disabled={zoom <= ZOOM_MIN}
          aria-label="Zoom out"
          title="Zoom out"
        >
          −
        </button>
        <button
          type="button"
          className="pg-button pg-zoom"
          onClick={handleZoomReset}
          title="Reset zoom"
        >
          {Math.round(zoom * 100)}%
        </button>
        <button
          type="button"
          className="pg-button"
          onClick={handleZoomIn}
          disabled={zoom >= ZOOM_MAX}
          aria-label="Zoom in"
          title="Zoom in"
        >
          +
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

        {status && <span className="pg-status">{status}</span>}
        <span className="pg-filename">{fileName}</span>
      </div>
    </div>
  );
}
