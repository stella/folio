/**
 * useDocumentLoader — thin React binding around DocumentLoaderManager.
 *
 * The manager owns the load-generation counter and the parse → reset → history
 * → fonts orchestration; this hook keeps the React glue: the prop-change
 * effects that drive a load, the original buffer ref, and re-binding the host
 * callbacks each render.
 */

import { useEffect, useRef, useState } from "react";

import type { DocxCompatibility } from "@stll/folio-core/docx/compatibility";
import { DocumentLoaderManager } from "@stll/folio-core/managers/DocumentLoaderManager";
import type { DocumentLoadState } from "@stll/folio-core/managers/DocumentLoaderManager";
import type { Document } from "@stll/folio-core/types/document";
import { getDocumentLoadSource } from "@stll/folio-core/utils/documentLoaderBehavior";
import type { DocxInput } from "@stll/folio-core/utils/docxInput";
import type { UseHistoryReturn } from "../../hooks/useHistory";

export type { DocumentLoadState } from "@stll/folio-core/managers/DocumentLoaderManager";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

type UseDocumentLoaderParams = {
  /** Raw DOCX input (ArrayBuffer, Uint8Array, Blob, or File). */
  documentBuffer: DocxInput | null | undefined;
  /** Pre-parsed document (alternative to documentBuffer). */
  initialDocument: Document | null | undefined;
  /** Password for Agile-encrypted .docx files (Office 2010+). */
  password?: string | undefined;
  /** History instance — used to reset/push document state. */
  history: UseHistoryReturn<Document | null>;
  /** Called when an unrecoverable parse error occurs. */
  onError: ((error: Error) => void) | undefined;
  /** Called after parsing to report whether editing can preserve fidelity. */
  onCompatibilityChange: ((compatibility: DocxCompatibility) => void) | undefined;
  /**
   * Callback invoked at the start of every load to let the host component
   * clear UI state that is coupled to the previous document (comments,
   * tracked-change sidebar, find-replace matches, etc.).
   */
  onReset: () => void;
  /** Set the document loading slice of EditorState. */
  setDocumentLoadState: (state: DocumentLoadState) => void;
};

type UseDocumentLoaderReturn = {
  /** Parse and load a raw DOCX buffer. */
  loadBuffer: (buffer: DocxInput, options?: { password?: string | undefined }) => Promise<void>;
  /** Load a pre-parsed Document. */
  loadParsedDocument: (doc: Document) => void;
  /** Reset internal + UI state for a fresh document. */
  resetForNewDocument: () => void;
  /** Ref holding the original ArrayBuffer for selective save / repack. */
  originalBufferRef: React.RefObject<ArrayBuffer | null>;
};

// ---------------------------------------------------------------------------
// Hook
// ---------------------------------------------------------------------------

export const useDocumentLoader = ({
  documentBuffer,
  initialDocument,
  password,
  history,
  onError,
  onCompatibilityChange,
  onReset,
  setDocumentLoadState,
}: UseDocumentLoaderParams): UseDocumentLoaderReturn => {
  /** Original DOCX buffer kept for selective save / full repack. */
  const originalBufferRef = useRef<ArrayBuffer | null>(null);

  // The manager instance is stable; its bound methods are created once so the
  // returned references stay referentially stable across renders.
  const [{ manager, api }] = useState(() => {
    const instance = new DocumentLoaderManager({
      history,
      onError,
      onCompatibilityChange,
      onReset,
      setDocumentLoadState,
    });
    return {
      manager: instance,
      api: {
        loadBuffer: instance.loadBuffer.bind(instance),
        loadParsedDocument: instance.loadParsedDocument.bind(instance),
        resetForNewDocument: instance.resetForNewDocument.bind(instance),
      },
    };
  });

  // Re-bind host callbacks so the manager always sees the latest closures.
  manager.setCallbacks({
    history,
    onError,
    onCompatibilityChange,
    onReset,
    setDocumentLoadState,
  });

  // React to document/documentBuffer prop changes.
  useEffect(() => {
    const source = getDocumentLoadSource({ documentBuffer, initialDocument });
    if (source.type === "none") {
      return;
    }

    if (source.type === "parsed-document") {
      api.loadParsedDocument(source.document);
      return;
    }

    void api.loadBuffer(source.buffer, { password });
  }, [documentBuffer, initialDocument, password, api]);

  // Keep decrypted ZIP bytes for save/export (falls back to the raw prop buffer).
  useEffect(() => {
    if (history.state?.originalBuffer) {
      originalBufferRef.current = history.state.originalBuffer;
      return;
    }
    if (documentBuffer instanceof ArrayBuffer) {
      originalBufferRef.current = documentBuffer;
    }
  }, [history.state, documentBuffer]);

  return {
    loadBuffer: api.loadBuffer,
    loadParsedDocument: api.loadParsedDocument,
    resetForNewDocument: api.resetForNewDocument,
    originalBufferRef,
  };
};
