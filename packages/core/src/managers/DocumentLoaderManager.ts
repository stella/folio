/**
 * DocumentLoaderManager
 *
 * Framework-agnostic DOCX load/state orchestration, extracted from the React
 * `useDocumentLoader` hook. Owns the monotonic load-generation counter that
 * discards stale async loads and runs the parse → reset → history → fonts
 * sequence. The pure load-source resolution lives in `documentLoaderBehavior`.
 *
 * The adapter keeps only React concerns: the prop-change effects, the original
 * buffer ref, and the host callbacks (which it re-binds via `setCallbacks`).
 */

import { inspectDocxCompatibility } from "../docx/compatibility";
import type { DocxCompatibility } from "../docx/compatibility";
import { parseDocx } from "../docx/parser";
import { recordDocumentLoadPhase } from "../layout-engine/layoutInstrumentation";
import type { Document } from "../types/document";
import { resetAuthorColors } from "../utils/authorColors";
import type { DocxInput } from "../utils/docxInput";
import { loadFontsWithMapping } from "../utils/fontLoader";

export type DocumentLoadState =
  | { status: "loading" }
  | { status: "ready" }
  | { status: "error"; message: string };

/**
 * Minimal structural view of the editor history the loader drives. Declared
 * here so core does not depend on the React `useHistory` return type.
 */
export type DocumentLoaderHistory = {
  /** The current document, or null before any load. */
  readonly state: Document | null;
  /** Reset history to a freshly loaded document. */
  reset: (document: Document) => void;
};

export type DocumentLoaderCallbacks = {
  /** Editor history to reset/seed with the loaded document. */
  history: DocumentLoaderHistory;
  /** Called when an unrecoverable parse error occurs. */
  onError: ((error: Error) => void) | undefined;
  /** Called after parsing to report whether editing can preserve fidelity. */
  onCompatibilityChange: ((compatibility: DocxCompatibility) => void) | undefined;
  /** Clears UI state coupled to the previous document at the start of a load. */
  onReset: () => void;
  /** Sets the document loading slice of editor state. */
  setDocumentLoadState: (state: DocumentLoadState) => void;
};

export class DocumentLoaderManager {
  private loadGeneration = 0;
  private callbacks: DocumentLoaderCallbacks;

  constructor(callbacks: DocumentLoaderCallbacks) {
    this.callbacks = callbacks;
  }

  /** Re-bind the host callbacks (the adapter refreshes these every render). */
  setCallbacks(callbacks: DocumentLoaderCallbacks): void {
    this.callbacks = callbacks;
  }

  /** Reset color/UI state coupled to the previous document. */
  resetForNewDocument(): void {
    resetAuthorColors();
    this.callbacks.onReset();
  }

  /** Load an already-parsed document. */
  loadParsedDocument(doc: Document): void {
    const { history, onCompatibilityChange, setDocumentLoadState } = this.callbacks;
    this.resetForNewDocument();
    history.reset(doc);
    onCompatibilityChange?.(inspectDocxCompatibility(doc));
    setDocumentLoadState({ status: "ready" });
    // Defer font loading so the first page renders immediately.
    if (doc.requiredFonts && doc.requiredFonts.length > 0) {
      loadFontsWithMapping(doc.requiredFonts).catch(() => undefined);
    }
  }

  /**
   * Parse and load a raw DOCX buffer. A monotonic generation counter discards
   * the result (and any error) when a newer load started while this one was in
   * flight.
   */
  async loadBuffer(
    buffer: DocxInput,
    options: { password?: string | undefined } = {},
  ): Promise<void> {
    const { history, onError, setDocumentLoadState } = this.callbacks;
    const generation = ++this.loadGeneration;
    const hasLoadedDocument = history.state !== null;
    if (!hasLoadedDocument) {
      setDocumentLoadState({ status: "loading" });
    }

    try {
      // Skip blocking font preload during parsing; fonts load asynchronously
      // in loadParsedDocument after the first render.
      const parseStartedAt = performance.now();
      let doc: Document;
      try {
        doc = await parseDocx(buffer, {
          detectVariables: false,
          preloadFonts: false,
          password: options.password,
        });
      } finally {
        recordDocumentLoadPhase("docx-parse", performance.now() - parseStartedAt);
      }
      if (this.loadGeneration !== generation) {
        return;
      }
      this.loadParsedDocument(doc);
    } catch (error) {
      if (this.loadGeneration !== generation) {
        return;
      }
      const message = error instanceof Error ? error.message : "Failed to parse document";
      setDocumentLoadState({ status: "error", message });
      onError?.(error instanceof Error ? error : new Error(message));
    }
  }
}
