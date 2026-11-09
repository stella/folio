import type { FlowBlock } from "./types";

export type LayoutRunReason =
  | "font-ready"
  | "hyphenation-ready"
  | "initial"
  | "layout-input"
  | "manual"
  | "transaction";

export type LayoutPhase =
  | "flow-blocks"
  | "header-footer"
  | "initial-fonts"
  | "layout-document"
  | "measure-blocks"
  | "render-pages";

export type HiddenEditorStateReason = "external-document" | "mount";

export type HiddenEditorPhase =
  "editor-state" | "editor-view" | "to-prose-doc" | "update-state";

export type DocumentLoadPhase = "docx-parse";

export type LayoutInstrumentation = {
  onDocumentLoadPhase?: (event: {
    durationMs: number;
    phase: DocumentLoadPhase;
  }) => void;
  onHiddenEditorPhase?: (event: {
    durationMs: number;
    phase: HiddenEditorPhase;
    reason: HiddenEditorStateReason;
  }) => void;
  onHiddenEditorStateCreate?: (event: {
    reason: HiddenEditorStateReason;
  }) => void;
  onHyphenationDictionaryError?: (event: {
    dictionary: string;
    message: string;
  }) => void;
  /**
   * A whole document was built: a deep copy of one, or one converted from an
   * editor state. Either costs time in proportion to the document.
   */
  onDocumentBuilt?: (event: { kind: DocumentBuildKind }) => void;
  onLayoutComplete?: (event: { reason: LayoutRunReason }) => void;
  onLayoutError?: (event: { message: string; reason: LayoutRunReason }) => void;
  onLayoutStart?: (event: { reason: LayoutRunReason }) => void;
  /**
   * A review view (the editor's `DisplayMode`) laid out part of the story as
   * authored: a revision there has no resolution.
   */
  onMarkupViewIncomplete?: (event: { view: string }) => void;
  onLayoutPhase?: (event: {
    durationMs: number;
    phase: LayoutPhase;
    reason: LayoutRunReason;
  }) => void;
  /** A block was measured: its lines were broken and its size computed. */
  onMeasureBlock?: (event: {
    blockIndex: number;
    blockKind: FlowBlock["kind"];
  }) => void;
  /** A block took a measure taken before, without measuring it again. */
  onReuseMeasure?: (event: {
    blockIndex: number;
    blockKind: FlowBlock["kind"];
  }) => void;
  onMeasureBlockError?: (event: {
    blockIndex: number;
    blockKind: FlowBlock["kind"];
    message: string;
  }) => void;
};

declare global {
  // Browser tests install this hook to count real production layout work.
  // It is undefined in normal use, so the instrumentation is a no-op.
  var __folioLayoutInstrumentation: LayoutInstrumentation | undefined;
}

/**
 * What giving a block its measure took: measuring it, reusing a measure taken
 * before, or nothing (a break or an image measures as a constant).
 */
export type MeasureOutcome = "measured" | "reused" | "constant";

/**
 * Report what giving a block its measure took. A block whose measure is a
 * constant (a break, an image) reports nothing.
 */
export function recordMeasureOutcome(
  blockIndex: number,
  block: FlowBlock,
  outcome: MeasureOutcome,
): void {
  const instrumentation = globalThis.__folioLayoutInstrumentation;
  if (!instrumentation) return;
  const event = { blockIndex, blockKind: block.kind };
  switch (outcome) {
    case "measured":
      instrumentation.onMeasureBlock?.(event);
      return;
    case "reused":
      instrumentation.onReuseMeasure?.(event);
      return;
    case "constant":
      return;
    default: {
      const unhandled: never = outcome;
      throw new TypeError(`Unhandled measure outcome ${String(unhandled)}`);
    }
  }
}

export function recordMeasureBlockError(
  blockIndex: number,
  block: FlowBlock,
  error: unknown,
): void {
  globalThis.__folioLayoutInstrumentation?.onMeasureBlockError?.({
    blockIndex,
    blockKind: block.kind,
    message: error instanceof Error ? error.message : String(error),
  });
}

export type DocumentBuildKind = "copy" | "conversion";

export function recordDocumentBuilt(kind: DocumentBuildKind): void {
  globalThis.__folioLayoutInstrumentation?.onDocumentBuilt?.({ kind });
}

export function recordLayoutComplete(reason: LayoutRunReason = "manual"): void {
  globalThis.__folioLayoutInstrumentation?.onLayoutComplete?.({ reason });
}

export function recordLayoutStart(reason: LayoutRunReason): void {
  globalThis.__folioLayoutInstrumentation?.onLayoutStart?.({ reason });
}

export function recordLayoutError(
  reason: LayoutRunReason,
  error: unknown,
): void {
  const message = error instanceof Error ? error.message : String(error);
  globalThis.__folioLayoutInstrumentation?.onLayoutError?.({
    message,
    reason,
  });
}

export function recordMarkupViewIncomplete(view: string): void {
  globalThis.__folioLayoutInstrumentation?.onMarkupViewIncomplete?.({ view });
}

/** A dictionary failed to load; layout continues without hyphenating that language. */
export function recordHyphenationDictionaryError(
  dictionary: string,
  error: Error,
): void {
  globalThis.__folioLayoutInstrumentation?.onHyphenationDictionaryError?.({
    dictionary,
    message: error.message,
  });
}

export function recordLayoutPhase(
  reason: LayoutRunReason,
  phase: LayoutPhase,
  durationMs: number,
): void {
  globalThis.__folioLayoutInstrumentation?.onLayoutPhase?.({
    durationMs,
    phase,
    reason,
  });
}

export function recordHiddenEditorStateCreate(
  reason: HiddenEditorStateReason,
): void {
  globalThis.__folioLayoutInstrumentation?.onHiddenEditorStateCreate?.({
    reason,
  });
}

export function recordHiddenEditorPhase(
  reason: HiddenEditorStateReason,
  phase: HiddenEditorPhase,
  durationMs: number,
): void {
  globalThis.__folioLayoutInstrumentation?.onHiddenEditorPhase?.({
    durationMs,
    phase,
    reason,
  });
}

export function recordDocumentLoadPhase(
  phase: DocumentLoadPhase,
  durationMs: number,
): void {
  globalThis.__folioLayoutInstrumentation?.onDocumentLoadPhase?.({
    durationMs,
    phase,
  });
}
