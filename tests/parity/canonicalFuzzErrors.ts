/** Private error sink: retain repeated refusals even when the status text is unchanged. */
import { CanonicalSessionRefusalError } from "../../packages/core/src/controller/hiddenEditorManager";
import type { CanonicalGap } from "../../packages/core/src/types/canonicalCapabilities";
import { CanonicalSaveDiagnosticError } from "../../packages/core/src/docx/canonicalSave";
import type { SaveDiagnostic } from "../../packages/core/src/docx/saveDiagnostics";
import type { BrowserInputAction } from "../visual/browserInputTrace";
import { CANONICAL_GAP } from "../../packages/core/src/types/canonicalCapabilities";
import { CANONICAL_SAVE_DIAGNOSTIC_DISPOSITIONS } from "../../test/canonicalSaveDiagnostics";

export type CanonicalFuzzPhase =
  | { type: "load" | "finalSave" | "reload" }
  | { type: "input" | "undo" | "redo" | "save"; index: number; action: BrowserInputAction["kind"] };

type CanonicalFuzzErrorDetails =
  | { status: "refusal"; type: "CanonicalSessionRefusalError"; gap: CanonicalGap }
  | {
      status: "saveDiagnostic";
      type: "CanonicalSaveDiagnosticError";
      gap: CanonicalSaveDiagnosticError["gap"];
      diagnostic: SaveDiagnostic;
    }
  | { status: "error"; type: string };

export type CanonicalFuzzError = CanonicalFuzzErrorDetails & {
  message: string;
  phase: CanonicalFuzzPhase | null;
};

export type CanonicalFuzzObservation = {
  phase: CanonicalFuzzPhase;
  errors: CanonicalFuzzError[];
};

/** The adapter reports a successful full-save fallback through its error channel. */
export const isCanonicalSaveFallback = (error: CanonicalFuzzError) => {
  if (error.status !== "saveDiagnostic") return false;
  const disposition = CANONICAL_SAVE_DIAGNOSTIC_DISPOSITIONS[error.diagnostic.type];
  return (
    disposition.status === "permitted" &&
    error.gap === CANONICAL_GAP.save &&
    error.diagnostic.part === disposition.diagnostic.part &&
    (error.phase?.type === "save" || error.phase?.type === "finalSave")
  );
};

declare global {
  var __folioCanonicalFuzzErrors: CanonicalFuzzError[] | undefined;
  var __folioCanonicalFuzzPhase: CanonicalFuzzPhase | undefined;
}

const canonicalFuzzErrorDetails = (error: Error): CanonicalFuzzErrorDetails => {
  if (error instanceof CanonicalSessionRefusalError) {
    return { status: "refusal", type: "CanonicalSessionRefusalError", gap: error.gap };
  }
  if (error instanceof CanonicalSaveDiagnosticError) {
    return {
      status: "saveDiagnostic",
      type: "CanonicalSaveDiagnosticError",
      gap: error.gap,
      diagnostic: error.diagnostic,
    };
  }
  return { status: "error", type: error.name };
};

export const recordCanonicalFuzzError = (error: Error) => {
  globalThis.__folioCanonicalFuzzErrors?.push({
    ...canonicalFuzzErrorDetails(error),
    message: error.message,
    phase: globalThis.__folioCanonicalFuzzPhase ?? null,
  });
};
