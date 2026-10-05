/** Private error sink: retain repeated refusals even when the status text is unchanged. */
import { CanonicalSaveDiagnosticError } from "../../packages/core/src/docx/canonicalSave";
import type { SaveDiagnostic } from "../../packages/core/src/docx/saveDiagnostics";
import type { BrowserInputAction } from "../visual/browserInputTrace";

export type CanonicalFuzzPhase =
  | { type: "load" | "finalSave" | "reload" }
  | { type: "input" | "undo" | "redo" | "save"; index: number; action: BrowserInputAction["kind"] };

type CanonicalFuzzErrorDetails =
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

declare global {
  var __folioCanonicalFuzzErrors: CanonicalFuzzError[] | undefined;
  var __folioCanonicalFuzzPhase: CanonicalFuzzPhase | undefined;
}

export const recordCanonicalFuzzError = (error: Error) => {
  const details =
    error instanceof CanonicalSaveDiagnosticError
      ? ({
          status: "saveDiagnostic",
          type: "CanonicalSaveDiagnosticError",
          gap: error.gap,
          diagnostic: error.diagnostic,
        } as const)
      : ({ status: "error", type: error.name } as const);
  globalThis.__folioCanonicalFuzzErrors?.push({
    ...details,
    message: error.message,
    phase: globalThis.__folioCanonicalFuzzPhase ?? null,
  });
};
