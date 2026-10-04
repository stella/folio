/** Private error sink: retain repeated refusals even when the status text is unchanged. */
import { CanonicalSessionRefusalError } from "../../packages/core/src/controller/hiddenEditorManager";
import type { CanonicalGap } from "../../packages/core/src/types/canonicalCapabilities";

type CanonicalFuzzError =
  | { status: "refusal"; type: "CanonicalSessionRefusalError"; gap: CanonicalGap; message: string }
  | { status: "error"; type: string; message: string };

declare global {
  var __folioCanonicalFuzzErrors: CanonicalFuzzError[] | undefined;
}

export const recordCanonicalFuzzError = (error: Error) => {
  const observed =
    error instanceof CanonicalSessionRefusalError
      ? ({
          status: "refusal",
          type: "CanonicalSessionRefusalError",
          gap: error.gap,
          message: error.message,
        } as const)
      : ({ status: "error", type: error.name, message: error.message } as const);
  globalThis.__folioCanonicalFuzzErrors?.push(observed);
};
