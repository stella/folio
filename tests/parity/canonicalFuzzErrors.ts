import { CanonicalSessionRefusalError } from "../../packages/core/src/controller/hiddenEditorManager";
import type { CanonicalGap } from "../../packages/core/src/types/canonicalCapabilities";

/** Private error sink: retain repeated refusals even when the status text is unchanged. */
declare global {
  var __folioCanonicalFuzzErrors:
    | { type: string; message: string; gap?: CanonicalGap }[]
    | undefined;
}

export const recordCanonicalFuzzError = (error: Error) =>
  globalThis.__folioCanonicalFuzzErrors?.push({
    type: error.name,
    message: error.message,
    ...(CanonicalSessionRefusalError.is(error) ? { gap: error.gap } : {}),
  });
