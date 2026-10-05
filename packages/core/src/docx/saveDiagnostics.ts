import { CANONICAL_GAP } from "../types/canonicalCapabilities";
/** Save diagnostics describe fidelity fallbacks without including document content. */
export type SaveDiagnostic =
  | { type: "sourceReplayMismatch"; part: string }
  | { type: "selectiveSaveRefused"; part: string }
  | { type: "sourceReplayUnavailable"; part: string }
  | {
      type: "canonicalResourceReplacement";
      gap: typeof CANONICAL_GAP.resourceReplacement;
      part: string;
    };
export type SaveDiagnosticOptions = {
  onDiagnostic?: ((diagnostic: SaveDiagnostic) => void) | undefined;
};
