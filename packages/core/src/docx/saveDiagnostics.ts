import type { SelectiveSaveDecline } from "../types/selectiveSaveDecline";

/** Save diagnostics describe fidelity fallbacks without including document content. */
export type SaveDiagnostic =
  | { type: "sourceReplayMismatch"; part: string }
  | { type: "selectiveSaveRefused"; part: string }
  | { type: "sourceReplayUnavailable"; part: string };
export type SaveDiagnosticOptions = {
  /** The condition that declined an attempted selective write. */
  onSelectiveDecline?: ((decline: SelectiveSaveDecline) => void) | undefined;
  onDiagnostic?: ((diagnostic: SaveDiagnostic) => void) | undefined;
};
