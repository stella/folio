/** Save diagnostics describe fidelity fallbacks without including document content. */
export type SaveDiagnostic =
  | { type: "sourceReplayMismatch"; part: string }
  | { type: "selectiveSaveRefused"; part: string }
  | { type: "sourceReplayUnavailable"; part: string };
export type SaveDiagnosticOptions = {
  onDiagnostic?: ((diagnostic: SaveDiagnostic) => void) | undefined;
};
