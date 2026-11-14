/** Save diagnostics describe fidelity fallbacks without including document content. */
export type SaveDiagnostic =
  | { type: "sourceReplayMismatch"; part: string }
  | { type: "selectiveSaveRefused"; part: string }
  | { type: "sourceReplayUnavailable"; part: string };
/** Which condition made the selective writer decline, so the save repacked. */
export const SELECTIVE_SAVE_DECLINE_REASONS = [
  "untrackedChanges",
  "tooLarge",
  "unmaterializedBodyResource",
  "unmaterializedStoryResource",
  "unmaterializedHeaderFooter",
  "pictureWatermark",
  "orphanCommentRanges",
  "unsynthesizedReplyRanges",
  "invalidModel",
  "unmaterializedHyperlink",
  "unsafePackagePath",
  "missingDocumentPart",
  "addedNumbering",
  "unroutedParagraphs",
  "paragraphPatch",
  "structuralRangeElement",
  "structuralSplice",
  "commentsExtended",
  "numberingPart",
  "styleMaterialization",
  "strictConformance",
  "packageError",
  "noSourcePackage",
] as const;
export type SelectiveSaveDeclineReason =
  (typeof SELECTIVE_SAVE_DECLINE_REASONS)[number];
export type SelectiveSaveDecline = {
  part: string;
  reason: SelectiveSaveDeclineReason;
};
export type SaveDiagnosticOptions = {
  onDiagnostic?: ((diagnostic: SaveDiagnostic) => void) | undefined;
  /** Why the selective writer declined, when it did. */
  onSelectiveDecline?: ((decline: SelectiveSaveDecline) => void) | undefined;
};
