/**
 * Names of the marks that carry run-level (character) formatting.
 *
 * Single source of truth shared by every place that must swap a run's live
 * formatting back to a previously recorded set: rejecting a `runPropertyChange`
 * (commands/comments.ts) and stripping a *suggested* `runPropertyChange` at the
 * serialization boundary (conversion/fromProseDoc.ts). Keeping one list makes
 * it structurally impossible for the two paths to drift and let a formatting
 * mark survive a revert.
 */
export const RUN_FORMATTING_MARK_NAMES = new Set([
  "bold",
  "italic",
  "underline",
  "strike",
  "textColor",
  "highlight",
  "runShading",
  "fontSize",
  "fontFamily",
  "language",
  "superscript",
  "subscript",
  "allCaps",
  "smallCaps",
  "characterSpacing",
  "emboss",
  "imprint",
  "hidden",
  "textShadow",
  "emphasisMark",
  "textOutline",
  "rtl",
  "textEffect",
  "runFormattingOverride",
  "characterStyle",
]);
