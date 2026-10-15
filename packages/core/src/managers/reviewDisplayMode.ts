import type { DisplayMode } from "./EditorModeManager";

/** The host's reading of pending revisions; it never changes the saved document. */
export const REVIEW_DISPLAY_MODES = ["all-markup", "final", "original"] as const;
export type ReviewDisplayMode = (typeof REVIEW_DISPLAY_MODES)[number];

const REVIEW_DISPLAY_MODE_VIEWS = {
  "all-markup": "all-markup",
  final: "no-markup",
  original: "original",
} as const satisfies Record<ReviewDisplayMode, DisplayMode>;

export const reviewDisplayModeToDisplayMode = (mode: ReviewDisplayMode): DisplayMode =>
  REVIEW_DISPLAY_MODE_VIEWS[mode];
