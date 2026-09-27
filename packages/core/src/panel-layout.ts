/**
 * One layout model for the editor's side panels: the page is centred between
 * a start track (the document outline) and an end track (the comments).
 *
 * The editor measures the width it has once, and this module decides how
 * each panel is presented. Every track the page shares the row with is
 * counted here, so the page, the outline and the comments cannot overlap:
 * a panel either gets a track of its own or becomes a drawer the user opens
 * on purpose. A drawer takes no track.
 *
 * Space goes to the page first, then the outline rail, then the comments
 * column, then the full outline column:
 *
 * - comments: a column when the page, the comments and the outline rail fit;
 *   otherwise a drawer.
 * - outline: a column when it fits beside the page and the comments column;
 *   otherwise a rail of heading ticks when that fits; otherwise a drawer.
 *
 * The tier names the result: `wide` when every panel has its full column,
 * `medium` when the outline is a rail, `narrow` when a panel is a drawer.
 * The thresholds derive from the page width (at the current zoom) and the
 * panel widths below, never from fixed breakpoints: a zoomed-out page or a
 * landscape section moves them.
 */

/** Widths, in CSS pixels, every tier is derived from. */
export const PANEL_METRICS = {
  /** The outline column in the `wide` tier, border included. */
  outlineColumnWidth: 256,
  /** The outline rail in the `medium` tier, border included. */
  outlineRailWidth: 32,
  /** A comment card column. */
  commentsWidth: 280,
  /** Space between the page edge and the comment cards, and after them. */
  commentsGap: 12,
  /** Canvas kept clear on either side of the page. */
  pageMargin: 20,
  /** An outline or comments drawer, capped by the width it opens over. */
  drawerWidth: 304,
} as const;

/** Width the comments take beside the page: gap, cards, gap. */
export const COMMENTS_TRACK_WIDTH = PANEL_METRICS.commentsWidth + 2 * PANEL_METRICS.commentsGap;

export type PanelLayoutTier = "wide" | "medium" | "narrow";

/** How the outline is shown: a column, a rail of ticks, a drawer, or not at all. */
export type OutlinePresentation = "none" | "column" | "rail" | "drawer";

/** How the comments are shown: a column beside the page, a drawer, or not at all. */
export type CommentsPresentation = "hidden" | "column" | "drawer";

export type PanelLayoutInput = {
  /** Width of the row the outline track and the scroll viewport share, scrollbar excluded. */
  availableWidth: number;
  /** Width of the widest page at the current zoom. */
  pageWidth: number;
  /** Whether the document has an outline to show. */
  outline: "absent" | "available";
  /** Whether the user (or the auto-open) wants the comments shown. */
  comments: "closed" | "open";
};

export type PanelLayout = {
  tier: PanelLayoutTier;
  outline: OutlinePresentation;
  comments: CommentsPresentation;
  /** Width of the outline track at the page's start (outside the scroll viewport). */
  outlineTrackWidth: number;
  /** Width reserved after the page inside the scroll viewport for the comments. */
  commentsGutter: number;
  /** Smallest available width for each tier, for the same page and panels. */
  thresholds: { wide: number; medium: number };
};

const pageTrackWidth = (pageWidth: number) => pageWidth + 2 * PANEL_METRICS.pageMargin;

/**
 * The smallest available width for each tier with the comments wanted as
 * given: `wide` fits every column, `medium` fits the outline rail and the
 * comments column.
 */
export const panelLayoutThresholds = ({
  pageWidth,
  outline,
  comments,
}: Omit<PanelLayoutInput, "availableWidth">): PanelLayout["thresholds"] => {
  const page = pageTrackWidth(pageWidth);
  const commentsTrack = comments === "open" ? COMMENTS_TRACK_WIDTH : 0;
  const hasOutline = outline === "available";
  return {
    wide: page + commentsTrack + (hasOutline ? PANEL_METRICS.outlineColumnWidth : 0),
    medium: page + commentsTrack + (hasOutline ? PANEL_METRICS.outlineRailWidth : 0),
  };
};

const OUTLINE_TRACK_WIDTH = {
  none: 0,
  column: PANEL_METRICS.outlineColumnWidth,
  rail: PANEL_METRICS.outlineRailWidth,
  drawer: 0,
} as const satisfies Record<OutlinePresentation, number>;

const tierFor = (outline: OutlinePresentation, comments: CommentsPresentation): PanelLayoutTier => {
  if (outline === "drawer" || comments === "drawer") return "narrow";
  if (outline === "rail") return "medium";
  return "wide";
};

/** Decide each panel's presentation for the width the editor has. */
export const computePanelLayout = (input: PanelLayoutInput): PanelLayout => {
  const { availableWidth } = input;
  const thresholds = panelLayoutThresholds(input);
  const page = pageTrackWidth(input.pageWidth);
  const comments = ((): CommentsPresentation => {
    if (input.comments === "closed") return "hidden";
    return availableWidth >= thresholds.medium ? "column" : "drawer";
  })();
  const commentsGutter = comments === "column" ? COMMENTS_TRACK_WIDTH : 0;
  const outline = ((): OutlinePresentation => {
    if (input.outline === "absent") return "none";
    const beside = page + commentsGutter;
    if (availableWidth >= thresholds.wide) return "column";
    if (availableWidth >= beside + PANEL_METRICS.outlineRailWidth) return "rail";
    return "drawer";
  })();
  return {
    tier: tierFor(outline, comments),
    outline,
    comments,
    outlineTrackWidth: OUTLINE_TRACK_WIDTH[outline],
    commentsGutter,
    thresholds,
  };
};

/** Which drawer is open over the page: at most one at a time. */
export type PanelOverlay = "none" | "outline" | "comments";
