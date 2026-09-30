/**
 * What a markup view reads: the document with every tracked change resolved
 * the way the view shows it, without changing the document.
 *
 * Word lays each review view out from the text that view shows. Original is
 * the document with every revision rejected, No Markup and Simple Markup the
 * document with every revision accepted, and All Markup the authored document
 * with its revisions inline. Laying a view out as All Markup and hiding runs
 * afterwards keeps All Markup's line breaks, justification and paragraph
 * heights, so the view is resolved here, before layout, by the same resolver
 * the Accept All and Reject All commands dispatch.
 *
 * The resolved document is a projection: the editor state keeps the authored
 * document, and `positionMap` maps its positions to the projection's so the
 * layout can address every painted character by the editor position it shows.
 */

import type { Node as PMNode } from "prosemirror-model";
import { Selection } from "prosemirror-state";
import type { EditorState } from "prosemirror-state";
import type { StepMap } from "prosemirror-transform";

import type { DisplayMode } from "../managers/EditorModeManager";
import { indexedPositionMap } from "../internal/indexedPositionMap";
import { finalRevisionParagraphRanges } from "../internal/revisionResolutionTracking";
import type { RevisionResolutionMode } from "../internal/revisionResolutionInline";
import { resolveWholeStory } from "../internal/wholeStoryRevisionResolution";
import { getDocumentNumbering } from "./plugins/documentNumbering";
import { getDocumentStyleResolver } from "./plugins/documentStyleState";

/** Resolve every revision in the state's story with the state's own styles and numbering. */
export const resolveStateStory = (state: EditorState, mode: RevisionResolutionMode) =>
  resolveWholeStory({
    doc: state.doc,
    mode,
    styleResolver: getDocumentStyleResolver(state),
    numbering: getDocumentNumbering(state),
  });

/** How a view marks the paragraphs whose authored text it does not show verbatim. */
type ChangeIndicators = "none" | "change-bars";

type MarkupViewResolution =
  | { type: "authored" }
  | { type: "resolved"; mode: RevisionResolutionMode; indicators: ChangeIndicators };

/**
 * Every view's reading of the document. Total over `DisplayMode`, so a new view
 * cannot ship without a decision about what it lays out.
 */
const MARKUP_VIEW_RESOLUTIONS = {
  "all-markup": { type: "authored" },
  "simple-markup": { type: "resolved", mode: "accept", indicators: "change-bars" },
  "no-markup": { type: "resolved", mode: "accept", indicators: "none" },
  original: { type: "resolved", mode: "reject", indicators: "none" },
} as const satisfies Record<DisplayMode, MarkupViewResolution>;

type PositionRange = { from: number; to: number };

export type MarkupViewProjection =
  /** The view lays the authored document out as it stands. */
  | { type: "authored" }
  | {
      type: "resolved";
      /** The document as the view reads it. */
      doc: PMNode;
      /** Authored (editor) positions to `doc` positions. */
      positionMap: StepMap;
      /** `positionMap`, indexed for lookups. */
      toView: ReturnType<typeof indexedPositionMap>;
      /** `positionMap` inverted and indexed: `doc` positions to editor positions. */
      toEditor: ReturnType<typeof indexedPositionMap>;
      /**
       * Ranges of `doc` whose paragraphs carry a change indicator (Simple
       * Markup's change bar); empty for a view without indicators.
       */
      changeBarRanges: readonly PositionRange[];
      /**
       * Whether the resolver left part of the story as authored (a table
       * revision it cannot resolve). The projection still shows every revision
       * it could resolve; the caller reports the rest.
       */
      completeness: "complete" | "partial";
    };

const AUTHORED: MarkupViewProjection = { type: "authored" };

const projectionCache = new WeakMap<EditorState, Map<DisplayMode, MarkupViewProjection>>();

/**
 * The document a markup view lays out, for the state's current document.
 * Cached per editor state: the resolver also reads its styles and numbering
 * plugins, which may change while the document node stays the same.
 */
/** How a view resolves the revisions it reads, or null when it shows them inline. */
export const markupViewResolutionMode = (view: DisplayMode): RevisionResolutionMode | null => {
  const resolution = MARKUP_VIEW_RESOLUTIONS[view];
  return resolution.type === "resolved" ? resolution.mode : null;
};

export const projectMarkupView = (state: EditorState, view: DisplayMode): MarkupViewProjection => {
  const resolution = MARKUP_VIEW_RESOLUTIONS[view];
  if (resolution.type === "authored") {
    return AUTHORED;
  }
  let byView = projectionCache.get(state);
  const cached = byView?.get(view);
  if (cached) {
    return cached;
  }
  const result = resolveStateStory(state, resolution.mode);
  const projection: MarkupViewProjection = result.resolved.eq(state.doc)
    ? AUTHORED
    : {
        type: "resolved",
        doc: result.resolved,
        positionMap: result.positionMap,
        toView: indexedPositionMap(result.positionMap),
        toEditor: indexedPositionMap(result.positionMap.invert()),
        changeBarRanges:
          resolution.indicators === "change-bars" ? finalRevisionParagraphRanges(result) : [],
        completeness: result.failed ? "partial" : "complete",
      };
  if (!byView) {
    byView = new Map();
    projectionCache.set(state, byView);
  }
  byView.set(view, projection);
  return projection;
};

/**
 * Where a caret at editor `position` is painted in a view. A position the view
 * shows is its own; one inside text the view does not show (a deletion in No
 * Markup, an insertion in Original, typed there with tracking on) moves to the
 * nearest place the view shows before it, the way the caret sits at the edge
 * of text it cannot enter.
 */
export const visibleCaretPosition = (
  projection: MarkupViewProjection,
  position: number,
): number => {
  if (projection.type === "authored") {
    return position;
  }
  const { pos: viewPosition, deletedAcross } = projection.toView.mapResult(position, -1);
  if (!deletedAcross) {
    return position;
  }
  const bounded = Math.max(0, Math.min(viewPosition, projection.doc.content.size));
  const near = Selection.near(projection.doc.resolve(bounded), -1);
  return projection.toEditor(near.head, -1);
};
