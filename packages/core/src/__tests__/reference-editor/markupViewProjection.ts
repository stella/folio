import type { EditorState } from "prosemirror-state";

import { indexedPositionMap } from "../../internal/indexedPositionMap";
import type { DisplayMode } from "../../managers/EditorModeManager";
import {
  type MarkupViewProjection,
  markupViewResolutionMode,
} from "../../prosemirror/markupViewProjection";
import { getDocumentNumbering } from "../../prosemirror/plugins/documentNumbering";
import { getDocumentStyleResolver } from "../../prosemirror/plugins/documentStyleState";
import type { RevisionResolutionMode } from "./revisionResolutionInline";
import { finalRevisionParagraphRanges } from "./revisionResolutionTracking";
import { resolveWholeStory } from "./wholeStoryRevisionResolution";

/** Resolve every revision in the state's story with the state's own styles and numbering. */
export const resolveStateStory = (
  state: EditorState,
  mode: RevisionResolutionMode,
) =>
  resolveWholeStory({
    doc: state.doc,
    mode,
    styleResolver: getDocumentStyleResolver(state),
    numbering: getDocumentNumbering(state),
  });

const AUTHORED: MarkupViewProjection = { type: "authored" };

const projectionCache = new WeakMap<
  EditorState,
  Map<DisplayMode, MarkupViewProjection>
>();

export const projectReferenceMarkupView = (
  state: EditorState,
  view: DisplayMode,
): MarkupViewProjection => {
  const mode = markupViewResolutionMode(view);
  if (mode === null) {
    return AUTHORED;
  }
  let byView = projectionCache.get(state);
  const cached = byView?.get(view);
  if (cached) {
    return cached;
  }
  const result = resolveStateStory(state, mode);
  const projection: MarkupViewProjection = result.resolved.eq(state.doc)
    ? AUTHORED
    : {
        type: "resolved",
        doc: result.resolved,
        positionMap: result.positionMap,
        toView: indexedPositionMap(result.positionMap),
        toEditor: indexedPositionMap(result.positionMap.invert()),
        changeBarRanges:
          view === "simple-markup" ? finalRevisionParagraphRanges(result) : [],
        completeness: result.failed ? "partial" : "complete",
      };
  if (!byView) {
    byView = new Map();
    projectionCache.set(state, byView);
  }
  byView.set(view, projection);
  return projection;
};
