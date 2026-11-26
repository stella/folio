import type { Node as PMNode } from "prosemirror-model";
import type { EditorState } from "prosemirror-state";
import type { Mapping, Step, StepMap } from "prosemirror-transform";

import { indexedPositionMap } from "../../internal/indexedPositionMap";
import type { RemovedSectionReference } from "../../internal/sectionEndpointResolution";
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
): {
  resolved: PMNode;
  mapping: Mapping;
  positionMap: StepMap;
  steps: Step[];
  changedRanges: {
    from: number;
    to: number;
  }[];
  preStructureMap: StepMap;
  tableRanges: {
    node: PMNode;
    inputNode: PMNode;
    sourceNodeSize: number;
    position: number;
    inputMap: StepMap;
    positionMap: StepMap;
    ranges: readonly {
      from: number;
      to: number;
    }[];
    paragraphOffsets: readonly {
      source: number;
      final: number;
    }[];
  }[];
  transfers: {
    displacedToken: string | null;
    selectedToken: string | null;
  }[];
  structuralRanges: {
    from: number;
    to: number;
  }[];
  structuralMap: StepMap;
  structural: boolean;
  failed: boolean;
  removedEndpointCount: number;
  removedReferences: RemovedSectionReference[];
  keptBoundaries: Set<PMNode>;
} =>
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
