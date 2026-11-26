import type { EditorState } from "prosemirror-state";

import { createDisplayListPagePainter } from "../display-list/editor/displayListPagePainter";
import { PAGE_RENDERER } from "../display-list/editor/pageRenderer";
import {
  type LayoutOutcome,
  type LayoutPipelineDeps,
  runLayoutPipelineWithPainter,
} from "./layoutPipelineExecution";
import type { LayoutRunOptions } from "./layoutRunOptions";

export type {
  LayoutOutcome,
  LayoutPipelineDeps,
} from "./layoutPipelineExecution";

/** Run a synchronous layout pass for consumers of the core pipeline. */
export function runLayoutPipeline<THfPMs>(
  deps: LayoutPipelineDeps<THfPMs>,
  state: EditorState,
  options?: LayoutRunOptions,
): LayoutOutcome {
  return runLayoutPipelineWithPainter(
    deps.pageRenderer === PAGE_RENDERER.displayList
      ? { ...deps, displayListPagePainter: createDisplayListPagePainter }
      : deps,
    state,
    options,
  );
}
