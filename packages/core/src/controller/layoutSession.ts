import type { EditorState } from "prosemirror-state";

import type { FlowBlock, Measure } from "../layout-engine/types";
import type {
  TemplatePreviewEntry,
  TemplatePreviewHiddenRange,
  TemplatePreviewValues,
} from "../prosemirror/plugins/templatePreviewValues";
import type { StyleDefinitions, Theme } from "../types/document";
import type { FontAlternates } from "../fonts/fontAlternates";

export type LayoutArtifacts = {
  blocks: FlowBlock[];
  blockWidths: number[];
  measures: Measure[];
};

export type LayoutTemplatePreview = {
  entries: readonly TemplatePreviewEntry[];
  hidden: readonly TemplatePreviewHiddenRange[];
  mode: TemplatePreviewValues["mode"];
};

/**
 * The inputs besides the ProseMirror document and block widths that a block's
 * measure depends on. An incremental pass keeps committed measures only when
 * these are the ones they were computed with. The `Document` model itself is
 * not among them: adapters hand back a new one after every edit, while a load
 * that changes how text measures brings new styles or font alternates.
 */
export type LayoutMeasureInputs = {
  styles: StyleDefinitions | null | undefined;
  theme: Theme | null | undefined;
  defaultTabStop: number | undefined;
  pageContentHeight: number;
  fontsLoaded: boolean;
  fontAlternates: FontAlternates;
};

// Controller-owned memory for the incremental layout loop: the previous run's
// artifacts and the inputs that produced them, so a relayout can decide whether
// it can reuse work or must recompute.
export type LayoutSession = {
  artifacts: LayoutArtifacts | null;
  lastEditorState: EditorState | null;
  lastPmDoc: EditorState["doc"] | null;
  lastMeasureInputs: LayoutMeasureInputs | null;
  usedLoadedFonts: boolean;
  lastTemplatePreview: LayoutTemplatePreview;
};

export const createLayoutSession = (): LayoutSession => ({
  artifacts: null,
  lastEditorState: null,
  lastPmDoc: null,
  lastMeasureInputs: null,
  usedLoadedFonts: false,
  lastTemplatePreview: { entries: [], hidden: [], mode: "plain" },
});
