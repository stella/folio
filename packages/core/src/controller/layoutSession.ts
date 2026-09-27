import type { EditorState } from "prosemirror-state";

import type { FlowBlock, Measure } from "../layout-engine/types";
import type { DisplayMode } from "../managers/EditorModeManager";
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
  /**
   * The font set (`readFontSetSignature`) the measures were taken in. A face
   * that loads afterwards changes it, so fallback measures are never reused.
   */
  fontSet: string;
  fontAlternates: FontAlternates;
  /** The review view: each lays out different text, so a measure is reusable only within one. */
  markupView: DisplayMode;
};

// Controller-owned memory for the incremental layout loop: the previous run's
// artifacts and the inputs that produced them, so a relayout can decide whether
// it can reuse work or must recompute.
export type LayoutSession = {
  artifacts: LayoutArtifacts | null;
  lastEditorState: EditorState | null;
  lastPmDoc: EditorState["doc"] | null;
  lastMeasureInputs: LayoutMeasureInputs | null;
  lastTemplatePreview: LayoutTemplatePreview;
};

export const createLayoutSession = (): LayoutSession => ({
  artifacts: null,
  lastEditorState: null,
  lastPmDoc: null,
  lastMeasureInputs: null,
  lastTemplatePreview: { entries: [], hidden: [], mode: "plain" },
});
