/**
 * Page-setup composable — owns the section-properties write path
 * (margins, page size, orientation) and the ruler-driven margin /
 * indent / tab-stop edit handlers. Bumps `stateTick` after every
 * write so `currentSectionProps` re-evaluates and the rulers redraw
 * with the new shape — Vue's shallowRef wouldn't see the deep
 * mutation otherwise.
 */

import type { EditorView } from "prosemirror-view";
import type { Ref } from "vue";

import {
  applyRulerParagraphEdit,
  type RulerParagraphEdit,
} from "@stll/folio-core/prosemirror/commands/rulerParagraph";
import type {
  Document,
  SectionProperties,
} from "@stll/folio-core/types/document";

export type UsePageSetupControlsOptions = {
  editorView: Ref<EditorView | null>;
  getDocument: () => Document | null;
  readOnly: Ref<boolean>;
  stateTick: Ref<number>;
  reLayout: () => void;
  /**
   * Notified after a direct section-properties write (margins, page size,
   * orientation) so the host's `onChange` fires. These writes bypass the PM
   * transaction pipeline, so this is their only change signal; ruler indent /
   * tab-stop edits run through the view's command owner and notify via the pipeline instead.
   */
  onChange: (doc: Document) => void;
  /** Canonical sessions publish section properties through the Document journal. */
  applySectionProperties?: (
    properties: Partial<SectionProperties>,
  ) => "applied" | "refused" | "unhandled";
};

type MarginProperty =
  "marginLeft" | "marginRight" | "marginTop" | "marginBottom";

export function usePageSetupControls(opts: UsePageSetupControlsOptions) {
  function handlePageSetupApply(sp: Partial<SectionProperties>) {
    if (opts.readOnly.value) return;
    const result = opts.applySectionProperties?.(sp);
    if (result === "applied" || result === "refused") return;
    const doc = opts.getDocument();
    if (!doc?.package?.document) return;
    const existing = doc.package.document.finalSectionProperties ?? {};
    doc.package.document.finalSectionProperties = { ...existing, ...sp };
    opts.stateTick.value++;
    opts.reLayout();
    opts.onChange(doc);
  }

  function applyMarginChange(property: MarginProperty, twips: number) {
    if (opts.readOnly.value) return;
    handlePageSetupApply({ [property]: twips });
  }

  function handleLeftMarginChange(twips: number) {
    applyMarginChange("marginLeft", twips);
  }
  function handleRightMarginChange(twips: number) {
    applyMarginChange("marginRight", twips);
  }
  function handleTopMarginChange(twips: number) {
    applyMarginChange("marginTop", twips);
  }
  function handleBottomMarginChange(twips: number) {
    applyMarginChange("marginBottom", twips);
  }

  function applyRulerEdit(edit: RulerParagraphEdit) {
    const view = opts.editorView.value;
    if (!view) return;
    applyRulerParagraphEdit(view, edit);
  }
  function handleIndentLeftChange(twips: number) {
    applyRulerEdit({ type: "indentLeft", twips });
  }
  function handleIndentRightChange(twips: number) {
    applyRulerEdit({ type: "indentRight", twips });
  }
  /** Negative twips are a hanging indent, as in React. */
  function handleFirstLineIndentChange(twips: number) {
    applyRulerEdit({ type: "indentFirstLine", twips });
  }
  function handleTabStopRemove(position: number) {
    applyRulerEdit({ type: "removeTabStop", position });
  }

  return {
    handlePageSetupApply,
    handleLeftMarginChange,
    handleRightMarginChange,
    handleTopMarginChange,
    handleBottomMarginChange,
    handleIndentLeftChange,
    handleIndentRightChange,
    handleFirstLineIndentChange,
    handleTabStopRemove,
  };
}
