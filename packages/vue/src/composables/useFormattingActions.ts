/**
 * Formatting/style composable — handles paragraph-style application,
 * page break / section break insertion, symbol insertion, and
 * clear-formatting.
 *
 * PORT-BLOCKED (agent ref-API entry points): upstream additionally returns
 * `applyFormatting`, `setParagraphStyle`, and `insertBreak`, which are thin
 * wrappers over the core `@stll/folio-core/prosemirror/applyFormatting`
 * orchestrator (`applyFormatting`/`setParagraphStyle`/`insertBreak`, plus the
 * `ApplyFormattingOptions`/`InsertBreakOptions` types). That orchestrator was
 * not ported into our core (our core exposes the lower-level
 * `prosemirror/commands/*` map instead, not the higher-level mark-toggle /
 * paraId-resolving façade). Reimplementing it here would duplicate ~265 LOC of
 * core logic in the wrong layer, so those three functions are omitted rather
 * than fabricated. Restore them once `prosemirror/applyFormatting` lands in
 * core; the toolbar-facing handlers below already work against our command map.
 */

import type { Ref } from "vue";
import type { EditorView } from "prosemirror-view";
import type { Document } from "@stll/folio-core/types/document";
import {
  applyStyle,
  type ResolvedStyleAttrs,
} from "@stll/folio-core/prosemirror/commands/paragraph";
import { createStyleResolver } from "@stll/folio-core/prosemirror";
import { getCachedNumberingMap } from "@stll/folio-core/docx";
import { clearFormatting } from "@stll/folio-core/prosemirror/commands/formatting";
import { insertPageBreak } from "@stll/folio-core/prosemirror/commands/pageBreak";
import { executeEditorCommand } from "@stll/folio-core/prosemirror/executeEditorCommand";
import {
  insertSectionBreakNextPage,
  insertSectionBreakContinuous,
} from "@stll/folio-core/prosemirror/commands/sectionBreak";

export type UseFormattingActionsOptions = {
  editorView: Ref<EditorView | null>;
  /**
   * The view interactive toolbar formatting should target. While a header or
   * footer is being edited this is its EditorView, so toolbar actions land in
   * the HF and not the body (#749). Falls back to the body `editorView`.
   */
  activeView?: Ref<EditorView | null>;
  getDocument: () => Document | null;
};

export function useFormattingActions(opts: UseFormattingActionsOptions) {
  const targetView = () => opts.activeView?.value ?? opts.editorView.value;

  function handleClearFormatting() {
    const view = targetView();
    if (!view) return;
    executeEditorCommand(view, clearFormatting);
    view.focus();
  }

  function handleApplyStyle(styleId: string) {
    const view = targetView();
    if (!view) return;
    const doc = opts.getDocument();
    const styles = doc?.package?.styles;
    if (styles) {
      const resolver = createStyleResolver(styles);
      const resolved = resolver.resolveParagraphStyle(styleId);
      // Build the attrs conditionally: under exactOptionalPropertyTypes an
      // explicit `undefined` is not assignable to the optional style fields.
      const attrs: ResolvedStyleAttrs = {
        numbering: doc?.package?.numbering ? getCachedNumberingMap(doc.package.numbering) : null,
      };
      if (resolved.paragraphFormatting) attrs.paragraphFormatting = resolved.paragraphFormatting;
      if (resolved.runFormatting) attrs.runFormatting = resolved.runFormatting;
      const styleName = resolver.getStyle(styleId)?.name;
      if (styleName) attrs.styleName = styleName;
      executeEditorCommand(view, applyStyle(styleId, attrs));
    } else {
      executeEditorCommand(view, applyStyle(styleId));
    }
    view.focus();
  }

  function handleInsertPageBreak() {
    const view = opts.editorView.value;
    if (!view) return;
    executeEditorCommand(view, insertPageBreak);
    view.focus();
  }

  function handleInsertSectionBreakNextPage() {
    const view = opts.editorView.value;
    if (!view) return;
    executeEditorCommand(view, insertSectionBreakNextPage);
    view.focus();
  }

  function handleInsertSectionBreakContinuous() {
    const view = opts.editorView.value;
    if (!view) return;
    executeEditorCommand(view, insertSectionBreakContinuous);
    view.focus();
  }

  function handleInsertSymbol(symbol: string) {
    const view = targetView();
    if (!view) return;
    const { from } = view.state.selection;
    const tr = view.state.tr.insertText(symbol, from);
    view.dispatch(tr.scrollIntoView());
    view.focus();
  }

  return {
    handleClearFormatting,
    handleApplyStyle,
    handleInsertPageBreak,
    handleInsertSectionBreakNextPage,
    handleInsertSectionBreakContinuous,
    handleInsertSymbol,
  };
}
