/**
 * Outline + sidebar composable — owns the outline panel toggle (which
 * lazily collects headings from the PM doc on open), the sidebar
 * toggle (which seeds extractCommentsAndChanges so the panel opens
 * with fresh data), the outline-navigate scroll-into-view, and the
 * "click-outside-sidebar clears active item" mouse-down handler.
 */

import type { Ref } from "vue";
import type { EditorView } from "prosemirror-view";
import { TextSelection } from "prosemirror-state";
import { collectHeadings } from "@stll/folio-core/utils/headingCollector";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";

export type UseOutlineSidebarOptions = {
  editorView: Ref<EditorView | null>;
  showOutline: Ref<boolean>;
  showSidebar: Ref<boolean>;
  outlineHeadings: Ref<HeadingInfo[]>;
  activeSidebarItem: Ref<string | null>;
  extractCommentsAndChanges: () => void;
  /**
   * Scroll the VISIBLE paged viewport to a PM position. The hidden PM
   * (`left: -9999px`) is never what the user sees, so `tr.scrollIntoView()`
   * on it does nothing; the outline must drive the painter viewport instead
   * (mirrors React's `scrollToPosition`; #930).
   */
  scrollToVisiblePosition: (pmPos: number) => void;
}

export function useOutlineSidebar(opts: UseOutlineSidebarOptions) {
  function recomputeHeadings() {
    const view = opts.editorView.value;
    if (view) {
      opts.outlineHeadings.value = collectHeadings(view.state.doc);
    }
  }

  function handleToggleOutline() {
    if (!opts.showOutline.value) {
      // Opening: collect headings so the panel isn't stale from before the
      // panel existed (or from before the doc-change recompute below fired).
      recomputeHeadings();
    }
    opts.showOutline.value = !opts.showOutline.value;
  }

  /**
   * Re-collect headings after a document change, but only while the outline
   * panel is open — a closed panel re-collects lazily on next open via
   * `handleToggleOutline`. Mirrors React's doc-change-triggered heading
   * recompute (`DocxEditor.tsx`'s `handleDocumentChange`, gated on
   * `showOutlineRef.current`).
   */
  function recomputeHeadingsIfOpen() {
    if (opts.showOutline.value) {
      recomputeHeadings();
    }
  }

  function handleOutlineNavigate(pmPos: number) {
    const view = opts.editorView.value;
    if (!view) return;
    // Put the caret on the heading in the hidden PM (so typing continues from
    // there), then scroll the VISIBLE pages — not the hidden PM — to it.
    const $pos = view.state.doc.resolve(Math.min(pmPos + 1, view.state.doc.content.size));
    const sel = TextSelection.near($pos);
    view.dispatch(view.state.tr.setSelection(sel));
    opts.scrollToVisiblePosition(pmPos);
    view.focus();
  }

  function handleToggleSidebar() {
    if (!opts.showSidebar.value) {
      opts.extractCommentsAndChanges();
    }
    opts.showSidebar.value = !opts.showSidebar.value;
  }

  function handleEditorScrollMouseDown(event: MouseEvent) {
    const target = event.target;
    if (!(target instanceof HTMLElement)) return;
    // Clicks inside the editor pages, sidebar card, or comment margin
    // markers are real interactions — leave activeSidebarItem alone.
    if (
      target.closest(".paged-editor__pages") ||
      target.closest(".unified-sidebar") ||
      target.closest(".docx-comment-margin-markers")
    ) {
      return;
    }
    opts.activeSidebarItem.value = null;
  }

  return {
    handleToggleOutline,
    handleOutlineNavigate,
    handleToggleSidebar,
    handleEditorScrollMouseDown,
    recomputeHeadingsIfOpen,
  };
}
