/**
 * Vue port of packages/react/src/paged-editor/useDragAutoScroll.ts.
 *
 * When the user is drag-selecting text and moves the mouse near the
 * top or bottom edge of the scroll container, this composable
 * auto-scrolls the container and continues extending the selection.
 *
 * Same numeric constants (40px edge zone, 12px/frame max speed) as
 * the React hook so the two adapters feel identical under drag.
 */
import { onScopeDispose, type Ref } from "vue";
import {
  getEditorScrollRoot,
  scrollEditorBy,
} from "@stll/folio-core/paged-layout/editorScrollRoot";
import { AUTO_SCROLL_EDGE_ZONE as EDGE_ZONE, computeAutoScrollDelta } from "../utils/autoScroll";

export type DragAutoScrollOptions = {
  pagesContainer: Ref<HTMLElement | null>;
  /** Called during auto-scroll to extend the selection at the current mouse position. */
  onScrollExtendSelection: (clientX: number, clientY: number) => void;
};

export type UseDragAutoScrollReturn = {
  updateMousePosition: (clientX: number, clientY: number) => void;
  stopAutoScroll: () => void;
};

export function useDragAutoScroll({
  pagesContainer,
  onScrollExtendSelection,
}: DragAutoScrollOptions): UseDragAutoScrollReturn {
  let rafId: number | null = null;
  let lastMouseX = 0;
  let lastMouseY = 0;
  let active = false;

  function getScrollParent(): HTMLElement | null {
    return getEditorScrollRoot(pagesContainer.value);
  }

  function stopAutoScroll() {
    active = false;
    if (rafId !== null) {
      cancelAnimationFrame(rafId);
      rafId = null;
    }
  }

  function tick() {
    if (!active) return;
    const container = getScrollParent();
    if (!container) return;

    const rect = container.getBoundingClientRect();
    const scrollDelta = computeAutoScrollDelta(rect, lastMouseY);

    if (scrollDelta !== 0) {
      scrollEditorBy(container, scrollDelta);
      onScrollExtendSelection(lastMouseX, lastMouseY);
    }
    rafId = requestAnimationFrame(tick);
  }

  function startAutoScroll() {
    if (active) return;
    active = true;
    rafId = requestAnimationFrame(tick);
  }

  /**
   * Call on every mousemove during drag.
   */
  function updateMousePosition(clientX: number, clientY: number) {
    lastMouseX = clientX;
    lastMouseY = clientY;
    if (!active) {
      const container = getScrollParent();
      if (!container) return;
      const rect = container.getBoundingClientRect();
      if (clientY < rect.top + EDGE_ZONE || clientY > rect.bottom - EDGE_ZONE) {
        startAutoScroll();
      }
    }
  }

  onScopeDispose(() => stopAutoScroll());

  return { updateMousePosition, stopAutoScroll };
}
