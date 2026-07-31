import { onScopeDispose, watch, type Ref, type ShallowRef } from "vue";
import type { EditorView } from "prosemirror-view";

import { BodySelectionOverlay } from "@stll/folio-core/render-dom/BodySelectionOverlay";
import { getCaretPositionFromDom } from "@stll/folio-core/layout-bridge/dom/clickToPositionDom";
import {
  resetImeCaretAnchor,
  syncImeCaretAnchor,
} from "@stll/folio-core/layout-bridge/dom/imeCaretAnchor";
import type { LayoutSelectionGate } from "@stll/folio-core/paged-layout/LayoutSelectionGate";

import type { ImageSelectionInfo } from "../components/imageSelectionTypes";
import { Z_INDEX } from "../styles/zIndex";

export type UseSelectionSyncOptions = {
  editorView: Ref<EditorView | null>;
  hiddenContainer: Ref<HTMLElement | null>;
  pagesRef: Ref<HTMLElement | null>;
  zoom: Ref<number>;
  selectedImage: ShallowRef<ImageSelectionInfo | null>;
  syncCoordinator: LayoutSelectionGate;
  imageInteracting?: Ref<boolean>;
};

export type UseSelectionSyncReturn = {
  clearOverlay: () => void;
  updateSelectionOverlay: () => void;
};

export const useSelectionSync = (opts: UseSelectionSyncOptions): UseSelectionSyncReturn => {
  const overlay = new BodySelectionOverlay();
  let animationFrame: number | null = null;

  const clearOverlay = (): void => {
    if (animationFrame !== null) {
      cancelAnimationFrame(animationFrame);
      animationFrame = null;
    }
    const pagesContainer = opts.pagesRef.value;
    if (pagesContainer) {
      overlay.clear(pagesContainer);
    }
  };

  const updateSelectionOverlay = (): void => {
    if (animationFrame !== null) {
      cancelAnimationFrame(animationFrame);
    }
    animationFrame = requestAnimationFrame(() => {
      animationFrame = null;
      if (!opts.syncCoordinator.isSafeToRender()) {
        return;
      }
      const pagesContainer = opts.pagesRef.value;
      const view = opts.editorView.value;
      if (!pagesContainer || !view) {
        return;
      }

      const result = overlay.sync({
        pagesContainer,
        state: view.state,
        zoom: opts.zoom.value,
        zIndex: Z_INDEX.selectionOverlay,
      });

      const { selection } = view.state;
      const pagesRect = pagesContainer.getBoundingClientRect();
      const caret = selection.empty
        ? getCaretPositionFromDom(pagesContainer, selection.head, pagesRect)
        : null;
      syncImeCaretAnchor({
        hiddenHost: opts.hiddenContainer.value,
        editorView: view,
        visibleCaret: caret
          ? {
              left: pagesRect.left + caret.x,
              top: pagesRect.top + caret.y,
            }
          : null,
      });

      if (result.type === "image") {
        const current = opts.selectedImage.value;
        const width = result.element.offsetWidth;
        const height = result.element.offsetHeight;
        if (
          current?.element !== result.element ||
          current.pmPos !== result.pmPos ||
          current.width !== width ||
          current.height !== height
        ) {
          opts.selectedImage.value = {
            element: result.element,
            pmPos: result.pmPos,
            width,
            height,
          };
        }
        return;
      }
      if (!opts.imageInteracting?.value) {
        opts.selectedImage.value = null;
      }
    });
  };

  const unsubscribeRender = opts.syncCoordinator.onRender(updateSelectionOverlay);
  watch([opts.zoom, opts.editorView, opts.pagesRef], updateSelectionOverlay, { flush: "post" });
  watch(
    opts.hiddenContainer,
    (hiddenContainer, _previous, onCleanup) => {
      if (!hiddenContainer) {
        return;
      }
      hiddenContainer.addEventListener("focusin", updateSelectionOverlay);
      hiddenContainer.addEventListener("compositionend", updateSelectionOverlay);
      onCleanup(() => {
        hiddenContainer.removeEventListener("focusin", updateSelectionOverlay);
        hiddenContainer.removeEventListener("compositionend", updateSelectionOverlay);
      });
    },
    { immediate: true },
  );
  const browserWindow = typeof window === "undefined" ? null : window;
  browserWindow?.addEventListener("scroll", updateSelectionOverlay, true);
  onScopeDispose(() => {
    unsubscribeRender();
    clearOverlay();
    browserWindow?.removeEventListener("scroll", updateSelectionOverlay, true);
    resetImeCaretAnchor(opts.hiddenContainer.value);
  });

  return { clearOverlay, updateSelectionOverlay };
};
