<template></template>

<script setup lang="ts">
import { onScopeDispose, watch } from "vue";
import type { EditorView } from "prosemirror-view";

import {
  resetImeCaretAnchor,
  syncImeCaretAnchor,
} from "@stll/folio-core/layout-bridge/dom/imeCaretAnchor";
import { PAINTER_PAINTED_EVENT } from "@stll/folio-core/layout-painter/renderPage";
import { HeaderFooterSelectionOverlay } from "@stll/folio-core/render-dom/HeaderFooterSelectionOverlay";
import type { HeaderFooterSelection } from "@stll/folio-core/render-dom/HeaderFooterSelectionOverlay";

const props = defineProps<{
  pagesContainer: HTMLElement | null;
  hiddenContainer: HTMLElement | null;
  editorView: EditorView | null;
  selection: HeaderFooterSelection | null;
  zoom: number;
}>();

const overlay = new HeaderFooterSelectionOverlay();
const sync = (): void => {
  if (!props.pagesContainer) return;
  overlay.sync(props.pagesContainer, props.selection, props.zoom);
  const paintedCaret = props.pagesContainer.querySelector<HTMLElement>("[data-testid='hf-caret']");
  const caretRect = paintedCaret?.getBoundingClientRect();
  syncImeCaretAnchor({
    hiddenHost: props.hiddenContainer,
    editorView: props.editorView,
    visibleCaret: caretRect ? { left: caretRect.left, top: caretRect.top } : null,
  });
};

watch(
  () => props.pagesContainer,
  (container, _previous, onCleanup) => {
    if (!container) return;
    container.addEventListener(PAINTER_PAINTED_EVENT, sync);
    sync();
    onCleanup(() => {
      container.removeEventListener(PAINTER_PAINTED_EVENT, sync);
      overlay.clear(container);
    });
  },
  { immediate: true },
);

watch(
  () => [
    props.selection?.rId,
    props.selection?.kind,
    props.selection?.from,
    props.selection?.to,
    props.selection?.pageNumber,
    props.hiddenContainer,
    props.editorView,
    props.zoom,
  ],
  sync,
);

watch(
  () => props.hiddenContainer,
  (hiddenContainer, _previous, onCleanup) => {
    if (!hiddenContainer) return;
    hiddenContainer.addEventListener("focusin", sync);
    hiddenContainer.addEventListener("compositionend", sync);
    onCleanup(() => {
      hiddenContainer.removeEventListener("focusin", sync);
      hiddenContainer.removeEventListener("compositionend", sync);
    });
  },
  { immediate: true },
);

let animationFrame: number | null = null;
const scheduleSync = (): void => {
  if (animationFrame !== null) cancelAnimationFrame(animationFrame);
  animationFrame = requestAnimationFrame(() => {
    animationFrame = null;
    sync();
  });
};
const browserWindow = typeof window === "undefined" ? null : window;
browserWindow?.addEventListener("scroll", scheduleSync, true);
onScopeDispose(() => {
  if (animationFrame !== null) cancelAnimationFrame(animationFrame);
  browserWindow?.removeEventListener("scroll", scheduleSync, true);
  resetImeCaretAnchor(props.hiddenContainer);
});
</script>
