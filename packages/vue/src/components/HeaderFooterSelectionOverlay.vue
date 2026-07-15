<template></template>

<script setup lang="ts">
import { watch } from "vue";

import { PAINTER_PAINTED_EVENT } from "@stll/folio-core/layout-painter/renderPage";
import { HeaderFooterSelectionOverlay } from "@stll/folio-core/render-dom/HeaderFooterSelectionOverlay";
import type { HeaderFooterSelection } from "@stll/folio-core/render-dom/HeaderFooterSelectionOverlay";

const props = defineProps<{
  pagesContainer: HTMLElement | null;
  selection: HeaderFooterSelection | null;
  zoom: number;
}>();

const overlay = new HeaderFooterSelectionOverlay();
const sync = (): void => {
  if (!props.pagesContainer) return;
  overlay.sync(props.pagesContainer, props.selection, props.zoom);
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
    props.zoom,
  ],
  sync,
);
</script>
