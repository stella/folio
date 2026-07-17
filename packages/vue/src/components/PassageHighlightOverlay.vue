<!--
  Passage Highlight Overlay (Vue)

  Paints a single persistent, translucent highlight over a text passage a
  consumer opened the document at (a citation chip's quoted passage, a
  find-in-document hit, an agent tool). The highlighted range is resolved in the
  hidden ProseMirror editor and pushed in imperatively via the ref API; this
  overlay projects it onto container-relative rectangles (via core's
  `projectRangesToRects`) and stacks one span per line rectangle.

  Port of the React `PassageHighlightOverlay`. The projection returns unscaled
  (page-space) rects — the Vue pages carry `transform: scale(zoom)` but this
  overlay does not, so every coordinate is multiplied back by `zoom` at render
  time (the same convention as AnonymizationRectsOverlay.vue).

  Pointer events stay disabled so the caret, selection, and the editor's own
  handlers keep working through the overlay. The highlight is ephemeral view
  state: the parent (DocxEditor.vue) clears the range on any doc-changing
  transaction and via the explicit clear method, so this overlay never has to
  reason about the document itself.
-->
<template>
  <div class="folio-passage-highlight-overlay" data-folio-passage-highlight-overlay="">
    <span
      v-for="(rect, idx) in rects"
      :key="`${idx}:${rect.left}:${rect.top}`"
      class="folio-passage-highlight"
      :style="{
        position: 'absolute',
        left: `${rect.left}px`,
        top: `${rect.top}px`,
        width: `${rect.width}px`,
        height: `${rect.height}px`,
      }"
    />
  </div>
</template>

<script setup lang="ts">
import { onBeforeUnmount, ref, watch } from "vue";

import { createLatestRequestGate } from "@stll/folio-core/controller/latestRequestGate";
import type { FlowBlock, Layout, Measure } from "@stll/folio-core/layout-engine/types";
import { projectRangesToRects } from "@stll/folio-core/paged-layout/rangeProjection";

/** Projected rectangle in overlay (post-zoom) pixel space. */
type OverlayRect = { left: number; top: number; width: number; height: number };

const props = defineProps<{
  /** PM range to highlight, or null to paint nothing. */
  range: { from: number; to: number } | null;
  getPagesContainer: () => HTMLElement | null;
  zoom: number;
  layout: Layout | null;
  blocks: FlowBlock[];
  measures: Measure[];
}>();

const rects = ref<OverlayRect[]>([]);

const requestGate = createLatestRequestGate();
let rafId: number | null = null;

function scheduleProject(): void {
  if (rafId !== null) {
    cancelAnimationFrame(rafId);
  }
  rafId = requestAnimationFrame(() => {
    rafId = null;
    void project();
  });
}

async function project(): Promise<void> {
  const isCurrentRequest = requestGate.begin();
  const range = props.range;
  const pagesContainer = props.getPagesContainer();
  if (range === null || !pagesContainer) {
    rects.value = [];
    return;
  }
  const projected = await projectRangesToRects([range], {
    pagesContainer,
    zoom: props.zoom,
    layout: props.layout,
    blocks: props.blocks,
    measures: props.measures,
  });
  if (!isCurrentRequest()) {
    return;
  }
  const z = props.zoom;
  rects.value = (projected.at(0)?.rects ?? []).map((rect) => ({
    left: rect.x * z,
    top: rect.y * z,
    width: rect.width * z,
    height: rect.height * z,
  }));
}

watch(
  () => [props.range, props.zoom, props.layout, props.blocks, props.measures],
  () => scheduleProject(),
  { immediate: true },
);

onBeforeUnmount(() => {
  requestGate.invalidate();
  if (rafId !== null) {
    cancelAnimationFrame(rafId);
    rafId = null;
  }
});
</script>

<style scoped>
.folio-passage-highlight-overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

/* Persistent translucent wash over the highlighted passage. Distinct from the
   yellow paragraph flash and the blue anonymization underline; a soft accent
   fill so the quoted span stays visible while the text underneath stays
   legible. Mirrors React editor.css. */
.folio-passage-highlight {
  --folio-passage-highlight-color: rgba(64, 158, 255, 0.28);
  background-color: var(--folio-passage-highlight-color);
  border-radius: 2px;
  box-shadow: inset 0 0 0 1px
    color-mix(in oklch, var(--folio-passage-highlight-color) 55%, transparent);
}
</style>
