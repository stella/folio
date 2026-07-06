<!--
  Anonymization Rects Overlay (Vue)

  Paints anonymization-term highlights on top of the rendered paged document.
  Decoration spans live in the hidden ProseMirror editor and never reach the
  visible page DOM, so we project the same ranges onto container-relative
  rectangles (via core's `projectRangesToRects`) and stack a coloured span per
  line on this absolutely-positioned overlay.

  Port of the React `AnonymizationRectsOverlay`. The projection returns unscaled
  (page-space) rects — the Vue pages carry `transform: scale(zoom)` but this
  overlay does not, so every coordinate is multiplied back by `zoom` at render
  time (the same convention as CommentMarginMarkers.vue).

  Pointer events stay disabled on every node so text selection, caret placement,
  and the editor's own handlers keep working through the overlay. The click
  hit-test instead listens on `document` and bounding-rect-tests the click
  position against the overlay spans, so the editor still receives the click and
  this fires alongside it.
-->
<template>
  <div ref="rootRef" class="folio-anonymization-overlay" data-folio-anonymization-overlay="">
    <template v-for="group in groups" :key="`${group.label}:${group.canonical}`">
      <span
        v-for="(rect, idx) in group.rects"
        :key="`${group.canonical}:${idx}:${rect.left}:${rect.top}`"
        :class="[
          'folio-anonymization-term',
          `folio-anonymization-term--${slugAnonymizationLabel(group.label)}`,
        ]"
        :data-folio-anonymization-label="group.label"
        :data-folio-anonymization-canonical="group.canonical"
        :data-folio-anonymization-selected="
          group.canonical === selectedCanonical ? 'true' : undefined
        "
        :title="`Anonymized: ${group.canonical}`"
        :style="{
          position: 'absolute',
          left: `${rect.left}px`,
          top: `${rect.top}px`,
          width: `${rect.width}px`,
          height: `${rect.height}px`,
        }"
      />
    </template>
  </div>
</template>

<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref, watch } from "vue";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import type { FlowBlock, Layout, Measure } from "@stll/folio-core/layout-engine/types";
import { projectRangesToRects } from "@stll/folio-core/paged-layout/rangeProjection";
import { prefersReducedMotionBehavior } from "@stll/folio-core/paged-layout/scrollNavigation";
import type { AnonymizationMatch } from "@stll/folio-core/prosemirror/plugins/anonymizationDecorations";
import {
  getAnonymizationMatches,
  slugAnonymizationLabel,
} from "@stll/folio-core/prosemirror/plugins/anonymizationDecorations";

/** One canonical's projected rectangles in overlay (post-zoom) pixel space. */
type AnonymizationRectGroup = {
  rects: { left: number; top: number; width: number; height: number }[];
  label: string;
  canonical: string;
};

const props = defineProps<{
  getView: () => EditorView | null;
  getPagesContainer: () => HTMLElement | null;
  /** Reactive editor state — changes on every transaction, including the
   *  meta-only term push that a doc edit does not cover. */
  editorState: EditorState | null;
  zoom: number;
  layout: Layout | null;
  blocks: FlowBlock[];
  measures: Measure[];
  onTermClick?: ((canonical: string, label: string) => void) | undefined;
  selectedCanonical?: string | null | undefined;
  selectionSeq?: number | undefined;
}>();

const rootRef = ref<HTMLDivElement | null>(null);
const groups = ref<AnonymizationRectGroup[]>([]);

// Monotonic guard: a stale async projection must not overwrite a newer one.
let requestSeq = 0;
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
  const seq = ++requestSeq;
  const view = props.getView();
  const pagesContainer = props.getPagesContainer();
  const state = props.editorState;
  if (!view || !pagesContainer || !state) {
    groups.value = [];
    return;
  }
  const matches = getAnonymizationMatches(state);
  if (matches.length === 0) {
    groups.value = [];
    return;
  }
  const projected = await projectRangesToRects<AnonymizationMatch>(matches, {
    pagesContainer,
    zoom: props.zoom,
    layout: props.layout,
    blocks: props.blocks,
    measures: props.measures,
  });
  if (seq !== requestSeq) {
    return;
  }
  const z = props.zoom;
  groups.value = projected.map(({ range, rects }) => ({
    label: range.label,
    canonical: range.canonical,
    rects: rects.map((rect) => ({
      left: rect.x * z,
      top: rect.y * z,
      width: rect.width * z,
      height: rect.height * z,
    })),
  }));
}

watch(
  () => [props.editorState, props.zoom, props.layout, props.blocks, props.measures],
  () => scheduleProject(),
);

// ---- Selection scroll-into-view -----------------------------------------
// Repeated sidebar selections of the same canonical step through every hit,
// wrapping at the end; a new canonical resets to the first occurrence.
let cycle: { canonical: string; index: number } | null = null;

function scrollSelectedIntoView(): void {
  const root = rootRef.value;
  const canonical = props.selectedCanonical;
  if (!root || !canonical) {
    cycle = null;
    return;
  }
  const spans = root.querySelectorAll<HTMLElement>(
    `[data-folio-anonymization-canonical="${CSS.escape(canonical)}"]`,
  );
  if (spans.length === 0) {
    return;
  }
  const nextIndex = cycle?.canonical === canonical ? (cycle.index + 1) % spans.length : 0;
  cycle = { canonical, index: nextIndex };
  spans.item(nextIndex)?.scrollIntoView({
    block: "center",
    behavior: prefersReducedMotionBehavior(),
  });
}

watch(
  () => [props.selectedCanonical, props.selectionSeq],
  () => scrollSelectedIntoView(),
  // Query the DOM after the overlay has (re)rendered, matching the React
  // adapter's post-commit `useEffect` timing.
  { flush: "post" },
);

// ---- Document-level click hit-test --------------------------------------
function handleDocumentClick(event: MouseEvent): void {
  if (event.button !== 0 || !props.onTermClick) {
    return;
  }
  const root = rootRef.value;
  if (!root) {
    return;
  }
  const spans = root.querySelectorAll<HTMLElement>("[data-folio-anonymization-canonical]");
  // Reverse paint order so the topmost rect wins when several overlap.
  for (let i = spans.length - 1; i >= 0; i -= 1) {
    const el = spans.item(i);
    if (!el) {
      continue;
    }
    const rect = el.getBoundingClientRect();
    const inside =
      event.clientX >= rect.left &&
      event.clientX < rect.right &&
      event.clientY >= rect.top &&
      event.clientY < rect.bottom;
    if (inside) {
      const canonical = el.dataset["folioAnonymizationCanonical"];
      const label = el.dataset["folioAnonymizationLabel"];
      if (canonical && label) {
        props.onTermClick(canonical, label);
      }
      return;
    }
  }
}

onMounted(() => {
  document.addEventListener("click", handleDocumentClick);
  scheduleProject();
});

onBeforeUnmount(() => {
  document.removeEventListener("click", handleDocumentClick);
  if (rafId !== null) {
    cancelAnimationFrame(rafId);
    rafId = null;
  }
});
</script>

<style scoped>
.folio-anonymization-overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
}

/* Single thin blue underline under every span the redaction pipeline would
   redact. No fill so the doc stays readable. Mirrors React editor.css. */
.folio-anonymization-term {
  background-color: transparent;
  box-shadow: inset 0 -2px 0 0 rgb(37 99 235 / 0.8);
  border-radius: 2px;
  transition:
    background-color 150ms ease,
    box-shadow 150ms ease;
}

.folio-anonymization-term:hover {
  background-color: rgb(37 99 235 / 0.12);
}

.folio-anonymization-term[data-folio-anonymization-selected="true"] {
  background-color: rgb(37 99 235 / 0.2);
  box-shadow:
    inset 0 -2px 0 0 rgb(37 99 235 / 0.9),
    0 0 0 2px rgb(37 99 235 / 0.35);
}
</style>
