<template>
  <AutocompleteCaretOverlay
    :caret="projection?.caret ?? null"
    :text="projection?.text ?? ''"
    :is-streaming="projection?.isStreaming ?? false"
  />
</template>

<script setup lang="ts">
import type { EditorState } from "prosemirror-state";
import { onBeforeUnmount, onMounted, ref, watch } from "vue";

import type { FolioAutocompleteState } from "@stll/folio-core/controller/autocompleteSuggestion";
import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import type { LayoutSelectionGate } from "@stll/folio-core/paged-layout/LayoutSelectionGate";
import { createRenderedDomContext } from "@stll/folio-core/render-dom/RenderedDomContext";

import type { AutocompleteCaretRect } from "./AutocompleteCaretOverlay.types";
import AutocompleteCaretOverlay from "./AutocompleteCaretOverlay.vue";

const MIN_GHOST_WIDTH = 48;

type AutocompleteProjection = {
  caret: AutocompleteCaretRect;
  text: string;
  isStreaming: boolean;
};

const props = defineProps<{
  getPagesContainer: () => HTMLElement | null;
  editorState: EditorState | null;
  zoom: number;
  syncCoordinator: LayoutSelectionGate;
  /** The editor's suggestion. */
  suggestion: FolioAutocompleteState;
  editorPositionAt: FolioEditor["editorPositionAt"];
}>();

const projection = ref<AutocompleteProjection | null>(null);
const renderEpoch = ref(0);
let rafId: number | null = null;
let unsubscribeRender: (() => void) | null = null;

type ResolveMaxWidthOptions = {
  pagesContainer: HTMLElement;
  x: number;
  y: number;
  zoom: number;
};

const resolveMaxWidth = ({
  pagesContainer,
  x,
  y,
  zoom,
}: ResolveMaxWidthOptions): number | undefined => {
  const safeZoom = zoom > 0 ? zoom : 1;
  const containerRect = pagesContainer.getBoundingClientRect();
  for (const page of pagesContainer.querySelectorAll<HTMLElement>(
    ".layout-page",
  )) {
    const pageRect = page.getBoundingClientRect();
    const pageTop = (pageRect.top - containerRect.top) / safeZoom;
    const pageBottom = (pageRect.bottom - containerRect.top) / safeZoom;
    if (y < pageTop || y > pageBottom) {
      continue;
    }
    const content = page.querySelector<HTMLElement>(".layout-page-content");
    if (!content) {
      return undefined;
    }
    const contentRect = content.getBoundingClientRect();
    const contentRight = (contentRect.right - containerRect.left) / safeZoom;
    return Math.max(MIN_GHOST_WIDTH, contentRight - x);
  }
  return undefined;
};

const syncProjection = (): void => {
  const pagesContainer = props.getPagesContainer();
  const state = props.editorState;
  if (!state || !pagesContainer || !props.syncCoordinator.isSafeToRender()) {
    projection.value = null;
    return;
  }

  const suggestion = props.suggestion;
  const anchor =
    suggestion.status === "idle"
      ? null
      : props.editorPositionAt(suggestion.anchor);
  if (
    suggestion.status === "idle" ||
    suggestion.text.length === 0 ||
    anchor === null
  ) {
    projection.value = null;
    return;
  }

  const context = createRenderedDomContext(pagesContainer, props.zoom);
  const coordinates = context.getCoordinatesForPosition(anchor);
  if (!coordinates) {
    projection.value = null;
    return;
  }

  const offset = context.getContainerOffset();
  const maxWidth = resolveMaxWidth({
    pagesContainer,
    x: coordinates.x,
    y: coordinates.y,
    zoom: props.zoom,
  });
  projection.value = {
    caret: {
      x: (coordinates.x + offset.x) * props.zoom,
      y: (coordinates.y + offset.y) * props.zoom,
      lineHeight: coordinates.height * props.zoom,
      maxWidth: maxWidth === undefined ? undefined : maxWidth * props.zoom,
    },
    text: suggestion.text,
    isStreaming: suggestion.status === "streaming",
  };
};

const scheduleSync = (): void => {
  if (rafId !== null) {
    cancelAnimationFrame(rafId);
  }
  rafId = requestAnimationFrame(() => {
    rafId = null;
    syncProjection();
  });
};

onMounted(() => {
  unsubscribeRender = props.syncCoordinator.onRender(() => {
    renderEpoch.value += 1;
  });
  scheduleSync();
});

onBeforeUnmount(() => {
  if (rafId !== null) {
    cancelAnimationFrame(rafId);
    rafId = null;
  }
  unsubscribeRender?.();
});

watch(
  () => [props.suggestion, props.editorState, props.zoom, renderEpoch.value],
  () => scheduleSync(),
);
</script>
