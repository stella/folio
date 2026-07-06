<!--
  Vue document outline panel — mirrors React's `<DocumentOutline>`
  (DocumentOutline.tsx). Slides in from the left edge of the viewport,
  overlays the editor without consuming layout space, uses the same
  240px width / arrow_back icon / title text / empty-state copy as
  React so the two adapters look identical.
-->
<template>
  <nav
    v-if="isOpen"
    class="doc-outline"
    :style="{ left: leftOffset + 'px', top: topPx + 'px' }"
    role="navigation"
    aria-label="Document outline"
    @mousedown.stop
  >
    <div class="doc-outline__header">
      <button
        class="doc-outline__back"
        :title="'Close outline'"
        :aria-label="'Close outline'"
        @click="$emit('close')"
      >
        <MaterialSymbol name="arrow_back" :size="20" />
      </button>
      <span class="doc-outline__title">Document Outline</span>
    </div>
    <div class="doc-outline__body">
      <div v-if="headings.length === 0" class="doc-outline__empty">No headings found</div>
      <OutlineRail
        v-else
        :items="items"
        :get-scroll-container="getScrollContainer ?? (() => null)"
        :on-jump="handleJump"
        aria-label="Document outline"
      />
    </div>
  </nav>
</template>

<script setup lang="ts">
import { computed } from "vue";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";
import MaterialSymbol from "./ui/MaterialSymbol.vue";
import type { OutlineItem } from "../ui/folio-ui";
import { useFolioUI } from "../ui/folio-ui";

const props = withDefaults(
  defineProps<{
    isOpen: boolean;
    headings: HeadingInfo[];
    /** Left anchor (px); host bumps it past the vertical ruler when shown. */
    leftOffset?: number;
    /** Top anchor (px); host bumps it past the sticky ruler row when shown. */
    topPx?: number;
    /** Getter for the scroll container, forwarded to the injected OutlineRail's
     *  `onJump` callback; DocxEditor.vue passes `() => pagesRef` (the same
     *  getter idiom `DecorationLayer.vue`'s `getPagesContainer` uses — Vue's
     *  template compiler auto-unwraps a bare ref, even inside an inline arrow
     *  function, so the getter reads the current element at click-time
     *  instead of passing the ref object across the prop boundary). Optional
     *  so the component still renders standalone. */
    getScrollContainer?: () => HTMLElement | null;
  }>(),
  { leftOffset: 12, topPx: 24 },
);

const { OutlineRail } = useFolioUI();

// Indent relative to the shallowest heading present (mirrors React) so a doc
// whose top sections are Heading 2 doesn't carry a phantom first-level indent.
const minLevel = computed(() =>
  props.headings.length ? Math.min(...props.headings.map((h) => h.level)) : 0,
);

const items = computed<OutlineItem[]>(() =>
  props.headings.map((h) => ({
    id: String(h.pmPos),
    label: h.text || "(untitled)",
    level: h.level - minLevel.value,
  })),
);

const emit = defineEmits<{
  (e: "close"): void;
  (e: "navigate", pmPos: number): void;
}>();

function handleJump(id: string) {
  emit("navigate", Number(id));
}
</script>

<style scoped>
/* Matches React DocumentOutline.tsx: position: absolute against the
   editor host, anchored 12px from the left (= the collapsed toggle's
   offset so the back arrow lands where the toggle was), 240px wide,
   full height. The wrapping `__editor-area` has position: relative so
   this lands on top of the page area without consuming flex space. The
   slide-in uses transform so it doesn't trigger layout. */
.doc-outline {
  position: absolute;
  /* `left` and `top` are set inline by the host: `left` via the leftOffset prop
     (page left anchor, bumped when the vertical ruler is shown), `top` at the
     page top edge, bumped past the sticky ruler row when one is shown. */
  bottom: 0;
  width: 240px;
  display: flex;
  flex-direction: column;
  font-family: 'Google Sans', Roboto, Arial, sans-serif;
  z-index: 40;
  animation: docOutlineIn 0.15s ease-out;
}
@keyframes docOutlineIn {
  /* Large enough to fully hide the 240px panel at any left anchor (12 or,
     when the vertical ruler is shown, 32 → right edge 272). */
  from {
    transform: translateX(-300px);
  }
  to {
    transform: translateX(0);
  }
}
/* No left padding so the back arrow sits at the nav anchor (= the
   collapsed toggle's position). */
.doc-outline__header {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 16px 16px 12px 0;
}
.doc-outline__back {
  background: none;
  border: none;
  cursor: pointer;
  padding: 4px;
  border-radius: 50%;
  display: flex;
  align-items: center;
  color: var(--doc-text-muted);
}
.doc-outline__back:hover {
  background: var(--doc-shadow-subtle);
}
.doc-outline__title {
  font-weight: 400;
  font-size: 14px;
  color: var(--doc-text);
  letter-spacing: 0.01em;
}
.doc-outline__body {
  flex: 1;
  overflow-y: auto;
  padding-left: 4px;
}
.doc-outline__empty {
  padding: 8px 16px;
  color: var(--doc-text-subtle);
  font-size: 13px;
  line-height: 20px;
}
</style>
