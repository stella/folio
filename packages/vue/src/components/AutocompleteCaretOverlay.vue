<template>
  <div
    v-if="caret && text.length > 0"
    class="folio-autocomplete-overlay"
    data-folio-autocomplete-overlay=""
  >
    <span
      class="folio-autocomplete-ghost"
      :style="{
        left: `${caret.x}px`,
        top: `${caret.y}px`,
        lineHeight: `${caret.lineHeight}px`,
        maxWidth: caret.maxWidth === undefined ? undefined : `${caret.maxWidth}px`,
      }"
    >
      {{ text
      }}<span
        :class="[
          'folio-autocomplete-caret',
          { 'folio-autocomplete-caret--streaming': isStreaming },
        ]"
        aria-hidden="true"
        >stella</span
      >
    </span>
  </div>
</template>

<script setup lang="ts">
import type { AutocompleteCaretOverlayProps } from "./AutocompleteCaretOverlay.types";

defineProps<AutocompleteCaretOverlayProps>();
</script>

<style scoped>
.folio-autocomplete-overlay {
  position: absolute;
  inset: 0;
  pointer-events: none;
  overflow: visible;
}

.folio-autocomplete-ghost {
  position: absolute;
  pointer-events: none;
  user-select: none;
  white-space: pre-wrap;
  overflow-wrap: anywhere;
  color: color-mix(in oklch, var(--doc-text, currentColor) 38%, transparent);
  font: inherit;
  letter-spacing: inherit;
}

.folio-autocomplete-caret {
  display: inline-flex;
  align-items: center;
  margin-inline-start: 0.4em;
  padding: 0 0.45em;
  height: 1.1em;
  border-radius: 9999px;
  background: color-mix(in oklch, var(--doc-text, currentColor) 8%, transparent);
  color: color-mix(in oklch, var(--doc-text, currentColor) 70%, transparent);
  font-size: 0.72em;
  font-weight: 500;
  letter-spacing: 0.02em;
  text-transform: lowercase;
  vertical-align: baseline;
  pointer-events: none;
  user-select: none;
}

.folio-autocomplete-caret--streaming {
  animation: folio-autocomplete-caret-pulse 1200ms ease-in-out infinite;
}

@keyframes folio-autocomplete-caret-pulse {
  0%,
  100% {
    opacity: 0.65;
  }
  50% {
    opacity: 1;
  }
}

@media (prefers-reduced-motion: reduce) {
  .folio-autocomplete-caret--streaming {
    animation: none;
  }
}
</style>
