<!--
  Vue port of packages/react/src/components/ui/ZoomControl.tsx —
  a compact, controlled dropdown for choosing the editor's document zoom
  level. Same DEFAULT_ZOOM_LEVELS (50/75/100/125/150/200%) and same
  near-preset matching as React so a continuous-zoom float (e.g. 0.9999)
  still highlights the 100% preset.

  Controlled + standalone, mirroring React: it neither owns zoom state nor
  consumes `useZoom` (that composable owns the wheel/keyboard path). The host
  passes the current `value` and reacts to `@change`.
-->
<template>
  <select
    class="docx-zoom-control"
    :class="[{ 'docx-zoom-control--compact': compact }, className]"
    :value="selectValue"
    :disabled="disabled"
    aria-label="Zoom level"
    @change="onChange"
  >
    <!-- A near-preset float has no matching <option>, so surface it as a
         transient option (e.g. "97%") that stays selected until the next
         preset is picked. -->
    <option v-if="!matchingLevel" :value="selectValue">{{ displayLabel }}</option>
    <option v-for="level in resolvedLevels" :key="level.value" :value="String(level.value)">
      {{ level.label }}
    </option>
  </select>
</template>

<script lang="ts">
/** A single zoom preset: `value` is a scale factor (1 = 100%). */
export type ZoomLevel = {
  value: number;
  label: string;
};

export type ZoomControlProps = {
  /** Current zoom (1 = 100%). */
  value?: number;
  /** Override the preset levels offered in the dropdown. */
  levels?: ZoomLevel[];
  disabled?: boolean;
  className?: string;
  /** Render the trigger at the smaller toolbar-chrome size. */
  compact?: boolean;
};

const DEFAULT_ZOOM_LEVELS: ZoomLevel[] = [
  { value: 0.5, label: "50%" },
  { value: 0.75, label: "75%" },
  { value: 1, label: "100%" },
  { value: 1.25, label: "125%" },
  { value: 1.5, label: "150%" },
  { value: 2, label: "200%" },
];
</script>

<script setup lang="ts">
import { computed } from "vue";

const props = withDefaults(defineProps<ZoomControlProps>(), {
  value: 1,
  disabled: false,
  compact: false,
});

const emit = defineEmits<{
  (e: "change", zoom: number): void;
}>();

const resolvedLevels = computed(() => props.levels ?? DEFAULT_ZOOM_LEVELS);

const matchingLevel = computed(() =>
  resolvedLevels.value.find((level) => Math.abs(level.value - props.value) < 0.001),
);

const displayLabel = computed(() =>
  matchingLevel.value ? matchingLevel.value.label : `${Math.round(props.value * 100)}%`,
);

// Use the matched preset's exact string so the active item highlights even
// when `value` is a near-preset float (e.g. 0.9999 from continuous zoom).
const selectValue = computed(() =>
  matchingLevel.value ? String(matchingLevel.value.value) : String(props.value),
);

function onChange(e: Event) {
  if (!(e.target instanceof HTMLSelectElement)) return;
  const zoom = Number.parseFloat(e.target.value);
  if (!Number.isNaN(zoom)) emit("change", zoom);
}
</script>

<style scoped>
.docx-zoom-control {
  height: 32px;
  width: 70px;
  padding: 0 4px;
  font-size: 14px;
  font-variant-numeric: tabular-nums;
  color: var(--doc-text-muted);
  background: transparent;
  border: 1px solid transparent;
  border-radius: 6px;
  cursor: pointer;
  outline: none;
}
.docx-zoom-control--compact {
  height: 28px;
  width: 55px;
  font-size: 12px;
}
.docx-zoom-control:hover:not(:disabled) {
  background: var(--doc-primary-light);
  color: var(--doc-text);
}
.docx-zoom-control:focus-visible {
  border-color: var(--doc-primary);
}
.docx-zoom-control:disabled {
  cursor: not-allowed;
  opacity: 0.4;
}
</style>
