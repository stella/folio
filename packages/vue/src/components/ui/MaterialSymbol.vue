<!--
  MaterialSymbol — Vue twin of `packages/react/src/components/ui/Icons.tsx`.

  Single SVG component that picks paths from a registry shared with the
  React side. Path data is extracted from `Icons.tsx` via
  `scripts/extract-icons.mjs` so the two adapters can never drift on the
  paths themselves — only the host component differs.

  Use:
    <MaterialSymbol name="format_bold" :size="20" />
-->
<template>
  <svg
    v-if="paths"
    xmlns="http://www.w3.org/2000/svg"
    :width="size"
    :height="size"
    viewBox="0 -960 960 960"
    fill="currentColor"
    :class="className"
    :style="{ display: 'inline-flex', flexShrink: 0, ...(style ?? {}) }"
    aria-hidden="true"
  >
    <path v-for="(d, i) in paths" :key="i" :d="d" />
  </svg>
  <!-- Fallback for unknown icon names: render the name so the bug is visible. -->
  <span
    v-else
    :class="className"
    :style="{
      fontSize: `${size}px`,
      width: `${size}px`,
      height: `${size}px`,
      display: 'inline-flex',
      ...(style ?? {}),
    }"
  >
    {{ name }}
  </span>
</template>

<script setup lang="ts">
import { computed, type CSSProperties } from "vue";
import iconPaths from "./icon-paths.json";

const props = withDefaults(
  defineProps<{
    name: string;
    size?: number;
    className?: string;
    style?: CSSProperties;
  }>(),
  { size: 20, className: "" },
);

// JSON import is typed structurally; widen to a plain registry so the lookup
// stays a checked index access (noUncheckedIndexedAccess -> string[] | undefined).
const registry: Record<string, string[]> = iconPaths;

const paths = computed<string[] | null>(() => {
  const entry = registry[props.name];
  if (!entry) {
    if (typeof console !== "undefined") console.warn(`Icon not found: ${props.name}`);
    return null;
  }
  return entry;
});
</script>
