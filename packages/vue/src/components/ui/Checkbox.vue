<!--
  Built-in, dependency-light Checkbox used when a consumer does not inject
  one. Vue port of packages/react/src/ui/defaults/checkbox.tsx: a native
  `<input type="checkbox">` styled via `accent-color` (no custom SVG
  indicator — Vue, unlike React's base-ui primitive, can style a native
  checkbox reliably enough that a bespoke indicator isn't needed for the
  dependency-free fallback). `checked` + `update:checked` support
  `v-model:checked`.
-->
<template>
  <input
    type="checkbox"
    class="folio-default-checkbox"
    :class="className"
    :checked="checked"
    @change="onChange"
  />
</template>

<script setup lang="ts">
defineProps<{
  checked?: boolean;
  className?: string;
}>();

const emit = defineEmits<{
  (e: 'update:checked', value: boolean): void;
}>();

function onChange(event: Event) {
  if (!(event.target instanceof HTMLInputElement)) return;
  emit('update:checked', event.target.checked);
}
</script>

<style scoped>
.folio-default-checkbox {
  width: 14px;
  height: 14px;
  margin: 0;
  accent-color: var(--doc-primary, #2563eb);
  cursor: pointer;
}
</style>
