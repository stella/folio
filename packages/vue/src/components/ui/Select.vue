<!--
  Built-in, dependency-light Select used when a consumer does not inject one.
  Vue port of packages/react/src/ui/defaults/select.tsx, collapsed into one
  native-`<select>`-backed component driven by a flat `items` array (Vue's
  contract has no Root/Trigger/Popup/Item part-object — see ui/folio-ui.ts's
  module docblock). No `<optgroup>` support; chrome consumers that need
  grouping (e.g. the font picker) stay on a native `<select>` directly.
-->
<template>
  <select
    class="folio-default-select"
    :class="className"
    :disabled="disabled"
    :value="value"
    @change="onChange"
  >
    <option v-if="placeholder && value === undefined" disabled value="">
      {{ placeholder }}
    </option>
    <option v-for="item in items" :key="item.value" :disabled="item.disabled" :value="item.value">
      {{ item.label }}
    </option>
  </select>
</template>

<script setup lang="ts">
import type { FolioSelectItem } from "../../ui/folio-ui";

defineProps<{
  value?: string;
  items: FolioSelectItem[];
  disabled?: boolean;
  placeholder?: string;
  className?: string;
}>();

const emit = defineEmits<{
  (e: "change", value: string): void;
}>();

function onChange(event: Event) {
  if (!(event.target instanceof HTMLSelectElement)) return;
  emit("change", event.target.value);
}
</script>

<style scoped>
.folio-default-select {
  height: 28px;
  padding: 0 6px;
  font-size: 13px;
  border: 1px solid var(--doc-border-dark, #ccc);
  border-radius: 4px;
  background: var(--doc-surface, #fff);
  color: var(--doc-text, inherit);
  cursor: pointer;
}
.folio-default-select:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}
</style>
