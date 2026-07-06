<!--
  Built-in, dependency-light DatePickerPopover used when a consumer does not
  inject one. Vue port of packages/react/src/ui/defaults/date-picker-popover.tsx:
  a native `<input type="date">` whose value is an ISO `yyyy-mm-dd` string;
  emits `change` with that string (or `null` when cleared), matching the
  design-system picker contract. See ui/folio-ui.ts's module docblock — no Vue
  chrome consumer renders this yet (content-control widgets are not ported to
  Vue).
-->
<template>
  <input
    class="folio-default-date-input"
    type="date"
    :value="isoValue"
    :autofocus="defaultOpen"
    @change="onChange"
  />
</template>

<script setup lang="ts">
import { computed } from 'vue';

const props = withDefaults(
  defineProps<{
    value: string | Date | null;
    defaultOpen?: boolean;
  }>(),
  { defaultOpen: false }
);

const emit = defineEmits<{
  (e: 'change', value: string | null): void;
}>();

function toISODate(value: string | Date | null): string {
  if (value === null) return '';
  if (value instanceof Date) {
    const year = value.getFullYear().toString().padStart(4, '0');
    const month = (value.getMonth() + 1).toString().padStart(2, '0');
    const day = value.getDate().toString().padStart(2, '0');
    return `${year}-${month}-${day}`;
  }
  return value.length >= 10 ? value.slice(0, 10) : value;
}

const isoValue = computed(() => toISODate(props.value));

function onChange(event: Event) {
  if (!(event.target instanceof HTMLInputElement)) return;
  emit('change', event.target.value || null);
}
</script>

<style scoped>
.folio-default-date-input {
  height: 28px;
  padding: 0 8px;
  font-size: 13px;
  border: 1px solid var(--doc-border-dark, #ccc);
  border-radius: 4px;
  background: var(--doc-surface, #fff);
  color: var(--doc-text, inherit);
}
</style>
