<!--
  Built-in, dependency-light Input used when a consumer does not inject one.
  Vue port of packages/react/src/ui/defaults/input.tsx: a thin wrapper over a
  native `<input>` that supports `v-model`. Every native `<input>` attribute
  (`type`, `placeholder`, `aria-*`, `min`/`max`/`step`, ...) falls through
  Vue's attribute inheritance and is not re-declared here.

  Exposes `focus()` / `select()` (mirroring React's `nativeInput` + forwarded
  `ref` escape hatch, used by FindReplaceDialog to focus-and-select on open):
  a Vue component `ref` resolves to the component instance, not its root DOM
  node, so a consumer wanting imperative access needs the child to
  `defineExpose` it. A host override that also needs `ref`-based `focus()` /
  `select()` (e.g. autofocus-on-open) should expose the same two methods.
-->
<template>
  <input
    ref="inputRef"
    class="folio-default-input"
    :class="[className, size && `folio-default-input--${size}`]"
    :value="modelValue"
    @input="onInput"
  />
</template>

<script setup lang="ts">
import { ref } from 'vue';

withDefaults(
  defineProps<{
    modelValue?: string | number;
    className?: string;
    size?: 'sm' | 'default' | 'lg';
  }>(),
  { size: 'default' }
);

const emit = defineEmits<{
  (e: 'update:modelValue', value: string): void;
}>();

function onInput(event: Event) {
  if (!(event.target instanceof HTMLInputElement)) return;
  emit('update:modelValue', event.target.value);
}

const inputRef = ref<HTMLInputElement | null>(null);
defineExpose({
  focus: () => inputRef.value?.focus(),
  select: () => inputRef.value?.select(),
});
</script>

<style scoped>
.folio-default-input {
  height: 28px;
  padding: 0 8px;
  font-size: 13px;
  border: 1px solid var(--doc-border-dark, #ccc);
  border-radius: 4px;
  background: var(--doc-surface, #fff);
  color: var(--doc-text, inherit);
}
.folio-default-input:focus {
  outline: none;
  border-color: var(--doc-primary, #2563eb);
}
.folio-default-input--sm {
  height: 24px;
  font-size: 12px;
}
.folio-default-input--lg {
  height: 32px;
}
</style>
