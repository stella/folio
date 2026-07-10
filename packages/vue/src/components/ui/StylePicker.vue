<!--
  Vue port of packages/react/src/components/ui/StylePicker.tsx —
  paragraph style picker (Normal / Title / Subtitle / Heading 1-3 / etc).
  Same default styles + same font-size/weight preview. Native
  <select> rather than radix-vue.
-->
<template>
  <Select
    class="docx-style-picker"
    :value="value ?? 'Normal'"
    :items="selectItems"
    :disabled="disabled"
    :class="className"
    aria-label="Paragraph style"
    @change="onSelectChange"
  />
</template>

<script setup lang="ts">
import { computed } from "vue";
import type { StyleOption } from "./StylePicker.types";
import { useFolioUI } from "../../ui/folio-ui";
import type { FolioSelectItem } from "../../ui/folio-ui";

// Resolve Select from the FolioUI injection provider so a host override
// takes effect here too (previously a native `<select>`).
const { Select } = useFolioUI();

const props = withDefaults(
  defineProps<{
    value?: string;
    styles?: StyleOption[];
    disabled?: boolean;
    className?: string;
  }>(),
  { disabled: false },
);

const emit = defineEmits<{
  (e: "change", styleId: string): void;
}>();

const DEFAULT_STYLES: StyleOption[] = [
  { styleId: "Normal", name: "Normal text" },
  { styleId: "Title", name: "Title" },
  { styleId: "Subtitle", name: "Subtitle" },
  { styleId: "Heading1", name: "Heading 1" },
  { styleId: "Heading2", name: "Heading 2" },
  { styleId: "Heading3", name: "Heading 3" },
];

const resolvedStyles = computed(() => props.styles ?? DEFAULT_STYLES);

const selectItems = computed<FolioSelectItem[]>(() =>
  resolvedStyles.value.map((s) => ({ value: s.styleId, label: s.name })),
);

function onSelectChange(value: string) {
  emit("change", value);
}
</script>

<style scoped>
.docx-style-picker {
  height: 28px;
  font-size: 13px;
  min-width: 100px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  padding: 0 6px;
  background: var(--doc-surface);
  cursor: pointer;
}
.docx-style-picker:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}
</style>
