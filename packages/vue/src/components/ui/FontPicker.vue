<!--
  Vue port of packages/react/src/components/ui/FontPicker.tsx —
  font family picker grouped by category. Shares the FontOption
  type with the rest of the adapter via ../../utils/fontOptions.
  Native <select> with <optgroup> rather than radix-vue.
-->
<template>
  <select
    class="docx-font-picker"
    :value="displayValue"
    :disabled="disabled"
    :class="className"
    aria-label="Font family"
    @change="onChange"
  >
    <optgroup v-if="groups['sans-serif'].length" label="Sans-serif">
      <option
        v-for="f in groups['sans-serif']"
        :key="f.name"
        :value="f.name"
        :style="showPreview ? { fontFamily: f.fontFamily } : undefined"
      >
        {{ f.name }}
      </option>
    </optgroup>
    <optgroup v-if="groups['serif'].length" label="Serif">
      <option
        v-for="f in groups['serif']"
        :key="f.name"
        :value="f.name"
        :style="showPreview ? { fontFamily: f.fontFamily } : undefined"
      >
        {{ f.name }}
      </option>
    </optgroup>
    <optgroup v-if="groups['monospace'].length" label="Monospace">
      <option
        v-for="f in groups['monospace']"
        :key="f.name"
        :value="f.name"
        :style="showPreview ? { fontFamily: f.fontFamily } : undefined"
      >
        {{ f.name }}
      </option>
    </optgroup>
    <optgroup v-if="groups['other'].length" label="Other">
      <option
        v-for="f in groups['other']"
        :key="f.name"
        :value="f.name"
        :style="showPreview ? { fontFamily: f.fontFamily } : undefined"
      >
        {{ f.name }}
      </option>
    </optgroup>
  </select>
</template>

<script setup lang="ts">
import { computed } from "vue";
import type { FontOption } from "../../utils/fontOptions";
import { getPrimaryFontFamily } from "./fontPickerValue";

export type { FontOption };

const props = withDefaults(
  defineProps<{
    value?: string;
    fonts?: FontOption[];
    disabled?: boolean;
    className?: string;
    placeholder?: string;
    showPreview?: boolean;
  }>(),
  { disabled: false, placeholder: "Font", showPreview: true },
);

const emit = defineEmits<{
  (e: "change", fontFamily: string): void;
}>();

const DEFAULT_FONTS: FontOption[] = [
  { name: "Arial", fontFamily: "Arial, Helvetica, sans-serif", category: "sans-serif" },
  { name: "Calibri", fontFamily: '"Calibri", Arial, sans-serif', category: "sans-serif" },
  { name: "Helvetica", fontFamily: "Helvetica, Arial, sans-serif", category: "sans-serif" },
  { name: "Verdana", fontFamily: "Verdana, Geneva, sans-serif", category: "sans-serif" },
  { name: "Open Sans", fontFamily: '"Open Sans", sans-serif', category: "sans-serif" },
  { name: "Roboto", fontFamily: "Roboto, sans-serif", category: "sans-serif" },
  { name: "Times New Roman", fontFamily: '"Times New Roman", Times, serif', category: "serif" },
  { name: "Georgia", fontFamily: "Georgia, serif", category: "serif" },
  { name: "Cambria", fontFamily: "Cambria, Georgia, serif", category: "serif" },
  { name: "Garamond", fontFamily: "Garamond, serif", category: "serif" },
  { name: "Courier New", fontFamily: '"Courier New", monospace', category: "monospace" },
  { name: "Consolas", fontFamily: "Consolas, monospace", category: "monospace" },
];

const resolvedFonts = computed(() => props.fonts ?? DEFAULT_FONTS);
const groups = computed(() => {
  const out: Record<"sans-serif" | "serif" | "monospace" | "other", FontOption[]> = {
    "sans-serif": [],
    serif: [],
    monospace: [],
    other: [],
  };
  for (const f of resolvedFonts.value) out[f.category ?? "other"].push(f);
  return out;
});

const displayValue = computed(() => {
  const value = props.value;
  if (!value) return props.placeholder;
  const lower = value.toLowerCase();
  const m = resolvedFonts.value.find(
    (f) =>
      f.fontFamily === value ||
      f.name.toLowerCase() === lower ||
      getPrimaryFontFamily(f.fontFamily).toLowerCase() === lower,
  );
  return m?.name ?? value;
});

function onChange(e: Event) {
  if (!(e.target instanceof HTMLSelectElement)) return;
  const name = e.target.value;
  const font = resolvedFonts.value.find((f) => f.name === name);
  emit("change", font ? getPrimaryFontFamily(font.fontFamily) || font.name : name);
}
</script>

<style scoped>
.docx-font-picker {
  height: 28px;
  font-size: 13px;
  min-width: 100px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  padding: 0 6px;
  background: var(--doc-surface);
  cursor: pointer;
}
.docx-font-picker:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}
</style>
