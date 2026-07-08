<template>
  <FolioDialog
    :open="isOpen"
    aria-label="Paste Special"
    :style="{ minWidth: '360px', maxWidth: '520px', width: '100%' }"
    @close="$emit('close')"
  >
    <div class="dialog__header">Paste Special</div>
    <div class="dialog__body">
      <fieldset class="options">
        <legend class="options__legend">Paste mode</legend>
        <label v-for="option in PASTE_MODES" :key="option.value" class="option">
          <input v-model="mode" type="radio" :value="option.value" />
          <span class="option__body">
            <span class="option__label">{{ option.label }}</span>
            <span class="option__description">{{ option.description }}</span>
          </span>
        </label>
      </fieldset>
    </div>
    <div class="dialog__footer">
      <button class="dialog__btn" @click="$emit('close')">Cancel</button>
      <button class="dialog__btn dialog__btn--primary" @click="paste">Paste</button>
    </div>
  </FolioDialog>
</template>

<script setup lang="ts">
import { ref, watch } from "vue";
import { useFolioUI } from "../../ui/folio-ui";

const { Dialog: FolioDialog } = useFolioUI();

export type PasteSpecialMode = "keepFormatting" | "mergeFormatting" | "plainText";

const PASTE_MODES: Array<{ value: PasteSpecialMode; label: string; description: string }> = [
  {
    value: "keepFormatting",
    label: "Keep source formatting",
    description: "Preserve styles, links, and inline formatting from the clipboard.",
  },
  {
    value: "mergeFormatting",
    label: "Merge formatting",
    description: "Keep useful inline formatting while matching the destination paragraph.",
  },
  {
    value: "plainText",
    label: "Unformatted text",
    description: "Paste text only and use the destination formatting.",
  },
];

const props = withDefaults(
  defineProps<{
    isOpen: boolean;
    defaultMode?: PasteSpecialMode;
  }>(),
  { defaultMode: "plainText" },
);

const emit = defineEmits<{
  (e: "close"): void;
  (e: "paste", mode: PasteSpecialMode): void;
}>();

const mode = ref<PasteSpecialMode>(props.defaultMode);

watch(
  () => props.isOpen,
  (open) => {
    if (open) mode.value = props.defaultMode;
  },
);

function paste() {
  emit("paste", mode.value);
  emit("close");
}
</script>

<style scoped>
.dialog__header {
  padding: 16px 20px 12px;
  border-bottom: 1px solid var(--doc-border);
  font-size: 16px;
  font-weight: 600;
  color: var(--doc-text);
}
.dialog__body {
  padding: 16px 20px;
  display: flex;
  flex-direction: column;
  gap: 16px;
}
.dialog__footer {
  padding: 12px 20px 16px;
  border-top: 1px solid var(--doc-border);
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
.options {
  display: flex;
  flex-direction: column;
  gap: 12px;
  padding: 0;
  margin: 0;
  border: 0;
}
.options__legend {
  margin-bottom: 4px;
  font-size: 13px;
  color: var(--doc-text-muted);
}
.option {
  display: flex;
  align-items: flex-start;
  gap: 12px;
  color: var(--doc-text);
}
.option__body {
  display: flex;
  flex-direction: column;
  gap: 4px;
}
.option__label {
  font-size: 13px;
  font-weight: 600;
}
.option__description {
  font-size: 12px;
  color: var(--doc-text-muted);
}
.dialog__btn {
  border: 1px solid var(--doc-border);
  border-radius: 4px;
  background: var(--doc-surface);
  color: var(--doc-text);
  padding: 6px 16px;
  font-size: 13px;
  cursor: pointer;
}
.dialog__btn--primary {
  border-color: var(--doc-primary);
  background: var(--doc-primary);
  color: white;
  font-weight: 600;
}
</style>
