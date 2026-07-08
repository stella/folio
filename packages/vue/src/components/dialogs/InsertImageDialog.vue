<template>
  <FolioDialog
    :open="isOpen"
    aria-label="Insert Image"
    :style="{ minWidth: '360px', maxWidth: '520px', width: '100%' }"
    @close="$emit('close')"
  >
    <div class="dialog__header">Insert Image</div>
    <div class="dialog__body">
      <label class="field">
        <span class="field__label">Image file</span>
        <input class="field__input" type="file" :accept="accept" @change="onFileChange" />
      </label>
      <label class="field">
        <span class="field__label">Alt text</span>
        <input v-model="alt" class="field__input" :placeholder="file?.name ?? ''" />
      </label>
      <div class="grid">
        <label class="field">
          <span class="field__label">Width</span>
          <input v-model="width" class="field__input" min="1" placeholder="Auto" type="number" />
        </label>
        <label class="field">
          <span class="field__label">Height</span>
          <input v-model="height" class="field__input" min="1" placeholder="Auto" type="number" />
        </label>
      </div>
    </div>
    <div class="dialog__footer">
      <button class="dialog__btn" @click="$emit('close')">Cancel</button>
      <button class="dialog__btn dialog__btn--primary" :disabled="!file" @click="insert">
        Insert
      </button>
    </div>
  </FolioDialog>
</template>

<script setup lang="ts">
import { ref, watch } from "vue";
import { useFolioUI } from "../../ui/folio-ui";

const { Dialog: FolioDialog } = useFolioUI();

export type InsertImageDialogData = {
  file: File;
  alt?: string;
  width?: number;
  height?: number;
};

const props = withDefaults(
  defineProps<{
    isOpen: boolean;
    accept?: string;
  }>(),
  { accept: "image/png,image/jpeg,image/gif,image/webp" },
);

const emit = defineEmits<{
  (e: "close"): void;
  (e: "insert", data: InsertImageDialogData): void;
}>();

const file = ref<File | null>(null);
const alt = ref("");
const width = ref("");
const height = ref("");

watch(
  () => props.isOpen,
  (open) => {
    if (!open) return;
    file.value = null;
    alt.value = "";
    width.value = "";
    height.value = "";
  },
);

function onFileChange(event: Event) {
  const input = event.target instanceof HTMLInputElement ? event.target : null;
  file.value = input?.files?.[0] ?? null;
}

function insert() {
  if (!file.value) return;
  const parsedWidth = toPositiveNumber(width.value);
  const parsedHeight = toPositiveNumber(height.value);
  emit("insert", {
    file: file.value,
    ...(alt.value.trim() ? { alt: alt.value.trim() } : {}),
    ...(parsedWidth ? { width: parsedWidth } : {}),
    ...(parsedHeight ? { height: parsedHeight } : {}),
  });
  emit("close");
}

function toPositiveNumber(value: string): number | undefined {
  const parsed = Number(value);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : undefined;
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
.grid {
  display: grid;
  grid-template-columns: minmax(0, 1fr) minmax(0, 1fr);
  gap: 12px;
}
.field {
  display: flex;
  flex-direction: column;
  gap: 4px;
  min-width: 0;
}
.field__label {
  font-size: 13px;
  color: var(--doc-text-muted);
}
.field__input {
  border: 1px solid var(--doc-border);
  border-radius: 4px;
  background: var(--doc-surface);
  color: var(--doc-text);
  padding: 6px 8px;
  font-size: 13px;
  outline: none;
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
.dialog__btn:disabled {
  cursor: not-allowed;
  opacity: 0.5;
}
</style>
