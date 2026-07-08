<template>
  <FolioDialog
    :open="isOpen"
    aria-label="Split Cell"
    :style="{ minWidth: '360px', maxWidth: '520px', width: '100%' }"
    @close="$emit('close')"
    @keydown.enter="split"
  >
    <div class="dialog__header">Split Cell</div>
    <div class="dialog__body">
      <div class="grid">
        <label class="field">
          <span class="field__label">Columns</span>
          <input
            v-model.number="columns"
            class="field__input"
            :max="MAX_PARTS"
            :min="MIN_PARTS"
            type="number"
          />
        </label>
        <label class="field">
          <span class="field__label">Rows</span>
          <input
            v-model.number="rows"
            class="field__input"
            :max="MAX_PARTS"
            :min="MIN_PARTS"
            type="number"
          />
        </label>
      </div>
      <label class="check">
        <input v-model="mergeBeforeSplit" type="checkbox" />
        <span>Merge selected cells before splitting</span>
      </label>
    </div>
    <div class="dialog__footer">
      <button class="dialog__btn" @click="$emit('close')">Cancel</button>
      <button class="dialog__btn dialog__btn--primary" :disabled="!canSplit" @click="split">
        Split
      </button>
    </div>
  </FolioDialog>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue";
import { useFolioUI } from "../../ui/folio-ui";

const { Dialog: FolioDialog } = useFolioUI();

export type SplitCellDialogData = {
  rows: number;
  columns: number;
  mergeBeforeSplit: boolean;
};

const MIN_PARTS = 1;
const MAX_PARTS = 63;

const props = withDefaults(
  defineProps<{
    isOpen: boolean;
    defaultRows?: number;
    defaultColumns?: number;
  }>(),
  { defaultRows: 1, defaultColumns: 2 },
);

const emit = defineEmits<{
  (e: "close"): void;
  (e: "split", data: SplitCellDialogData): void;
}>();

const rows = ref(props.defaultRows);
const columns = ref(props.defaultColumns);
const mergeBeforeSplit = ref(false);

const normalizedRows = computed(() => clampInteger(rows.value));
const normalizedColumns = computed(() => clampInteger(columns.value));
const canSplit = computed(() => normalizedRows.value > 1 || normalizedColumns.value > 1);

watch(
  () => props.isOpen,
  (open) => {
    if (!open) return;
    rows.value = props.defaultRows;
    columns.value = props.defaultColumns;
    mergeBeforeSplit.value = false;
  },
);

function split() {
  if (!canSplit.value) return;
  emit("split", {
    rows: normalizedRows.value,
    columns: normalizedColumns.value,
    mergeBeforeSplit: mergeBeforeSplit.value,
  });
  emit("close");
}

function clampInteger(value: number): number {
  return Math.min(MAX_PARTS, Math.max(MIN_PARTS, Math.trunc(value)));
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
.field__label,
.check {
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
.check {
  display: flex;
  align-items: center;
  gap: 8px;
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
