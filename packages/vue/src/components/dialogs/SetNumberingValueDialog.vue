<!--
  Set Numbering Value: start the selected list item's list over at a value.
  Mirrors React's SetNumberingValueDialog.
-->
<template>
  <FolioDialog
    :open="isOpen"
    :aria-label="t('dialogs.setNumberingValue.title')"
    :style="{ minWidth: '320px', maxWidth: '420px', width: '100%' }"
    @close="$emit('close')"
    @keydown.enter="apply"
  >
    <div class="dialog__header">{{ t("dialogs.setNumberingValue.title") }}</div>
    <div class="dialog__body">
      <label class="field">
        <span class="field__label">{{ t("dialogs.setNumberingValue.valueLabel") }}</span>
        <input
          v-model.number="value"
          class="field__input"
          :max="MAX_VALUE"
          :min="MIN_VALUE"
          type="number"
        />
      </label>
    </div>
    <div class="dialog__footer">
      <button class="dialog__btn" @click="$emit('close')">{{ t("common.cancel") }}</button>
      <button class="dialog__btn dialog__btn--primary" @click="apply">
        {{ t("common.apply") }}
      </button>
    </div>
  </FolioDialog>
</template>

<script setup lang="ts">
import { ref, watch } from "vue";

import { useTranslation } from "../../i18n";
import { useFolioUI } from "../../ui/folio-ui";

const { Dialog: FolioDialog } = useFolioUI();
const { t } = useTranslation();

/** `w:startOverride` is a non-negative decimal; the input stops far below its limit. */
const MIN_VALUE = 0;
const MAX_VALUE = 32767;
const DEFAULT_VALUE = 1;

const props = defineProps<{ isOpen: boolean }>();

const emit = defineEmits<{
  (e: "close"): void;
  (e: "apply", value: number): void;
}>();

const value = ref(DEFAULT_VALUE);

watch(
  () => props.isOpen,
  (open) => {
    if (open) value.value = DEFAULT_VALUE;
  },
);

function apply() {
  const number = Number.isFinite(value.value) ? Math.trunc(value.value) : MIN_VALUE;
  emit("apply", Math.min(MAX_VALUE, Math.max(MIN_VALUE, number)));
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
</style>
