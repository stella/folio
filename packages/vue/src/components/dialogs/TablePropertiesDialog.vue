<!--
  Table Properties dialog — ported from the upstream docx-editor Vue adapter.
-->
<template>
  <FolioDialog
    :open="isOpen"
    :aria-label="t('dialogs.tableProperties.title')"
    :style="{ minWidth: '360px', maxWidth: '440px', width: '100%' }"
    @close="$emit('close')"
    @keydown.enter="apply"
  >
    <div class="tpd-dialog__header">{{ t("dialogs.tableProperties.title") }}</div>
    <div class="tpd-dialog__body">
      <div class="tpd-row">
        <label class="tpd-label">{{ t("dialogs.tableProperties.widthType") }}</label>
        <select v-model="widthType" class="tpd-select">
          <option value="auto">{{ t("dialogs.tableProperties.widthTypes.auto") }}</option>
          <option value="dxa">{{ t("dialogs.tableProperties.widthTypes.fixed") }}</option>
          <option value="pct">{{ t("dialogs.tableProperties.widthTypes.percentage") }}</option>
        </select>
      </div>
      <div v-if="widthType !== 'auto'" class="tpd-row">
        <label class="tpd-label">{{ t("dialogs.tableProperties.widthLabel") }}</label>
        <input
          v-model.number="width"
          type="number"
          class="tpd-input"
          :min="0"
          :step="widthType === 'pct' ? 5 : 100"
        />
        <span class="tpd-unit">{{
          widthType === "pct"
            ? t("dialogs.tableProperties.units.fiftiethsPercent")
            : t("dialogs.tableProperties.units.twips")
        }}</span>
      </div>
      <div class="tpd-row">
        <label class="tpd-label">{{ t("dialogs.tableProperties.alignmentLabel") }}</label>
        <select v-model="justification" class="tpd-select">
          <option value="left">{{ t("dialogs.tableProperties.alignOptions.left") }}</option>
          <option value="center">{{ t("dialogs.tableProperties.alignOptions.center") }}</option>
          <option value="right">{{ t("dialogs.tableProperties.alignOptions.right") }}</option>
        </select>
      </div>
    </div>
    <div class="tpd-dialog__footer">
      <button class="tpd-btn" @click="$emit('close')">{{ t("common.cancel") }}</button>
      <button class="tpd-btn tpd-btn--primary" @click="apply">{{ t("common.apply") }}</button>
    </div>
  </FolioDialog>
</template>

<script setup lang="ts">
import { ref, watch } from "vue";
import { useTranslation } from "../../i18n";
import { useFolioUI } from "../../ui/folio-ui";

const { t } = useTranslation();

// Resolve Dialog from the FolioUI injection provider so a host override
// takes effect here too (previously a hand-rolled overlay + dialog div pair).
const { Dialog: FolioDialog } = useFolioUI();

type TableJustification = "left" | "center" | "right";

export type TableProperties = {
  width?: number | null;
  widthType?: string | null;
  justification?: TableJustification | null;
};

const props = defineProps<{
  isOpen: boolean;
  currentProps?: { width?: number; widthType?: string; justification?: string };
}>();

const emit = defineEmits<{
  (e: "close"): void;
  (e: "apply", props: TableProperties): void;
}>();

const width = ref(0);
const widthType = ref("auto");
const justification = ref<TableJustification>("left");

function toJustification(value: string | undefined): TableJustification {
  return value === "center" || value === "right" ? value : "left";
}

watch(
  () => props.isOpen,
  (open) => {
    if (open) {
      width.value = props.currentProps?.width ?? 0;
      widthType.value = props.currentProps?.widthType ?? "auto";
      justification.value = toJustification(props.currentProps?.justification);
    }
  },
);

function apply() {
  const result: TableProperties =
    widthType.value === "auto"
      ? { justification: justification.value, width: null, widthType: "auto" }
      : { justification: justification.value, width: width.value, widthType: widthType.value };
  emit("apply", result);
  emit("close");
}
</script>

<style scoped>
/* The overlay + dialog box (background, border, shadow, centering) now come
   from the injected Dialog primitive (see the `<FolioDialog>` usage above);
   its width constraints are passed as an inline `:style` since Vue's scoped
   CSS can't reach into a child component's own template nodes. */
.tpd-dialog__header {
  padding: 16px 20px 12px;
  border-bottom: 1px solid var(--doc-border);
  font-size: 16px;
  font-weight: 600;
  color: var(--doc-text);
}
.tpd-dialog__body {
  padding: 16px 20px;
  display: flex;
  flex-direction: column;
  gap: 12px;
}
.tpd-dialog__footer {
  padding: 12px 20px 16px;
  border-top: 1px solid var(--doc-border);
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
.tpd-row {
  display: flex;
  align-items: center;
  gap: 12px;
}
.tpd-label {
  width: 80px;
  font-size: 13px;
  color: var(--doc-text-muted);
}
.tpd-input,
.tpd-select {
  flex: 1;
  padding: 6px 8px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  font-size: 13px;
}
.tpd-unit {
  font-size: 11px;
  color: var(--doc-text-subtle);
}
.tpd-btn {
  padding: 6px 16px;
  font-size: 13px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  cursor: pointer;
  background: var(--doc-surface);
}
.tpd-btn--primary {
  background: var(--doc-primary);
  color: var(--doc-on-primary);
  border-color: var(--doc-primary);
}
</style>
