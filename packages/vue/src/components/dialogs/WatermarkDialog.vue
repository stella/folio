<template>
  <FolioDialog
    :open="isOpen"
    aria-label="Watermark"
    :style="{ minWidth: '360px', maxWidth: '520px', width: '100%' }"
    @close="$emit('close')"
  >
    <div class="dialog__header">Watermark</div>
    <div class="dialog__body">
      <label class="field">
        <span class="field__label">Type</span>
        <select v-model="mode" class="field__input">
          <option value="text">Text watermark</option>
          <option value="picture">Picture watermark</option>
          <option value="none">No watermark</option>
        </select>
      </label>

      <div v-if="mode === 'text'" class="grid">
        <label class="field field--wide">
          <span class="field__label">Text</span>
          <input v-model="text" class="field__input" />
        </label>
        <label class="field">
          <span class="field__label">Font</span>
          <input v-model="font" class="field__input" />
        </label>
        <label class="field">
          <span class="field__label">Color</span>
          <input v-model="color" class="field__input field__input--color" type="color" />
        </label>
        <label class="field">
          <span class="field__label">Opacity</span>
          <input
            v-model.number="opacityPercent"
            class="field__input"
            max="100"
            min="0"
            type="number"
          />
        </label>
        <label class="check">
          <input v-model="diagonal" type="checkbox" />
          <span>Diagonal</span>
        </label>
      </div>

      <div v-if="mode === 'picture'" class="grid">
        <label class="field field--wide">
          <span class="field__label">Header image relationship id</span>
          <input v-model="imageRId" class="field__input" placeholder="rId12" />
        </label>
        <label class="field field--wide">
          <span class="field__label">Image target</span>
          <input v-model="imageTarget" class="field__input" placeholder="word/media/image1.png" />
        </label>
        <label class="field">
          <span class="field__label">Scale</span>
          <input
            v-model.number="scalePercent"
            class="field__input"
            max="300"
            min="1"
            type="number"
          />
        </label>
        <label class="check">
          <input v-model="washout" type="checkbox" />
          <span>Washout</span>
        </label>
        <label class="check field--wide">
          <input v-model="imageTargetExternal" type="checkbox" />
          <span>Target is an external URL</span>
        </label>
      </div>
    </div>
    <div class="dialog__footer">
      <button class="dialog__btn" @click="$emit('close')">Cancel</button>
      <button class="dialog__btn dialog__btn--primary" :disabled="!canApply" @click="apply">
        Apply
      </button>
    </div>
  </FolioDialog>
</template>

<script setup lang="ts">
import { computed, ref, watch } from "vue";
import type { Watermark } from "@stll/folio-core/watermark";
import { useFolioUI } from "../../ui/folio-ui";

const { Dialog: FolioDialog } = useFolioUI();

type WatermarkMode = "none" | "text" | "picture";

const props = defineProps<{
  isOpen: boolean;
  currentWatermark?: Watermark;
}>();

const emit = defineEmits<{
  (e: "close"): void;
  (e: "apply", watermark: Watermark | undefined): void;
}>();

const DEFAULT_TEXT_COLOR = "#C0C0C0";

const mode = ref<WatermarkMode>("text");
const text = ref("CONFIDENTIAL");
const font = ref("Calibri");
const color = ref(DEFAULT_TEXT_COLOR);
const diagonal = ref(true);
const opacityPercent = ref(50);
const imageRId = ref("");
const imageTarget = ref("");
const imageTargetExternal = ref(false);
const scalePercent = ref(100);
const washout = ref(true);

const canApply = computed(
  () =>
    mode.value === "none" ||
    (mode.value === "text" && text.value.trim().length > 0) ||
    (mode.value === "picture" && imageRId.value.trim().length > 0),
);

watch(
  () => props.isOpen,
  (open) => {
    if (!open) return;
    const watermark = props.currentWatermark;
    if (!watermark) {
      mode.value = "text";
      text.value = "CONFIDENTIAL";
      font.value = "Calibri";
      color.value = DEFAULT_TEXT_COLOR;
      diagonal.value = true;
      opacityPercent.value = 50;
      imageRId.value = "";
      imageTarget.value = "";
      imageTargetExternal.value = false;
      scalePercent.value = 100;
      washout.value = true;
      return;
    }
    mode.value = watermark.kind;
    if (watermark.kind === "text") {
      text.value = watermark.text;
      font.value = watermark.font ?? "Calibri";
      color.value = toColorInputValue(watermark.color);
      diagonal.value = watermark.diagonal ?? true;
      opacityPercent.value = Math.round((watermark.opacity ?? 0.5) * 100);
      return;
    }
    imageRId.value = watermark.imageRId;
    imageTarget.value = watermark.imageTarget ?? "";
    imageTargetExternal.value = watermark.imageTargetExternal ?? false;
    scalePercent.value = Math.round((watermark.scale ?? 1) * 100);
    washout.value = watermark.washout ?? true;
  },
  { immediate: true },
);

function apply() {
  if (!canApply.value) return;
  if (mode.value === "none") {
    emit("apply", undefined);
    emit("close");
    return;
  }
  if (mode.value === "text") {
    emit("apply", {
      kind: "text",
      text: text.value.trim(),
      ...(font.value.trim() ? { font: font.value.trim() } : {}),
      color: stripHash(color.value),
      diagonal: diagonal.value,
      opacity: clampPercent(opacityPercent.value) / 100,
    });
    emit("close");
    return;
  }
  emit("apply", {
    kind: "picture",
    imageRId: imageRId.value.trim(),
    ...(imageTarget.value.trim() ? { imageTarget: imageTarget.value.trim() } : {}),
    ...(imageTarget.value.trim() ? { imageTargetExternal: imageTargetExternal.value } : {}),
    scale: clampPercent(scalePercent.value) / 100,
    washout: washout.value,
  });
  emit("close");
}

function clampPercent(value: number): number {
  return Math.min(100, Math.max(0, value));
}

function stripHash(value: string): string {
  return value.startsWith("#") ? value.slice(1) : value;
}

function toColorInputValue(value: string | undefined): string {
  if (!value || value === "auto") return DEFAULT_TEXT_COLOR;
  return value.startsWith("#") ? value : `#${value}`;
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
.field--wide {
  grid-column: 1 / -1;
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
.field__input--color {
  min-height: 34px;
  padding: 2px;
}
.check {
  display: flex;
  align-items: center;
  gap: 8px;
  padding-top: 22px;
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
