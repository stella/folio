<template>
  <div class="hf-inline-editor" :style="containerStyle" role="presentation">
    <div
      class="hf-separator-bar"
      :class="{ 'hf-separator-bar--footer': edit.position === 'footer' }"
    >
      <span class="hf-editor-label">{{ label }}</span>
      <div class="hf-options" @mousedown.stop>
        <FolioButton
          variant="ghost"
          size="xs"
          class-name="hf-options-button"
          @click.stop="showOptions = !showOptions"
        >
          {{ t("headerFooter.options") }} ▾
        </FolioButton>
        <div
          v-if="showOptions"
          class="hf-options-dropdown"
          :class="{ 'hf-options-dropdown--up': edit.position === 'footer' }"
        >
          <FolioButton variant="ghost" size="xs" @click="insertField('PAGE')">
            {{ t("headerFooter.insertPageNumber") }}
          </FolioButton>
          <FolioButton variant="ghost" size="xs" @click="insertField('NUMPAGES')">
            {{ t("headerFooter.insertPageCount") }}
          </FolioButton>
          <div class="hf-options-divider" />
          <FolioButton variant="ghost" size="xs" @click="emit('remove')">
            {{
              edit.position === "header"
                ? t("headerFooter.removeHeader")
                : t("headerFooter.removeFooter")
            }}
          </FolioButton>
          <FolioButton variant="ghost" size="xs" @click="emit('close')">
            {{
              edit.position === "header"
                ? t("headerFooter.closeHeaderEditing")
                : t("headerFooter.closeFooterEditing")
            }}
          </FolioButton>
        </div>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, type CSSProperties } from "vue";
import type { EditorView } from "prosemirror-view";

import { schema } from "@stll/folio-core/prosemirror";

import type { HfEditState } from "../composables/usePagesPointer";
import { useTranslation } from "../i18n";
import { useFolioUI } from "../ui/folio-ui";

const props = defineProps<{
  edit: HfEditState;
  getView: () => EditorView | null;
}>();

const emit = defineEmits<{
  close: [];
  remove: [];
}>();

const { Button: FolioButton } = useFolioUI();
const { t } = useTranslation();
const showOptions = ref(false);

const label = computed(() =>
  props.edit.position === "header" ? t("headerFooter.header") : t("headerFooter.footer"),
);

const containerStyle = computed<CSSProperties>(() => ({
  position: "absolute",
  top: `${props.edit.targetRect?.top ?? 0}px`,
  left: `${props.edit.targetRect?.left ?? 0}px`,
  width: `${props.edit.targetRect?.width ?? 0}px`,
  height: `${props.edit.targetRect?.height ?? 0}px`,
  pointerEvents: "none",
  zIndex: 10,
}));

const insertField = (fieldType: "PAGE" | "NUMPAGES"): void => {
  showOptions.value = false;
  const view = props.getView();
  const field = schema.nodes["field"];
  if (!view || !field) return;
  const { $from, from } = view.state.selection;
  const marks = view.state.storedMarks ?? $from.marks();
  const node = field.create({
    fieldType,
    instruction: ` ${fieldType} \\* MERGEFORMAT `,
    fieldKind: "simple",
    dirty: true,
  });
  view.dispatch(view.state.tr.insert(from, node.mark(marks)));
  view.focus();
};

const handleEscape = (event: KeyboardEvent): void => {
  if (event.key !== "Escape") return;
  event.preventDefault();
  event.stopPropagation();
  emit("close");
};

let focusFrame: number | null = null;
onMounted(() => {
  document.addEventListener("keydown", handleEscape, true);
  let attempts = 0;
  const focusView = (): void => {
    const view = props.getView();
    if (view && !view.hasFocus()) view.focus();
    attempts++;
    if (attempts < 30) focusFrame = requestAnimationFrame(focusView);
  };
  focusFrame = requestAnimationFrame(focusView);
});

onBeforeUnmount(() => {
  document.removeEventListener("keydown", handleEscape, true);
  if (focusFrame !== null) cancelAnimationFrame(focusFrame);
});
</script>

<style scoped>
.hf-separator-bar {
  position: absolute;
  right: 0;
  bottom: 100%;
  left: 0;
  display: flex;
  align-items: center;
  justify-content: space-between;
  color: var(--doc-link);
  font-size: 11px;
  pointer-events: auto;
}

.hf-separator-bar--footer {
  top: 100%;
  bottom: auto;
}

.hf-editor-label {
  font-weight: 500;
  letter-spacing: 0.3px;
}

.hf-options {
  position: relative;
}

.hf-options-dropdown {
  position: absolute;
  top: 100%;
  right: 0;
  display: flex;
  min-width: 220px;
  flex-direction: column;
  align-items: stretch;
  padding: 4px;
  border: 1px solid var(--border);
  border-radius: 6px;
  background: var(--background);
  box-shadow: 0 4px 12px rgb(0 0 0 / 12%);
  pointer-events: auto;
}

.hf-options-dropdown--up {
  top: auto;
  bottom: 100%;
}

.hf-options-dropdown :deep(button) {
  justify-content: flex-start;
}

.hf-options-divider {
  height: 1px;
  margin: 4px;
  background: var(--border);
}
</style>
