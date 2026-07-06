<!--
  Vue port of packages/react/src/components/FormattingBar.tsx —
  a clean, minimal, controlled formatting bar for legal document editing.

  Controls (left to right):
  Undo Redo | Style ▾ | B I U | A▾ | ≡ 1. ◁ ▷

  Standalone + controlled, mirroring React: it owns no editor view. The host
  passes the selection's current formatting via `currentFormatting` and wires
  the emitted `format` / `undo` / `redo` events back into the document (React
  threads the same intent through its `onFormat` / `onUndo` / `onRedo` props).
  Everything else (font, size, highlight, comments) is reachable via keyboard
  shortcuts or the host app's chrome, matching React's minimal subset.

  Composes the same Vue `ui/` pickers the full <Toolbar> uses (StylePicker,
  ColorPicker, AlignmentButtons, ListButtons) so the two rails stay visually
  and behaviourally aligned.
-->
<template>
  <div
    class="docx-formatting-bar"
    :class="className"
    role="toolbar"
    :aria-label="t('formattingToolbar')"
    data-folio-toolbar="true"
  >
    <!-- Undo / Redo -->
    <span class="docx-formatting-bar__group" role="group" :aria-label="t('historyGroup')">
      <button
        type="button"
        class="docx-formatting-bar__btn"
        :disabled="disabled || !canUndo"
        :title="t('undoShortcut')"
        :aria-label="t('undo')"
        @mousedown.prevent="emit('undo')"
      >
        <MaterialSymbol name="undo" :size="18" />
      </button>
      <button
        type="button"
        class="docx-formatting-bar__btn"
        :disabled="disabled || !canRedo"
        :title="t('redoShortcut')"
        :aria-label="t('redo')"
        @mousedown.prevent="emit('redo')"
      >
        <MaterialSymbol name="redo" :size="18" />
      </button>
    </span>

    <span class="docx-formatting-bar__divider" role="separator" />

    <!-- Paragraph style -->
    <template v-if="showStylePicker">
      <!-- `styles` is passed only when supplied so StylePicker keeps its own
           default-preset fallback (its prop is non-`undefined` under
           exactOptionalPropertyTypes). -->
      <StylePicker
        :value="currentFormatting.styleId ?? 'Normal'"
        v-bind="stylePickerStyles"
        :disabled="disabled"
        @change="(styleId: string) => emit('format', { type: 'applyStyle', value: styleId })"
      />
      <span class="docx-formatting-bar__divider" role="separator" />
    </template>

    <!-- Bold / Italic / Underline -->
    <span class="docx-formatting-bar__group" role="group" :aria-label="t('textFormattingGroup')">
      <button
        type="button"
        class="docx-formatting-bar__btn"
        :class="{ 'docx-formatting-bar__btn--active': currentFormatting.bold }"
        :disabled="disabled"
        :title="t('boldShortcut')"
        :aria-label="t('bold')"
        :aria-pressed="!!currentFormatting.bold"
        @mousedown.prevent="emit('format', 'bold')"
      >
        <MaterialSymbol name="format_bold" :size="18" />
      </button>
      <button
        type="button"
        class="docx-formatting-bar__btn"
        :class="{ 'docx-formatting-bar__btn--active': currentFormatting.italic }"
        :disabled="disabled"
        :title="t('italicShortcut')"
        :aria-label="t('italic')"
        :aria-pressed="!!currentFormatting.italic"
        @mousedown.prevent="emit('format', 'italic')"
      >
        <MaterialSymbol name="format_italic" :size="18" />
      </button>
      <button
        type="button"
        class="docx-formatting-bar__btn"
        :class="{ 'docx-formatting-bar__btn--active': currentFormatting.underline }"
        :disabled="disabled"
        :title="t('underlineShortcut')"
        :aria-label="t('underline')"
        :aria-pressed="!!currentFormatting.underline"
        @mousedown.prevent="emit('format', 'underline')"
      >
        <MaterialSymbol name="format_underlined" :size="18" />
      </button>
    </span>

    <!-- Text color -->
    <template v-if="showTextColorPicker">
      <span class="docx-formatting-bar__divider" role="separator" />
      <ColorPicker
        mode="text"
        :value="textColorHex"
        :theme="theme ?? null"
        :disabled="disabled"
        @change="onTextColor"
      />
    </template>

    <span class="docx-formatting-bar__divider" role="separator" />

    <!-- Alignment -->
    <AlignmentButtons
      v-if="showAlignmentButtons"
      :value="currentFormatting.alignment ?? 'left'"
      :disabled="disabled"
      @change="
        (alignment: ParagraphAlignment) => emit('format', { type: 'alignment', value: alignment })
      "
    />

    <!-- Lists + indent -->
    <ListButtons
      v-if="showListButtons"
      :list-state="listState"
      :disabled="disabled"
      :can-outdent="canOutdent"
      @bullet-list="emit('format', 'bulletList')"
      @numbered-list="emit('format', 'numberedList')"
      @indent="emit('format', 'indent')"
      @outdent="emit('format', 'outdent')"
    />
  </div>
</template>

<script setup lang="ts">
import { computed } from "vue";
import MaterialSymbol from "./ui/MaterialSymbol.vue";
import StylePicker from "./ui/StylePicker.vue";
import AlignmentButtons from "./ui/AlignmentButtons.vue";
import ListButtons from "./ui/ListButtons.vue";
import { createDefaultListState } from "../utils/listState";
import { useTranslation } from "../i18n";
import { useFolioUI } from "../ui/folio-ui";

// Resolve ColorPicker from the FolioUI injection provider so a host
// `components.ColorPicker` override renders here; falls back to the package
// default (ui/ColorPicker.vue) when rendered outside a provider.
const { ColorPicker } = useFolioUI();
import type { ColorValue } from "@stll/folio-core/types/document";
import type { ListState } from "../utils/listState";
import type { ParagraphAlignment } from "./ui/AlignmentButtons.types";
import type { StyleOption } from "./ui/StylePicker.types";
import type { FormattingAction, FormattingBarProps, SelectionFormatting } from "./FormattingBar.types";

const props = withDefaults(defineProps<FormattingBarProps>(), {
  disabled: false,
  canUndo: false,
  canRedo: false,
  showStylePicker: true,
  showTextColorPicker: true,
  showAlignmentButtons: true,
  showListButtons: true,
});

const emit = defineEmits<{
  (e: "format", action: FormattingAction): void;
  (e: "undo"): void;
  (e: "redo"): void;
}>();

const { t } = useTranslation();

const currentFormatting = computed<SelectionFormatting>(() => props.currentFormatting ?? {});

// Omit `styles` entirely when the host passes none, so StylePicker falls back
// to its built-in presets instead of receiving `undefined`.
const stylePickerStyles = computed<{ styles: StyleOption[] } | Record<string, never>>(() =>
  props.documentStyles ? { styles: props.documentStyles } : {},
);

// ListButtons requires a concrete ListState (its prop is non-`undefined` under
// exactOptionalPropertyTypes); fall back to the empty default like React does.
const listState = computed<ListState>(
  () => currentFormatting.value.listState ?? createDefaultListState(),
);

// Hex (no `#`, uppercase) of the applied text color, matching the active
// swatch in the picker — mirrors Toolbar.vue's `currentTextColorHex`.
const textColorHex = computed<string | undefined>(() => {
  const color = currentFormatting.value.color;
  return color ? color.replace(/^#/, "").toUpperCase() : undefined;
});

// Outdent is only meaningful in a nested list level or when the paragraph
// already carries a left indent — mirrors Toolbar.vue's `canOutdent`.
const canOutdent = computed(() => {
  const current = listState.value;
  const inNestedList = current.isInList && current.level > 0;
  return inNestedList || (currentFormatting.value.indentLeft ?? 0) > 0;
});

function onTextColor(color: ColorValue | string) {
  emit("format", { type: "textColor", value: color });
}
</script>

<style scoped>
.docx-formatting-bar {
  display: flex;
  align-items: center;
  gap: 1px;
  width: 100%;
  height: 48px;
  padding: 0 16px;
  overflow-x: auto;
  border-bottom: 1px solid var(--doc-border);
  background: var(--doc-page);
  scrollbar-width: none;
}
.docx-formatting-bar::-webkit-scrollbar {
  display: none;
}
.docx-formatting-bar__group {
  display: inline-flex;
  align-items: center;
  gap: 2px;
}
.docx-formatting-bar__btn {
  display: inline-flex;
  align-items: center;
  justify-content: center;
  min-width: 32px;
  height: 32px;
  border: none;
  border-radius: 6px;
  background: transparent;
  color: var(--doc-text-muted);
  cursor: pointer;
  flex-shrink: 0;
  transition:
    background-color 0.1s ease,
    color 0.1s ease;
}
.docx-formatting-bar__btn:focus-visible {
  outline: 2px solid var(--doc-primary);
  outline-offset: 2px;
}
.docx-formatting-bar__btn:hover:not(:disabled) {
  background: var(--doc-primary-light);
  color: var(--doc-text);
}
.docx-formatting-bar__btn--active,
.docx-formatting-bar__btn--active:hover {
  background: var(--doc-primary-light);
  color: var(--doc-text);
}
.docx-formatting-bar__btn:disabled {
  cursor: not-allowed;
  opacity: 0.35;
}
.docx-formatting-bar__divider {
  width: 1px;
  height: 24px;
  margin: 0 8px;
  background: var(--doc-border);
  flex-shrink: 0;
}
</style>
