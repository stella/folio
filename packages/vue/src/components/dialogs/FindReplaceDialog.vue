<!--
  Find / Replace dialog — ported from the upstream docx-editor Vue adapter.

  Most labels remap to the folio `findReplace.*` catalog; the replace-specific
  ones resolve `dialogs.findReplace.*`.
-->
<template>
  <div
    v-if="isOpen"
    class="find-replace-dialog"
    data-testid="find-replace-dialog"
    @mousedown.stop
    @keydown.stop
  >
    <div class="find-replace-dialog__header">
      <span class="find-replace-dialog__title">{{
        replaceMode ? t("findReplace.findAndReplace") : t("findReplace.find")
      }}</span>
      <FolioButton
        class="find-replace-dialog__close"
        variant="ghost"
        size="icon-xs"
        @click="close"
        :title="t('findReplace.close')"
      >
        ✕
      </FolioButton>
    </div>

    <div class="find-replace-dialog__body">
      <!-- Search row -->
      <div class="find-replace-dialog__row">
        <FolioInput
          ref="searchInputRef"
          v-model="searchText"
          class="find-replace-dialog__input"
          data-testid="find-input"
          :placeholder="t('findReplace.findPlaceholder')"
          :aria-label="t('findReplace.findText')"
          @keydown="handleSearchKeyDown"
          @input="performSearch"
        />
        <span class="find-replace-dialog__count">{{ matchCountText }}</span>
        <FolioButton
          variant="ghost"
          size="icon-xs"
          :title="t('findReplace.previous')"
          @mousedown.prevent="findPrevious"
          >▲</FolioButton
        >
        <FolioButton
          variant="ghost"
          size="icon-xs"
          :title="t('findReplace.next')"
          @mousedown.prevent="findNext"
          >▼</FolioButton
        >
      </div>

      <!-- Options row -->
      <div class="find-replace-dialog__options">
        <label>
          <FolioCheckbox
            :checked="matchCase"
            @update:checked="
              (v: boolean) => {
                matchCase = v;
                performSearch();
              }
            "
          />
          {{ t("findReplace.matchCase") }}
        </label>
        <label>
          <FolioCheckbox
            :checked="matchWholeWord"
            @update:checked="
              (v: boolean) => {
                matchWholeWord = v;
                performSearch();
              }
            "
          />
          {{ t("findReplace.wholeWords") }}
        </label>
        <FolioButton
          class="find-replace-dialog__toggle"
          data-testid="find-replace-toggle"
          variant="ghost"
          size="xs"
          :class="{ active: replaceMode }"
          @mousedown.prevent="replaceMode = !replaceMode"
          :title="t('dialogs.findReplace.toggleReplace')"
        >
          ↔ {{ t("dialogs.findReplace.replaceButton") }}
        </FolioButton>
      </div>

      <!-- Replace row -->
      <div v-if="replaceMode" class="find-replace-dialog__row">
        <FolioInput
          v-model="replaceText"
          class="find-replace-dialog__input"
          data-testid="replace-input"
          :placeholder="t('findReplace.replacePlaceholder')"
          :aria-label="t('findReplace.replaceText')"
          @keydown.enter.prevent="handleReplace"
        />
        <FolioButton
          variant="ghost"
          size="xs"
          data-testid="replace-current"
          :title="t('findReplace.replaceCurrent')"
          @mousedown.prevent="handleReplace"
        >
          {{ t("dialogs.findReplace.replaceButton") }}
        </FolioButton>
        <FolioButton
          variant="ghost"
          size="xs"
          data-testid="replace-all"
          :title="t('findReplace.replaceAll')"
          @mousedown.prevent="handleReplaceAll"
        >
          {{ t("findReplace.replaceAll") }}
        </FolioButton>
      </div>
    </div>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, ref, toRef, watch } from "vue";
import type { EditorView } from "prosemirror-view";
import { useTranslation } from "../../i18n";
import { useFolioUI } from "../../ui/folio-ui";
import { useFindReplace } from "../../composables/useFindReplace";

const { t } = useTranslation();

// Resolve Button/Input/Checkbox from the FolioUI injection provider so a host
// override takes effect here too (previously static markup — this dialog is a
// floating panel, not a modal, so it does not use the Dialog primitive).
const { Button: FolioButton, Input: FolioInput, Checkbox: FolioCheckbox } = useFolioUI();

const props = defineProps<{
  isOpen: boolean;
  view: EditorView | null;
  scrollVisiblePositionIntoView?: (pmPos: number) => void;
}>();

const emit = defineEmits<{
  (e: "close"): void;
}>();

const searchInputRef = ref<HTMLInputElement | null>(null);
const replaceMode = ref(false);
const {
  searchText,
  replaceText,
  options,
  matches,
  currentIndex,
  performSearch,
  findNext,
  findPrevious,
  replaceCurrent: handleReplace,
  replaceAll: handleReplaceAll,
  clear: clearHighlights,
} = useFindReplace({
  editorView: toRef(props, "view"),
  scrollVisiblePositionIntoView: (pmPos) => props.scrollVisiblePositionIntoView?.(pmPos),
});
const matchCase = toRef(options, "matchCase");
const matchWholeWord = toRef(options, "matchWholeWord");

const matchCountText = computed(() => {
  if (!searchText.value.trim()) return "";
  if (matches.value.length === 0) return t("findReplace.noResults");
  if (currentIndex.value < 0)
    return t("dialogs.findReplace.matchesFound", { total: matches.value.length });
  return t("findReplace.matchCounter", {
    current: currentIndex.value + 1,
    total: matches.value.length,
  });
});

// Focus search input when dialog opens
watch(
  () => props.isOpen,
  async (open) => {
    if (open) {
      await nextTick();
      searchInputRef.value?.focus();
      searchInputRef.value?.select();
      performSearch();
    } else {
      clearHighlights();
    }
  },
);

function close() {
  clearHighlights();
  emit("close");
}

function handleSearchKeyDown(e: KeyboardEvent) {
  if (e.key === "Enter") {
    e.preventDefault();
    if (e.shiftKey) {
      findPrevious();
    } else {
      findNext();
    }
  } else if (e.key === "Escape") {
    close();
  }
}
</script>

<style scoped>
.find-replace-dialog {
  /* `position: fixed` + a top-of-stack z-index so the panel floats above
     toolbar dropdowns and the agent panel instead of behind them (it was
     `position: absolute; z-index: 100`, which loses to both). Matches the
     React Find/Replace dialog and the `contextMenu` tier in
     src/styles/zIndex.ts. */
  position: fixed;
  top: 8px;
  right: 16px;
  z-index: 10000;
  background: var(--doc-surface);
  border: 1px solid var(--doc-border-dark);
  border-radius: 8px;
  box-shadow: 0 4px 12px var(--doc-shadow);
  min-width: 360px;
  font-size: 13px;
}
.find-replace-dialog__header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 8px 12px;
  border-bottom: 1px solid var(--doc-border);
}
.find-replace-dialog__title {
  font-weight: 600;
  color: var(--doc-text);
}
.find-replace-dialog__close {
  border: none;
  background: transparent;
  cursor: pointer;
  color: var(--doc-text-muted);
  font-size: 14px;
  width: 24px;
  height: 24px;
  display: flex;
  align-items: center;
  justify-content: center;
  border-radius: 4px;
}
.find-replace-dialog__close:hover {
  background: var(--doc-bg-hover);
}
.find-replace-dialog__body {
  padding: 8px 12px 12px;
}
.find-replace-dialog__row {
  display: flex;
  align-items: center;
  gap: 4px;
  margin-bottom: 6px;
}
.find-replace-dialog__input {
  flex: 1;
  padding: 5px 8px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  font-size: 13px;
  outline: none;
}
.find-replace-dialog__input:focus {
  border-color: var(--doc-primary);
  box-shadow: 0 0 0 2px var(--doc-focus-ring);
}
.find-replace-dialog__count {
  font-size: 11px;
  color: var(--doc-text-muted);
  white-space: nowrap;
  min-width: 60px;
  text-align: center;
}
.find-replace-dialog__options {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 6px;
  font-size: 12px;
  color: var(--doc-text-muted);
}
.find-replace-dialog__options label {
  display: flex;
  align-items: center;
  gap: 4px;
  cursor: pointer;
}
.find-replace-dialog__options input[type="checkbox"] {
  margin: 0;
}
.find-replace-dialog__toggle {
  margin-left: auto;
  font-size: 11px;
  padding: 2px 8px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  background: transparent;
  cursor: pointer;
  color: var(--doc-text-muted);
}
.find-replace-dialog__toggle.active {
  background: var(--doc-accent-bg);
  border-color: var(--doc-accent);
  color: var(--doc-accent);
}
.find-replace-dialog__row button {
  padding: 4px 8px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  background: var(--doc-surface);
  cursor: pointer;
  font-size: 12px;
  color: var(--doc-text-muted);
}
.find-replace-dialog__row button:hover {
  background: var(--doc-bg-hover);
}
</style>
