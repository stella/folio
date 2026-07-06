<!--
  Mirror of React's review controls (DocxEditor.tsx `toolbarPriorityExtra`,
  gated on `showReviewControls`). Two pieces:

    1. Track-changes toggle — a pill button flipping the editor between
       `editing` and `suggesting`. Active (suggesting) state gets the primary
       accent; disabled while read-only. Label mirrors React's
       `trackingOn` / `trackingOff` catalog strings.
    2. Markup display-mode selector — an eye-icon dropdown choosing how tracked
       changes render (All Markup / Simple / No Markup / Original). Emits the
       chosen `DisplayMode`; the host applies the `folio-root--<mode>` class,
       exactly as React does.

  The toggle only reflects/sets track-changes; it never touches the third
  `viewing` mode (that stays reachable through the separate EditingModeDropdown).
-->
<template>
  <div class="review-controls">
    <button
      type="button"
      class="review-controls__toggle"
      :class="{ 'review-controls__toggle--on': trackChangesOn }"
      :disabled="readOnly"
      :aria-pressed="trackChangesOn"
      :aria-label="t('toggleTrackChanges')"
      :title="t('toggleTrackChanges')"
      @mousedown.prevent
      @click.prevent="$emit('toggle-track-changes')"
    >
      <MaterialSymbol name="edit_note" :size="18" />
      <span class="review-controls__toggle-label">{{
        trackChangesOn ? t('trackingOn') : t('trackingOff')
      }}</span>
    </button>

    <Popover
      :open="isOpen"
      placement="bottom-right"
      @update:open="(v) => (isOpen = v)"
      @close="isOpen = false"
    >
      <template #trigger="{ toggle }">
        <button
          type="button"
          class="review-controls__display"
          :class="{ 'review-controls__display--open': isOpen }"
          :title="currentLabel"
          @mousedown.prevent
          @click.prevent="toggle"
        >
          <MaterialSymbol name="visibility" :size="16" />
          <span class="review-controls__display-label">{{ currentLabel }}</span>
          <MaterialSymbol name="arrow_drop_down" :size="16" />
        </button>
      </template>
      <template #panel>
        <div class="review-controls__panel" @mousedown.prevent>
          <button
            v-for="mode in DISPLAY_MODES"
            :key="mode"
            type="button"
            class="review-controls__option"
            @click.prevent="select(mode)"
          >
            <span class="review-controls__option-label">{{ labelFor(mode) }}</span>
            <MaterialSymbol
              v-if="mode === displayMode"
              name="check"
              :size="18"
              class="review-controls__check"
            />
          </button>
        </div>
      </template>
    </Popover>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from 'vue';
import { DISPLAY_MODES } from '@stll/folio-core/managers/EditorModeManager';
import type { DisplayMode } from '@stll/folio-core/managers/EditorModeManager';
import MaterialSymbol from './ui/MaterialSymbol.vue';
import Popover from './ui/Popover.vue';
import { useTranslation } from '../i18n';

const { t } = useTranslation();

const props = withDefaults(
  defineProps<{
    trackChangesOn: boolean;
    displayMode: DisplayMode;
    readOnly?: boolean;
  }>(),
  { readOnly: false }
);

const emit = defineEmits<{
  (e: 'toggle-track-changes'): void;
  (e: 'update:display-mode', mode: DisplayMode): void;
}>();

const isOpen = ref(false);

// Catalog keys mirror React's `DISPLAY_MODE_LABELS` (markupView.* namespace).
const DISPLAY_MODE_LABEL_KEYS: Record<DisplayMode, string> = {
  'all-markup': 'markupView.allMarkup',
  'simple-markup': 'markupView.simple',
  'no-markup': 'markupView.noMarkup',
  original: 'markupView.original',
};

function labelFor(mode: DisplayMode): string {
  return t(DISPLAY_MODE_LABEL_KEYS[mode]);
}

const currentLabel = computed(() => labelFor(props.displayMode));

function select(mode: DisplayMode): void {
  emit('update:display-mode', mode);
  isOpen.value = false;
}
</script>

<style scoped>
.review-controls {
  display: flex;
  align-items: center;
  gap: 2px;
  flex-shrink: 0;
}

/* --- Track-changes toggle ------------------------------------------------ */
.review-controls__toggle {
  display: flex;
  align-items: center;
  gap: 4px;
  height: 28px;
  padding: 2px 8px 2px 6px;
  border: 1px solid transparent;
  border-radius: 4px;
  background: transparent;
  cursor: pointer;
  font-size: 13px;
  font-weight: 400;
  color: var(--doc-text-muted);
  white-space: nowrap;
}
.review-controls__toggle:hover:not(:disabled) {
  border-color: var(--doc-border);
  background: var(--doc-bg-hover);
  color: var(--doc-text);
}
.review-controls__toggle--on {
  border-color: var(--doc-primary);
  background: var(--doc-bg-hover);
  color: var(--doc-text);
  box-shadow: 0 0 0 1px var(--doc-primary);
}
.review-controls__toggle:disabled {
  opacity: 0.35;
  cursor: default;
  color: var(--doc-text-subtle);
}
.review-controls__toggle-label {
  padding: 0 2px;
}

/* --- Display-mode selector ----------------------------------------------- */
.review-controls__display {
  display: flex;
  align-items: center;
  gap: 4px;
  height: 28px;
  padding: 2px 4px 2px 6px;
  border: none;
  border-radius: 4px;
  background: transparent;
  cursor: pointer;
  font-size: 13px;
  font-weight: 400;
  color: var(--doc-text-muted);
  white-space: nowrap;
}
.review-controls__display:hover,
.review-controls__display--open {
  background: var(--doc-bg-hover);
  color: var(--doc-text);
}
.review-controls__display-label {
  padding: 0 2px;
}
.review-controls__panel {
  padding: 4px 0;
  min-width: 180px;
}
.review-controls__option {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 12px;
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 13px;
  color: var(--doc-text);
  width: 100%;
  text-align: left;
}
.review-controls__option:hover {
  background: var(--doc-bg-hover);
}
.review-controls__option-label {
  flex: 1 1 auto;
}
.review-controls__check {
  margin-left: auto;
  color: var(--doc-primary);
}
</style>
