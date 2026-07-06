<!--
  Footnote / Endnote Properties dialog — ported from the upstream docx-editor Vue
  adapter. Upstream cast every emitted field with `as any`; here the reactive refs
  are typed with the model unions (FootnotePosition, EndnotePosition,
  NoteNumberRestart, NumberFormat) so the emit is structurally sound without casts.

  TODO(i18n): all `dialogs.footnoteProperties.*` and `common.*` keys are kept
  verbatim; the folio catalog has no equivalents yet.
-->
<template>
  <FolioDialog
    :open="isOpen"
    :aria-label="t('dialogs.footnoteProperties.title')"
    :style="{ width: '440px', maxWidth: '90vw' }"
    @close="$emit('close')"
  >
      <div class="fnpd-dialog__header">{{ t('dialogs.footnoteProperties.title') }}</div>

      <div class="fnpd-dialog__body">
        <!-- Footnotes -->
        <fieldset class="fnpd-fieldset">
          <legend class="fnpd-legend">{{ t('dialogs.footnoteProperties.footnotes') }}</legend>
          <div class="fnpd-row">
            <label class="fnpd-label">{{ t('dialogs.footnoteProperties.position') }}</label>
            <select v-model="fnPosition" class="fnpd-select">
              <option value="pageBottom">
                {{ t('dialogs.footnoteProperties.footnotePositions.bottomOfPage') }}
              </option>
              <option value="beneathText">
                {{ t('dialogs.footnoteProperties.footnotePositions.belowText') }}
              </option>
            </select>
          </div>
          <div class="fnpd-row">
            <label class="fnpd-label">{{ t('dialogs.footnoteProperties.numberFormat') }}</label>
            <select v-model="fnNumFmt" class="fnpd-select">
              <option v-for="fmt in numberFormats" :key="fmt.value" :value="fmt.value">
                {{ fmt.label }}
              </option>
            </select>
          </div>
          <div class="fnpd-row">
            <label class="fnpd-label">{{ t('dialogs.footnoteProperties.startAt') }}</label>
            <input v-model.number="fnStart" type="number" class="fnpd-input" min="1" />
          </div>
          <div class="fnpd-row">
            <label class="fnpd-label">{{ t('dialogs.footnoteProperties.numbering') }}</label>
            <select v-model="fnRestart" class="fnpd-select">
              <option value="continuous">
                {{ t('dialogs.footnoteProperties.numberingOptions.continuous') }}
              </option>
              <option value="eachSect">
                {{ t('dialogs.footnoteProperties.numberingOptions.restartSection') }}
              </option>
              <option value="eachPage">
                {{ t('dialogs.footnoteProperties.numberingOptions.restartPage') }}
              </option>
            </select>
          </div>
        </fieldset>

        <!-- Endnotes -->
        <fieldset class="fnpd-fieldset">
          <legend class="fnpd-legend">{{ t('dialogs.footnoteProperties.endnotes') }}</legend>
          <div class="fnpd-row">
            <label class="fnpd-label">{{ t('dialogs.footnoteProperties.position') }}</label>
            <select v-model="enPosition" class="fnpd-select">
              <option value="docEnd">
                {{ t('dialogs.footnoteProperties.endnotePositions.endOfDocument') }}
              </option>
              <option value="sectEnd">
                {{ t('dialogs.footnoteProperties.endnotePositions.endOfSection') }}
              </option>
            </select>
          </div>
          <div class="fnpd-row">
            <label class="fnpd-label">{{ t('dialogs.footnoteProperties.numberFormat') }}</label>
            <select v-model="enNumFmt" class="fnpd-select">
              <option v-for="fmt in numberFormats" :key="fmt.value" :value="fmt.value">
                {{ fmt.label }}
              </option>
            </select>
          </div>
          <div class="fnpd-row">
            <label class="fnpd-label">{{ t('dialogs.footnoteProperties.startAt') }}</label>
            <input v-model.number="enStart" type="number" class="fnpd-input" min="1" />
          </div>
          <div class="fnpd-row">
            <label class="fnpd-label">{{ t('dialogs.footnoteProperties.numbering') }}</label>
            <select v-model="enRestart" class="fnpd-select">
              <option value="continuous">
                {{ t('dialogs.footnoteProperties.numberingOptions.continuous') }}
              </option>
              <option value="eachSect">
                {{ t('dialogs.footnoteProperties.numberingOptions.restartSection') }}
              </option>
            </select>
          </div>
        </fieldset>
      </div>

      <div class="fnpd-dialog__footer">
        <button class="fnpd-btn" @click="$emit('close')">{{ t('common.cancel') }}</button>
        <button class="fnpd-btn fnpd-btn--primary" @click="apply">{{ t('common.apply') }}</button>
      </div>
  </FolioDialog>
</template>

<script setup lang="ts">
import { computed, ref, watch } from 'vue';
import type {
  EndnotePosition,
  EndnoteProperties,
  FootnotePosition,
  FootnoteProperties,
  NoteNumberRestart,
} from '@stll/folio-core/types/content';
import type { NumberFormat } from '@stll/folio-core/types/document';
import { useTranslation } from '../../i18n';
import { useFolioUI } from '../../ui/folio-ui';

const { t } = useTranslation();

// Resolve Dialog from the FolioUI injection provider so a host override
// takes effect here too (previously a hand-rolled overlay + dialog div pair).
const { Dialog: FolioDialog } = useFolioUI();

const props = defineProps<{
  isOpen: boolean;
  footnotePr?: FootnoteProperties;
  endnotePr?: EndnoteProperties;
}>();

const emit = defineEmits<{
  (e: 'close'): void;
  (e: 'apply', footnotePr: FootnoteProperties, endnotePr: EndnoteProperties): void;
}>();

const numberFormats = computed<{ value: NumberFormat; label: string }[]>(() => [
  { value: 'decimal', label: t('dialogs.footnoteProperties.formats.decimal') },
  { value: 'lowerLetter', label: t('dialogs.footnoteProperties.formats.lowerAlpha') },
  { value: 'upperLetter', label: t('dialogs.footnoteProperties.formats.upperAlpha') },
  { value: 'lowerRoman', label: t('dialogs.footnoteProperties.formats.lowerRoman') },
  { value: 'upperRoman', label: t('dialogs.footnoteProperties.formats.upperRoman') },
  { value: 'chicago', label: t('dialogs.footnoteProperties.formats.symbols') },
]);

const fnPosition = ref<FootnotePosition>('pageBottom');
const fnNumFmt = ref<NumberFormat>('decimal');
const fnStart = ref(1);
const fnRestart = ref<NoteNumberRestart>('continuous');
const enPosition = ref<EndnotePosition>('docEnd');
const enNumFmt = ref<NumberFormat>('lowerRoman');
const enStart = ref(1);
const enRestart = ref<NoteNumberRestart>('continuous');

watch(
  () => props.isOpen,
  (open) => {
    if (!open) return;
    fnPosition.value = props.footnotePr?.position ?? 'pageBottom';
    fnNumFmt.value = props.footnotePr?.numFmt ?? 'decimal';
    fnStart.value = props.footnotePr?.numStart ?? 1;
    fnRestart.value = props.footnotePr?.numRestart ?? 'continuous';
    enPosition.value = props.endnotePr?.position ?? 'docEnd';
    enNumFmt.value = props.endnotePr?.numFmt ?? 'lowerRoman';
    enStart.value = props.endnotePr?.numStart ?? 1;
    enRestart.value = props.endnotePr?.numRestart ?? 'continuous';
  }
);

function apply() {
  emit(
    'apply',
    {
      position: fnPosition.value,
      numFmt: fnNumFmt.value,
      numStart: fnStart.value,
      numRestart: fnRestart.value,
    },
    {
      position: enPosition.value,
      numFmt: enNumFmt.value,
      numStart: enStart.value,
      numRestart: enRestart.value,
    }
  );
  emit('close');
}
</script>

<style scoped>
/* The overlay + dialog box (background, border, shadow, centering) now come
   from the injected Dialog primitive (see the `<FolioDialog>` usage above);
   its width constraints are passed as an inline `:style` since Vue's scoped
   CSS can't reach into a child component's own template nodes. */
.fnpd-dialog__header {
  padding: 16px 20px 12px;
  border-bottom: 1px solid var(--doc-border);
  font-size: 16px;
  font-weight: 600;
  color: var(--doc-text);
}
.fnpd-dialog__body {
  padding: 12px 20px;
  display: flex;
  flex-direction: column;
  gap: 12px;
  max-height: 60vh;
  overflow-y: auto;
}
.fnpd-dialog__footer {
  padding: 12px 20px 16px;
  border-top: 1px solid var(--doc-border);
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
.fnpd-fieldset {
  border: 1px solid var(--doc-border);
  border-radius: 6px;
  padding: 10px 12px;
  margin: 0;
}
.fnpd-legend {
  font-size: 12px;
  font-weight: 600;
  color: var(--doc-text-muted);
  padding: 0 4px;
}
.fnpd-row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 6px;
}
.fnpd-row:last-child {
  margin-bottom: 0;
}
.fnpd-label {
  width: 100px;
  font-size: 13px;
  color: var(--doc-text-muted);
  flex-shrink: 0;
}
.fnpd-input,
.fnpd-select {
  flex: 1;
  padding: 5px 8px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  font-size: 13px;
}
.fnpd-btn {
  padding: 6px 16px;
  font-size: 13px;
  border: 1px solid var(--doc-border-dark);
  border-radius: 4px;
  cursor: pointer;
  background: var(--doc-surface);
}
.fnpd-btn--primary {
  background: var(--doc-primary);
  color: var(--doc-on-primary);
  border-color: var(--doc-primary);
}
</style>
