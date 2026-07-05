<template>
  <button
    type="button"
    class="docx-outline-toggle docx-editor-vue__outline-toggle"
    :style="{ left: leftOffset + 'px', top: topPx + 'px' }"
    :title="t('editor.showDocumentOutline')"
    :aria-label="t('editor.showDocumentOutline')"
    @click="$emit('toggle')"
    @mousedown.stop
  >
    <MaterialSymbol name="format_list_bulleted" :size="20" />
  </button>
</template>

<script setup lang="ts">
// Outline toggle — Vue counterpart of React's OutlineToggleButton. Inject-based
// translation (host calls provideLocale) so the label tracks the user's i18n prop.
//
// TODO(i18n): `editor.showDocumentOutline` is not in folio's flat catalog
// (`@stll/folio-core/i18n/messages/en.json`); `t()` falls back to the key path
// until the entry is added. Upstream key kept verbatim — do not invent one.
import { useTranslation } from '../i18n';
import MaterialSymbol from './ui/MaterialSymbol.vue';

withDefaults(
  defineProps<{
    /**
     * Left anchor (px) from the editor area's left edge. The host bumps it past
     * the vertical ruler when one is shown so the button doesn't overlap it.
     */
    leftOffset: number;
    /**
     * Top anchor (px) from the editor area's top edge. The host bumps it past
     * the sticky ruler row when one is shown so the button doesn't overlap it.
     */
    topPx?: number;
  }>(),
  { topPx: 24 }
);

defineEmits<{ toggle: [] }>();

const { t } = useTranslation();
</script>
