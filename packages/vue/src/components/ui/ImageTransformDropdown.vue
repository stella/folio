<!--
  Vue port of packages/react/src/components/ui/ImageTransformDropdown.tsx —
  rotate CW / CCW / flip H / flip V via IconGridDropdown.

  TODO(i18n): every `imageTransform.*` key below is absent from folio's flat
  `folio` catalog (`packages/core/src/i18n/messages/en.json`); kept verbatim
  (not invented) so `t(...)` currently returns the raw key path until entries
  land.
-->
<template>
  <IconGridDropdown
    :options="options"
    trigger-icon="rotate_right"
    :tooltip-content="t('imageTransform.tooltip')"
    :disabled="disabled"
    @select="(v: TransformAction) => $emit('transform', v)"
  />
</template>

<script setup lang="ts">
import { computed } from 'vue';
import IconGridDropdown, { type IconGridOption } from './IconGridDropdown.vue';
import { useTranslation } from '../../i18n';

export type TransformAction = 'rotateCW' | 'rotateCCW' | 'flipH' | 'flipV';

defineProps<{
  disabled?: boolean;
}>();

defineEmits<{
  (e: 'transform', action: TransformAction): void;
}>();

const { t } = useTranslation();

const OPTION_DEFS: { value: TransformAction; labelKey: string; iconName: string }[] = [
  { value: 'rotateCW', labelKey: 'imageTransform.rotateClockwise', iconName: 'rotate_right' },
  {
    value: 'rotateCCW',
    labelKey: 'imageTransform.rotateCounterClockwise',
    iconName: 'rotate_left',
  },
  { value: 'flipH', labelKey: 'imageTransform.flipHorizontal', iconName: 'swap_horiz' },
  { value: 'flipV', labelKey: 'imageTransform.flipVertical', iconName: 'swap_vert' },
];

const options = computed<IconGridOption<TransformAction>[]>(() =>
  OPTION_DEFS.map((o) => ({ value: o.value, label: t(o.labelKey), iconName: o.iconName }))
);
</script>
