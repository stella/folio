<!-- Thin Vue renderer for the shared content-control widget controller. -->
<template>
  <Teleport to="body">
    <div
      v-if="snapshot.status !== 'closed'"
      :class="['content-control-picker', portalClass]"
      :style="pickerStyle"
      role="dialog"
      :aria-label="
        snapshot.status === 'dropdown'
          ? t('contentControlDropdownAriaLabel')
          : t('contentControlDateAriaLabel')
      "
      @mousedown.stop
    >
      <div v-if="snapshot.status === 'dropdown'" class="content-control-picker__items" role="menu">
        <div v-if="snapshot.items.length === 0" class="content-control-picker__empty">
          {{ t("contentControlDropdownNoOptions") }}
        </div>
        <template v-else>
          <FolioButton
            v-for="item in snapshot.items"
            :key="`${item.value}::${item.displayText}`"
            class-name="content-control-picker__item"
            type="button"
            role="menuitem"
            variant="ghost"
            size="sm"
            @mousedown.prevent
            @click="controller.pickDropdown(item.value)"
          >
            {{ item.displayText }}
          </FolioButton>
        </template>
      </div>
      <DatePickerPopover
        v-else-if="snapshot.status === 'date'"
        :value="snapshot.currentValue"
        :default-open="true"
        @change="onDateChange"
      />
    </div>
  </Teleport>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, shallowRef, watch, type CSSProperties } from "vue";
import type { EditorView } from "prosemirror-view";

import { ContentControlWidgetController } from "@stll/folio-core/controller/contentControlWidgetController";
import { useDocxPortalClass } from "../composables/usePortalClass";
import { useTranslation } from "../i18n";
import { useFolioUI } from "../ui/folio-ui";

const props = defineProps<{ view: EditorView | null }>();

const { t } = useTranslation();
const { Button: FolioButton, DatePickerPopover } = useFolioUI();
const portalClass = useDocxPortalClass();
const controller = new ContentControlWidgetController();
const snapshot = shallowRef(controller.getSnapshot());
const unsubscribe = controller.subscribe(() => {
  snapshot.value = controller.getSnapshot();
});

watch(
  () => props.view,
  (view) => controller.bind(view),
  { immediate: true },
);

onBeforeUnmount(() => {
  unsubscribe();
  controller.destroy();
});

const pickerStyle = computed<CSSProperties>(() => {
  if (snapshot.value.status === "closed") {
    return {};
  }
  return {
    insetInlineStart: `${snapshot.value.position.x}px`,
    top: `${snapshot.value.position.y}px`,
  };
});

const onDateChange = (value: string | null): void => {
  if (value) {
    controller.pickDate(value);
  }
};
</script>

<style scoped>
.content-control-picker {
  position: fixed;
  z-index: 9999;
  min-width: 10rem;
  padding: 4px;
  color: var(--doc-text);
  background: var(--doc-surface);
  border: 1px solid var(--doc-border);
  border-radius: 6px;
  box-shadow: var(--doc-shadow-md, 0 4px 12px rgb(0 0 0 / 15%));
}

.content-control-picker__items {
  display: flex;
  flex-direction: column;
}

.content-control-picker__item {
  width: 100%;
  padding: 6px 8px;
  color: inherit;
  font: inherit;
  text-align: start;
  cursor: pointer;
  background: transparent;
  border: 0;
  border-radius: 4px;
}

.content-control-picker__item:hover,
.content-control-picker__item:focus-visible {
  background: var(--doc-bg-hover);
  outline: none;
}

.content-control-picker__empty {
  padding: 6px 8px;
  color: var(--doc-text-muted);
  font-size: 13px;
}
</style>
