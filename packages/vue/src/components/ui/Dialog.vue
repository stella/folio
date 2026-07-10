<!--
  Built-in, dependency-light Dialog used when a consumer does not inject one.
  Vue port of packages/react/src/ui/defaults/dialog.tsx, collapsed into one
  monolithic component (Vue's contract has no Root/Portal/Backdrop/Popup/Title/
  Close part-object, see ui/folio-ui.ts's module docblock): teleports to
  `document.body`, shows a click-to-close backdrop, and closes on Escape.
  Folio's chrome positions/sizes the popup itself via `className` (each dialog
  keeps its own scoped stylesheet for the box), so this default carries no
  layout opinion beyond centering.
-->
<template>
  <Teleport to="body">
    <div
      v-if="open"
      class="folio-default-dialog-backdrop"
      :class="backdropClass"
      @mousedown.self="onBackdropMousedown"
    >
      <div
        class="folio-default-dialog-popup"
        :class="className"
        role="dialog"
        :aria-label="ariaLabel"
        v-bind="$attrs"
        @keydown.escape="requestClose"
        @mousedown.stop
      >
        <slot />
      </div>
    </div>
  </Teleport>
</template>

<script setup lang="ts">
// Listeners a host attaches to `<FolioDialog>` (e.g. `@keydown.enter="apply"`)
// fall through via `$attrs`; `inheritAttrs: false` + the explicit `v-bind`
// keeps them off the (conditionally rendered) Teleport wrapper and only on
// the popup element, merging with this component's own `@keydown.escape`.
defineOptions({ inheritAttrs: false });

const props = withDefaults(
  defineProps<{
    open: boolean;
    ariaLabel?: string;
    className?: string;
    backdropClass?: string;
    closeOnBackdrop?: boolean;
  }>(),
  { closeOnBackdrop: true },
);

const emit = defineEmits<{
  (e: "update:open", value: boolean): void;
  (e: "close"): void;
}>();

function requestClose() {
  emit("update:open", false);
  emit("close");
}

function onBackdropMousedown() {
  if (props.closeOnBackdrop) requestClose();
}
</script>

<style scoped>
.folio-default-dialog-backdrop {
  position: fixed;
  inset: 0;
  z-index: 10000;
  display: flex;
  align-items: center;
  justify-content: center;
  background: var(--doc-overlay, rgb(0 0 0 / 50%));
}
.folio-default-dialog-popup {
  background: var(--doc-surface, #fff);
  border-radius: 8px;
  box-shadow: 0 4px 20px var(--doc-shadow, rgb(0 0 0 / 20%));
}
</style>
