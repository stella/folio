<!--
  Floating-popover cluster for DocxEditor — collects the click-anchored popups
  that surface above the editor: the selection / table context menu and the image
  context menu. Mounted at the SFC root (after the editor scroll region) so their
  absolute positioning is not constrained by the pages-viewport's stacking
  context.

  Visibility is owned by the host's context-menu state; this component just routes
  events back so the host can dispatch into them.
-->
<template>
  <TextContextMenu
    :is-open="contextMenu.isOpen"
    :position="contextMenu.position"
    :has-selection="contextMenu.hasSelection"
    :is-editable="!readOnly"
    :in-table="contextMenu.inTable"
    :on-image="contextMenu.onImage"
    :can-merge-cells="contextMenu.canMergeCells"
    :can-split-cell="contextMenu.canSplitCell"
    :custom-items="customTextMenuItems"
    @action="(action: string) => emit('context-menu-action', action)"
    @close="emit('close-context-menu')"
  />

  <ImageContextMenu
    :state="imageContextMenu"
    :text-actions="imageContextMenuTextActions"
    :can-open-properties="canOpenImageProperties"
    @close="emit('close-image-context-menu')"
    @select="(target: ImageLayoutTarget) => emit('image-wrap-select', target)"
    @text-action="(action: string) => emit('context-menu-action', action)"
    @open-properties="emit('open-image-properties')"
  />
</template>

<script setup lang="ts">
import type { ImageLayoutTarget } from "@stll/folio-core/layout-painter/imageLayout";

import ImageContextMenu from "../ImageContextMenu.vue";
import type { ImageContextMenuState, ImageContextMenuTextAction } from "../imageContextMenuTypes";
import TextContextMenu from "../TextContextMenu.vue";
import type { CustomTextMenuItem } from "../TextContextMenu.vue";

/** Text/table context-menu state (mirrors useContextMenus `TextContextMenuState`). */
type TextContextMenuState = {
  isOpen: boolean;
  position: { x: number; y: number };
  hasSelection: boolean;
  inTable: boolean;
  onImage: boolean;
  canMergeCells: boolean;
  canSplitCell: boolean;
};

defineProps<{
  readOnly: boolean;
  contextMenu: TextContextMenuState;
  imageContextMenu: ImageContextMenuState | null;
  imageContextMenuTextActions: ImageContextMenuTextAction[];
  canOpenImageProperties: boolean;
  /** Host-injected menu entries leading the text menu. */
  customTextMenuItems: readonly CustomTextMenuItem[];
}>();

const emit = defineEmits<{
  (e: "context-menu-action", action: string): void;
  (e: "close-context-menu"): void;
  (e: "image-wrap-select", target: ImageLayoutTarget): void;
  (e: "close-image-context-menu"): void;
  (e: "open-image-properties"): void;
}>();
</script>
