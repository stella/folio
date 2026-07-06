<!--
  Title-bar block for DocxEditor — DocumentName + MenuBar centred between the
  left/right chrome slots. Renders nothing when `showMenuBar` is false. The host
  SFC wraps this and the Toolbar inside the `__toolbar-shell` (shared background),
  mirroring React's `<TitleBar>` inside `<EditorToolbar>` arrangement.

  Fork note: the fork's `DocxEditorProps` does not (yet) expose a document-name /
  logo / title-bar-right surface, so this component keeps the slot-driven layout
  and lets the host pass a name only when it has one. `renderLogo` /
  `renderTitleBarRight` render props are PORT-BLOCKED (React-style Component props
  are not part of the fork's prop contract); use the named slots instead.
-->
<template>
  <div v-if="showMenuBar" class="docx-editor-vue__title-bar">
    <div class="docx-editor-vue__title-bar-left">
      <slot name="title-bar-left" />
    </div>
    <div class="docx-editor-vue__title-bar-center">
      <DocumentName
        :model-value="documentName"
        :editable="documentNameEditable"
        @update:model-value="(name: string) => emit('rename', name)"
      />
      <MenuBar
        :show-file-open="showFileOpen"
        :show-help-menu="showHelpMenu"
        :show-table-insert="showTableInsert"
        @action="(action: string) => emit('menu-action', action)"
        @insert-table="(rows: number, cols: number) => emit('insert-table', rows, cols)"
      />
    </div>
    <div class="docx-editor-vue__title-bar-right">
      <slot name="title-bar-right" />
    </div>
  </div>
</template>

<script setup lang="ts">
import DocumentName from "../DocumentName.vue";
import MenuBar from "../MenuBar.vue";

withDefaults(
  defineProps<{
    showMenuBar?: boolean;
    documentName?: string;
    documentNameEditable?: boolean;
    showFileOpen?: boolean;
    showHelpMenu?: boolean;
    /** Gates the Insert > Table menu item (mirrors the toolbar's showTableInsert). */
    showTableInsert?: boolean;
  }>(),
  {
    showMenuBar: true,
    documentName: "",
    documentNameEditable: true,
    showFileOpen: true,
    showHelpMenu: true,
    showTableInsert: true,
  },
);

const emit = defineEmits<{
  (e: "rename", name: string): void;
  (e: "menu-action", action: string): void;
  (e: "insert-table", rows: number, cols: number): void;
}>();
</script>
