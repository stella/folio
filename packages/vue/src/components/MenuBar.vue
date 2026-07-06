<!--
  File / Format / Insert / Help menus. Items emit a string `action` event the
  host (DocxEditor.vue) routes to the matching insert/format side-effect; Insert
  > Table opens an inline grid picker (TableGridInline) and emits `insert-table`
  with the chosen dimensions.

  Insert > Table is hidden when `showTableInsert` is false, mirroring how React's
  toolbar gates its Insert Table control (FormattingBar `showTableInsert`).

  TODO(i18n): the `toolbar.*` keys used below are not in folio's flat catalog
  (`@stll/folio-core/i18n/messages/en.json`); `t()` falls back to the key path
  until they are added. Upstream keys kept verbatim — do not invent entries.
-->
<template>
  <div class="menu-bar" role="menubar">
    <MenuDropdown :label="t('toolbar.file')" :items="fileItems" />
    <MenuDropdown :label="t('toolbar.format')" :items="formatItems" />
    <MenuDropdown :label="t('toolbar.insert')" :items="insertItems">
      <template #submenu="{ item, closeMenu }">
        <TableGridInline
          v-if="item.key === 'table'"
          @insert="
            (rows: number, cols: number) => {
              emit('insert-table', rows, cols);
              closeMenu();
            }
          "
        />
        <div v-else-if="item.key === 'break'" class="menu-bar__break-submenu">
          <button
            v-for="b in breakItems"
            :key="b.action"
            type="button"
            class="menu-bar__break-item"
            @click.prevent="
              () => {
                emit('action', b.action);
                closeMenu();
              }
            "
          >
            <MaterialSymbol :name="b.icon" :size="18" />
            <span>{{ b.label }}</span>
          </button>
        </div>
      </template>
    </MenuDropdown>
    <MenuDropdown v-if="showHelpMenu" :label="t('toolbar.help')" :items="helpItems" />
  </div>
</template>

<script setup lang="ts">
import { computed } from 'vue';
import { useTranslation } from '../i18n';
import MaterialSymbol from './ui/MaterialSymbol.vue';
import type { MenuEntry } from './ui/MenuDropdown.vue';
import TableGridInline from './ui/TableGridInline.vue';
import { useFolioUI } from '../ui/folio-ui';

// Resolve Menu from the FolioUI injection provider so a host override takes
// effect here too (previously a static import of MenuDropdown.vue).
const { Menu: MenuDropdown } = useFolioUI();

const props = withDefaults(
  defineProps<{ showFileOpen?: boolean; showHelpMenu?: boolean; showTableInsert?: boolean }>(),
  {
    showFileOpen: true,
    showHelpMenu: true,
    showTableInsert: true,
  },
);

const emit = defineEmits<{
  (e: 'action', action: string): void;
  (e: 'insert-table', rows: number, cols: number): void;
}>();

const { t } = useTranslation();

function act(action: string) {
  return () => emit('action', action);
}

const fileItems = computed<MenuEntry[]>(() => {
  const items: MenuEntry[] = [];
  if (props.showFileOpen) {
    items.push({
      icon: 'file_upload',
      label: t('toolbar.open'),
      shortcut: t('toolbar.openShortcut'),
      onClick: act('open'),
    });
  }
  items.push(
    {
      icon: 'file_download',
      label: t('toolbar.save'),
      shortcut: t('toolbar.saveShortcut'),
      onClick: act('save'),
    },
    { type: 'separator' },
    {
      label: t('toolbar.findReplace'),
      shortcut: t('toolbar.findReplaceShortcut'),
      onClick: act('findReplace'),
    },
    { type: 'separator' },
    { icon: 'settings', label: t('toolbar.pageSetup'), onClick: act('pageSetup') }
  );
  return items;
});

const formatItems = computed<MenuEntry[]>(() => [
  { icon: 'format_textdirection_l_to_r', label: t('toolbar.leftToRight'), onClick: act('dirLTR') },
  { icon: 'format_textdirection_r_to_l', label: t('toolbar.rightToLeft'), onClick: act('dirRTL') },
]);

const insertItems = computed<MenuEntry[]>(() => {
  const items: MenuEntry[] = [{ icon: 'image', label: t('toolbar.image'), onClick: act('insertImage') }];
  // Gate the Insert > Table item on `showTableInsert`, mirroring React's
  // FormattingBar Insert-Table control gating.
  if (props.showTableInsert) {
    items.push({ icon: 'grid_on', label: t('toolbar.table'), key: 'table', submenu: true });
  }
  items.push(
    { type: 'separator' },
    { icon: 'page_break', label: t('toolbar.break'), key: 'break', submenu: true },
    { icon: 'format_list_numbered', label: t('toolbar.tableOfContents'), onClick: act('insertTOC') },
    { icon: 'branding_watermark', label: t('toolbar.watermark'), onClick: act('watermark') },
  );
  return items;
});

const breakItems = computed(() => [
  { icon: 'page_break', label: t('toolbar.pageBreak'), action: 'insertPageBreak' },
  {
    icon: 'horizontal_rule',
    label: t('toolbar.sectionBreakNextPage'),
    action: 'insertSectionBreakNextPage',
  },
  {
    icon: 'border_horizontal',
    label: t('toolbar.sectionBreakContinuous'),
    action: 'insertSectionBreakContinuous',
  },
]);

const helpItems = computed<MenuEntry[]>(() => [
  { label: t('toolbar.reportIssue'), onClick: act('reportIssue') },
]);
</script>

<style scoped>
.menu-bar {
  display: flex;
  align-items: center;
}
.menu-bar__break-submenu {
  display: flex;
  flex-direction: column;
  min-width: 220px;
}
.menu-bar__break-item {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 12px;
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 13px;
  color: var(--doc-text);
  width: 100%;
  text-align: left;
  white-space: nowrap;
}
.menu-bar__break-item:hover {
  background: var(--doc-bg-hover);
}
</style>
