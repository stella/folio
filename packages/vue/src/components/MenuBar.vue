<!--
  File / Format / Insert / Help menus — mirrors the MenuBar in
  packages/react/src/components/TitleBar.tsx. Items emit a string `action`
  event; Insert > Table opens an inline grid picker and emits `insert-table`.

  PORT-BLOCKED: `./ui/MenuDropdown.vue` and `./ui/TableGridInline.vue` are not
  yet ported. They are stubbed inline in <script setup> (no new files, so a
  parallel port of the real primitives can't collide) so this component compiles
  and keeps its full menu-item wiring; the stubs render nothing until the real
  ui primitives land. TODO: import the ported `./ui/MenuDropdown.vue` (+ its
  `MenuEntry` type) and `./ui/TableGridInline.vue` and delete the stubs.

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
import { computed, defineComponent, type PropType, type SlotsType } from 'vue';
import { useTranslation } from '../i18n';
import MaterialSymbol from './ui/MaterialSymbol.vue';

interface MenuItem {
  icon?: string;
  label: string;
  shortcut?: string;
  onClick?: () => void;
  disabled?: boolean;
  /** Stable id so the `#submenu` slot can tell which item it's rendering for. */
  key?: string;
  /** When true, the item shows a right-chevron and opens the `#submenu` slot on hover. */
  submenu?: boolean;
}
interface MenuSeparator {
  type: 'separator';
}
type MenuEntry = MenuItem | MenuSeparator;

// PORT-BLOCKED stubs — see the file header. No-op components so MenuBar
// type-checks and mounts; they paint nothing until ./ui/MenuDropdown.vue and
// ./ui/TableGridInline.vue are ported.
const MenuDropdown = defineComponent({
  name: 'MenuDropdownStub',
  props: {
    label: { type: String, required: true },
    items: { type: Array as PropType<MenuEntry[]>, required: true },
  },
  slots: Object as SlotsType<{
    submenu?: (slotProps: { item: MenuItem; closeMenu: () => void }) => unknown;
  }>,
  setup() {
    return () => null;
  },
});
const TableGridInline = defineComponent({
  name: 'TableGridInlineStub',
  emits: {
    insert: (_rows: number, _cols: number) => true,
  },
  setup() {
    return () => null;
  },
});

const props = withDefaults(defineProps<{ showFileOpen?: boolean; showHelpMenu?: boolean }>(), {
  showFileOpen: true,
  showHelpMenu: true,
});

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
    { icon: 'settings', label: t('toolbar.pageSetup'), onClick: act('pageSetup') }
  );
  return items;
});

const formatItems = computed<MenuEntry[]>(() => [
  { icon: 'format_textdirection_l_to_r', label: t('toolbar.leftToRight'), onClick: act('dirLTR') },
  { icon: 'format_textdirection_r_to_l', label: t('toolbar.rightToLeft'), onClick: act('dirRTL') },
]);

const insertItems = computed<MenuEntry[]>(() => [
  { icon: 'image', label: t('toolbar.image'), onClick: act('insertImage') },
  { icon: 'grid_on', label: t('toolbar.table'), key: 'table', submenu: true },
  { type: 'separator' },
  { icon: 'page_break', label: t('toolbar.break'), key: 'break', submenu: true },
  { icon: 'format_list_numbered', label: t('toolbar.tableOfContents'), onClick: act('insertTOC') },
  { icon: 'branding_watermark', label: t('toolbar.watermark'), onClick: act('watermark') },
]);

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
