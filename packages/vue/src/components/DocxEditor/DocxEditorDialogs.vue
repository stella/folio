<!--
  Modal-dialog cluster for DocxEditor — collects the dialogs the editor surfaces
  (find/replace, hyperlink, insert symbol, image properties, page setup, table
  properties) so the host template does not carry the dialog markup just to
  wire show-flags and close handlers. (Image insertion has no dialog — Insert >
  Image opens the OS file picker directly.)

  Visibility is passed via `v-model:show-*` so the host owns the boolean refs and
  dialogs close themselves through the standard `update:` emit pattern. Action
  emits (`insert-symbol`, `hyperlink-submit`, `page-setup-apply`,
  `table-properties-apply`) bubble up so the host's feature composables keep
  ownership of the document mutations.

  Fork note: the Watermark and Keyboard-Shortcuts dialogs are PORT-BLOCKED (not
  ported to the fork's `dialogs/` set yet) and are omitted here.
-->
<template>
  <FindReplaceDialog
    :is-open="showFindReplace"
    :view="view"
    :scroll-visible-position-into-view="scrollVisiblePositionIntoView"
    @close="emit('update:showFindReplace', false)"
  />

  <HyperlinkDialog
    :is-open="showHyperlink"
    :view="view"
    :bookmarks="bookmarks"
    @close="emit('update:showHyperlink', false)"
    @submit="(data) => emit('hyperlink-submit', data)"
    @remove="emit('hyperlink-remove')"
  />

  <InsertSymbolDialog
    :is-open="showInsertSymbol"
    @close="emit('update:showInsertSymbol', false)"
    @insert="(symbol) => emit('insert-symbol', symbol)"
  />

  <ImagePropertiesDialog
    :is-open="showImageProperties"
    :view="view"
    :pm-pos="selectedImagePmPos"
    @close="emit('update:showImageProperties', false)"
  />

  <PageSetupDialog
    :is-open="showPageSetup"
    :section-properties="sectionProperties"
    @close="emit('update:showPageSetup', false)"
    @apply="(props) => emit('page-setup-apply', props)"
  />

  <TablePropertiesDialog
    :is-open="showTableProperties"
    :current-props="tableProperties"
    @close="emit('update:showTableProperties', false)"
    @apply="(props) => emit('table-properties-apply', props)"
  />
</template>

<script setup lang="ts">
import type { EditorView } from "prosemirror-view";

import type { SectionProperties } from "@stll/folio-core/types/document";

import FindReplaceDialog from "../dialogs/FindReplaceDialog.vue";
import HyperlinkDialog from "../dialogs/HyperlinkDialog.vue";
import ImagePropertiesDialog from "../dialogs/ImagePropertiesDialog.vue";
import InsertSymbolDialog from "../dialogs/InsertSymbolDialog.vue";
import PageSetupDialog from "../dialogs/PageSetupDialog.vue";
import TablePropertiesDialog from "../dialogs/TablePropertiesDialog.vue";

type BookmarkOption = {
  name: string;
  label?: string;
};

type HyperlinkSubmitPayload = {
  url?: string;
  bookmark?: string;
  displayText: string;
  tooltip: string;
};

type TableProperties = {
  width?: number | null;
  widthType?: string | null;
  justification?: "left" | "center" | "right" | null;
};

type CurrentTableProperties = {
  width?: number;
  widthType?: string;
  justification?: string;
};

defineProps<{
  view: EditorView | null;
  bookmarks: BookmarkOption[];
  selectedImagePmPos: number | null;
  sectionProperties: SectionProperties | null;
  scrollVisiblePositionIntoView: (pmPos: number) => void;
  showFindReplace: boolean;
  showHyperlink: boolean;
  showInsertSymbol: boolean;
  showImageProperties: boolean;
  showPageSetup: boolean;
  showTableProperties: boolean;
  tableProperties: CurrentTableProperties;
}>();

const emit = defineEmits<{
  (e: "update:showFindReplace", value: boolean): void;
  (e: "update:showHyperlink", value: boolean): void;
  (e: "update:showInsertSymbol", value: boolean): void;
  (e: "update:showImageProperties", value: boolean): void;
  (e: "update:showPageSetup", value: boolean): void;
  (e: "update:showTableProperties", value: boolean): void;
  (e: "insert-symbol", symbol: string): void;
  (e: "hyperlink-submit", data: HyperlinkSubmitPayload): void;
  (e: "hyperlink-remove"): void;
  (e: "page-setup-apply", props: Partial<SectionProperties>): void;
  (e: "table-properties-apply", props: TableProperties): void;
}>();
</script>
