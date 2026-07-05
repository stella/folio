<!--
  DocxEditor.vue — top-level orchestration shell for the fork's Vue adapter.

  Composes the load-bearing pipeline (`useDocxEditor`: DOCX/Document → PM state →
  layout → paint) with the editor chrome and exposes the imperative ref described
  by `DocxEditorRef` via `useDocxEditorRefApi`.

  Responsibilities (mirrors React's `DocxEditor`):
   1. Owns the two template refs the pipeline paints against: `hiddenPmRef` (the
      off-screen ProseMirror host) and `pagesRef` (the painted-pages
      `HTMLDivElement`).
   2. Drives loading from prop watchers on `documentBuffer` / `document` (via
      `useDocumentLifecycle`), calling `loadBuffer` / `loadDocument`.
   3. Subscribes to the reactive `isReady` / `parseError` / `layout` surface for
      chrome and to `editor.on("selectionChange" | "docChange" | "layoutComplete")`
      to bump `stateTick` so the toolbar/overlays recompute.
   4. Passes `readOnly` / `mode` / `author` / `documentKey` and the callback props
      through to the pipeline.

  PORT-BLOCKED (fork has not ported the backing composable/component yet):
   - externalPlugins / i18n-locale props: not on the fork's `DocxEditorProps`, so
     `externalPlugins` is `[]` and the locale defaults to `en`.
   - colorMode prop + useColorMode: dark mode is fixed light here.
   - Comment lifecycle/management, image actions, context menus, header/footer
     editing, decoration overlay, hyperlink management, rulers: the composables
     are absent, so the sidebar/outline/context-menu chrome renders as empty
     shells and hyperlink/menu-table actions are no-ops.
-->
<template>
  <div
    :class="[
      'docx-editor-vue ep-root paged-editor',
      className,
      { 'paged-editor--readonly': readOnly },
    ]"
    :style="style"
  >
    <div class="docx-editor-vue__toolbar-shell">
      <DocxEditorMenuBar
        @rename="(name: string) => emit('rename', name)"
        @menu-action="handleMenuAction"
        @insert-table="handleMenuTableInsert"
      >
        <template #title-bar-left><slot name="title-bar-left" /></template>
        <template #title-bar-right><slot name="title-bar-right" /></template>
      </DocxEditorMenuBar>

      <Toolbar
        v-if="showToolbar"
        :view="editorView"
        :get-commands="getCommands"
        :state-tick="stateTick"
        :zoom-percent="zoomPercent"
        :is-min-zoom="isMinZoom"
        :is-max-zoom="isMaxZoom"
        :zoom-presets="ZOOM_PRESETS"
        :show-zoom-control="showZoomControl"
        :editor-mode="editorMode"
        :comments-sidebar-open="showSidebar"
        :image-context="null"
        :theme="theme ?? null"
        v-bind="toolbarDynamicProps"
        @find-replace="showFindReplace = true"
        @insert-link="showHyperlink = true"
        @insert-symbol="showInsertSymbol = true"
        @insert-page-break="handleInsertPageBreak"
        @page-setup="showPageSetup = true"
        @toggle-outline="showOutline = !showOutline"
        @apply-style="handleApplyStyle"
        @zoom-in="zoomIn"
        @zoom-out="zoomOut"
        @zoom-set="setZoom"
        @toggle-sidebar="showSidebar = !showSidebar"
        @mode-change="setEditorMode"
        @image-properties="showImageProperties = true"
      >
        <template v-if="$slots['toolbar-extra']" #toolbar-extra>
          <slot name="toolbar-extra" />
        </template>
      </Toolbar>
    </div>

    <DocxEditorDialogs
      v-model:show-find-replace="showFindReplace"
      v-model:show-hyperlink="showHyperlink"
      v-model:show-insert-symbol="showInsertSymbol"
      v-model:show-image-properties="showImageProperties"
      v-model:show-page-setup="showPageSetup"
      :view="editorView"
      :bookmarks="bookmarks"
      :selected-image-pm-pos="null"
      :section-properties="currentSectionProps"
      @insert-symbol="handleInsertSymbol"
      @hyperlink-submit="() => {}"
      @hyperlink-remove="() => {}"
      @page-setup-apply="handlePageSetupApply"
    />

    <div v-if="parseError" class="docx-editor-vue__error">{{ parseError }}</div>
    <div v-if="!isReady && !parseError" class="docx-editor-vue__loading">Loading…</div>

    <div ref="hiddenPmRef" class="docx-editor-vue__hidden-pm paged-editor__hidden-pm" />

    <div class="docx-editor-vue__editor-scroll">
      <div class="docx-editor-vue__editor-area">
        <div ref="pagesViewportRef" class="docx-editor-vue__pages-viewport" @wheel="handleZoomWheel">
          <div
            class="docx-editor-vue__editor-content-wrapper"
            :style="{ position: 'relative', display: 'flex', flexDirection: 'column', minHeight: '100%' }"
          >
            <div ref="pagesRef" class="docx-editor-vue__pages paged-editor__pages" :style="pagesContainerStyle" />
          </div>

          <CommentMarginMarkers
            :comments="comments"
            :pages-container="pagesRef"
            :zoom="zoom"
            :page-width-px="pageWidthPx"
            :sidebar-open="showSidebar"
            :resolved-comment-ids="resolvedCommentIds"
            @marker-click="() => {}"
          />

          <UnifiedSidebar
            :is-open="showSidebar"
            :comments="comments"
            :tracked-changes="trackedChanges"
            :show-resolved="true"
            :pages-container="pagesRef"
            :page-width-px="pageWidthPx"
            :zoom="zoom"
            :active-item-id="activeSidebarItem"
            @close="showSidebar = false"
            @update:active-item-id="(id: string | null) => (activeSidebarItem = id)"
          />
        </div>

        <OutlineToggleButton v-if="!showOutline" :left-offset="12" @toggle="showOutline = true" />

        <PageIndicator
          v-if="scrollPageInfo.totalPages > 1"
          :current-page="scrollPageInfo.currentPage"
          :total-pages="scrollPageInfo.totalPages"
          :visible="scrollPageInfo.visible"
        />

        <DocumentOutline
          :is-open="showOutline"
          :headings="outlineHeadings"
          @close="showOutline = false"
          @navigate="() => {}"
        />
      </div>
    </div>

    <DocxEditorOverlays
      :read-only="readOnly"
      :context-menu="contextMenu"
      :image-context-menu="imageContextMenu"
      :image-context-menu-text-actions="imageContextMenuTextActions"
      :can-open-image-properties="false"
      @context-menu-action="() => {}"
      @close-context-menu="contextMenu.isOpen = false"
      @image-wrap-select="() => {}"
      @close-image-context-menu="imageContextMenu = null"
      @open-image-properties="showImageProperties = true"
    />
  </div>
</template>

<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref, shallowRef, watch } from "vue";

import { extractSelectionState } from "@stll/folio-core/prosemirror";
import { twipsToPixels } from "@stll/folio-core/paged-layout/sectionGeometry";
import type { Comment } from "@stll/folio-core/types/content";
import type { Document, SectionProperties, Style } from "@stll/folio-core/types/document";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";

import CommentMarginMarkers from "./CommentMarginMarkers.vue";
import DocumentOutline from "./DocumentOutline.vue";
import DocxEditorDialogs from "./DocxEditor/DocxEditorDialogs.vue";
import DocxEditorMenuBar from "./DocxEditor/DocxEditorMenuBar.vue";
import DocxEditorOverlays from "./DocxEditor/DocxEditorOverlays.vue";
import type { DocxEditorProps, EditorMode } from "./DocxEditor/types";
import type { ImageContextMenuState, ImageContextMenuTextAction } from "./imageContextMenuTypes";
import OutlineToggleButton from "./OutlineToggleButton.vue";
import PageIndicator from "./PageIndicator.vue";
import type { TrackedChangeEntry } from "./sidebar/sidebarUtils";
import Toolbar from "./Toolbar.vue";
import UnifiedSidebar from "./UnifiedSidebar.vue";

import { useDocumentLifecycle } from "../composables/useDocumentLifecycle";
import { useDocxEditor } from "../composables/useDocxEditor";
import { useDocxEditorRefApi } from "../composables/useDocxEditorRefApi";
import { useFormattingActions } from "../composables/useFormattingActions";
import { usePageSetupControls } from "../composables/usePageSetupControls";
import { provideDocxPortalClass } from "../composables/usePortalClass";
import { useZoom } from "../composables/useZoom";
import type { FontOption } from "../utils/fontOptions";
import { provideLocale } from "../i18n";

const props = withDefaults(defineProps<DocxEditorProps>(), {
  documentBuffer: null,
  document: null,
  showToolbar: true,
  showZoomControl: true,
  readOnly: false,
  author: "User",
  mode: "editing",
  initialZoom: 1,
  showOutline: false,
  className: "",
});

const emit = defineEmits<{
  (e: "change", doc: Document): void;
  (e: "update:document", doc: Document | null): void;
  (e: "error", error: Error): void;
  (e: "ready"): void;
  (e: "rename", name: string): void;
  (e: "menu-action", action: string): void;
  (e: "mode-change", mode: EditorMode): void;
}>();

// PORT-BLOCKED: no i18n/locale prop on the fork's DocxEditorProps yet — default to `en`.
provideLocale();
// PORT-BLOCKED: no colorMode prop; share a fixed-light token scope with teleported chrome.
const isDark = ref(false);
provideDocxPortalClass(isDark);

const editorMode = ref<EditorMode>(props.mode);
const readOnly = computed(() => props.readOnly || editorMode.value === "viewing");

// ---- Template refs (paint targets) --------------------------------------
const hiddenPmRef = ref<HTMLElement | null>(null);
const pagesRef = ref<HTMLElement | null>(null);
const pagesViewportRef = ref<HTMLElement | null>(null);

// ---- Chrome state -------------------------------------------------------
const stateTick = ref(0);
const showFindReplace = ref(false);
const showHyperlink = ref(false);
const showInsertSymbol = ref(false);
const showImageProperties = ref(false);
const showPageSetup = ref(false);
const showOutline = ref(props.showOutline);
const showSidebar = ref(false);
const activeSidebarItem = ref<string | null>(null);
const bookmarks = shallowRef<{ name: string; label?: string }[]>([]);

// Comment / tracked-change / outline payloads stay empty until the comment and
// outline composables are ported (PORT-BLOCKED).
const comments = shallowRef<Comment[]>([]);
const trackedChanges = shallowRef<TrackedChangeEntry[]>([]);
const outlineHeadings = shallowRef<HeadingInfo[]>([]);

// PORT-BLOCKED: useContextMenus is absent — drive the overlay cluster with a
// static closed state so it renders (nothing) and typechecks.
const contextMenu = ref({
  isOpen: false,
  position: { x: 0, y: 0 },
  hasSelection: false,
  inTable: false,
  onImage: false,
  canMergeCells: false,
  canSplitCell: false,
});
const imageContextMenu = ref<ImageContextMenuState | null>(null);
const imageContextMenuTextActions: ImageContextMenuTextAction[] = [];

const { zoom, zoomPercent, isMinZoom, isMaxZoom, setZoom, zoomIn, zoomOut, handleWheel: handleZoomWheel, ZOOM_PRESETS } =
  useZoom(props.initialZoom);

// ---- Pipeline -----------------------------------------------------------
const {
  editor,
  editorView,
  isReady,
  parseError,
  layout,
  loadBuffer,
  loadDocument,
  save,
  getDocument,
  getCommands,
  focus,
  reLayout,
} = useDocxEditor({
  hiddenContainer: hiddenPmRef,
  pagesContainer: pagesRef,
  readOnly,
  editorMode,
  author: () => props.author,
  documentKey: () => props.documentKey,
  // PORT-BLOCKED: no externalPlugins prop on the fork's DocxEditorProps yet.
  externalPlugins: [],
  onChange: (doc) => {
    props.onChange?.(doc);
    emit("change", doc);
    emit("update:document", doc);
  },
  onError: (err) => {
    props.onError?.(err);
    emit("error", err);
  },
  onSelectionUpdate: (state) => {
    props.onSelectionChange?.(extractSelectionState(state));
  },
  onEditorViewReady: (view) => props.onEditorViewReady?.(view),
  onReadOnlyEditAttempt: () => props.onReadonlyEditAttempt?.(),
});

// ---- Document-derived computed refs -------------------------------------
const currentSectionProps = computed<SectionProperties | null>(() => {
  void stateTick.value;
  const body = getDocument()?.package.document;
  if (!body) {
    return null;
  }
  return body.finalSectionProperties ?? body.sections?.[0]?.properties ?? null;
});

// Optional Toolbar props that must be OMITTED (not passed as `undefined`) under
// exactOptionalPropertyTypes: bind via `v-bind` so an absent value drops the key.
const toolbarDynamicProps = computed(() => {
  void stateTick.value;
  const dynamic: {
    fontFamilies?: ReadonlyArray<string | FontOption>;
    documentStyles?: Style[];
  } = {};
  if (props.fontFamilies !== undefined) {
    dynamic.fontFamilies = props.fontFamilies;
  }
  const styles = getDocument()?.package.styles?.styles;
  if (styles !== undefined) {
    dynamic.documentStyles = styles;
  }
  return dynamic;
});

const pageWidthPx = computed(() => twipsToPixels(currentSectionProps.value?.pageWidth ?? 12240) * zoom.value);

const resolvedCommentIds = computed(() => new Set<number>());

const scrollPageInfo = computed(() => ({
  currentPage: 1,
  totalPages: layout.value?.pages.length ?? 0,
  visible: false,
}));

const pagesContainerStyle = computed(() => ({
  transform: zoom.value === 1 ? undefined : `scale(${zoom.value})`,
  transformOrigin: "top center",
  transition: "transform 0.2s ease",
}));

// ---- Feature handlers ----------------------------------------------------
const { handleApplyStyle, handleInsertSymbol, handleInsertPageBreak } = useFormattingActions({
  editorView,
  getDocument,
});

// PORT-BLOCKED: usePageSetupControls emits host notifications we do not re-route yet.
const { handlePageSetupApply } = usePageSetupControls({
  editorView,
  getDocument,
  readOnly,
  stateTick,
  reLayout,
  emit: () => {},
});

function setEditorMode(mode: EditorMode): void {
  if (editorMode.value === mode) {
    return;
  }
  editorMode.value = mode;
  props.onModeChange?.(mode);
  emit("mode-change", mode);
}

function handleMenuAction(action: string): void {
  switch (action) {
    case "find":
    case "findReplace":
      showFindReplace.value = true;
      break;
    case "insertSymbol":
      showInsertSymbol.value = true;
      break;
    case "insertHyperlink":
      showHyperlink.value = true;
      break;
    case "pageSetup":
      showPageSetup.value = true;
      break;
    case "toggleOutline":
      showOutline.value = !showOutline.value;
      break;
    case "toggleSidebar":
      showSidebar.value = !showSidebar.value;
      break;
    default:
      break;
  }
  emit("menu-action", action);
}

// PORT-BLOCKED: menu-driven table insertion needs a command wrapper, not ported.
function handleMenuTableInsert(_rows: number, _cols: number): void {}

// ---- Loading + lifecycle -------------------------------------------------
const sidebarAutoOpenedRef = ref(false);
useDocumentLifecycle({
  documentBuffer: () => props.documentBuffer ?? null,
  document: () => props.document ?? null,
  loadDocumentBuffer: loadBuffer,
  loadDocument,
  sidebarAutoOpenedRef,
});

watch(
  () => props.mode,
  (mode) => {
    if (mode && mode !== editorMode.value) {
      editorMode.value = mode;
    }
  },
);

watch(
  () => props.showOutline,
  (next) => {
    showOutline.value = !!next;
  },
);

watch(isReady, (ready) => {
  if (ready) {
    emit("ready");
  }
});

// Recompute chrome state on every selection / doc / layout event (Requirement 3).
onMounted(() => {
  const offSelection = editor.on("selectionChange", () => {
    stateTick.value++;
  });
  const offDoc = editor.on("docChange", () => {
    stateTick.value++;
  });
  const offLayout = editor.on("layoutComplete", () => {
    stateTick.value++;
  });
  onBeforeUnmount(() => {
    offSelection();
    offDoc();
    offLayout();
  });
});

// ---- Imperative ref surface ---------------------------------------------
const { exposed } = useDocxEditorRefApi({
  editor,
  editorView,
  layout,
  pagesRef,
  pagesViewportRef,
  zoom,
  focus,
  getDocument,
  setZoom,
  save,
  loadDocument,
  loadDocumentBuffer: loadBuffer,
  onPrint: props.onPrint,
});
defineExpose(exposed);
</script>

<style scoped>
.docx-editor-vue {
  display: flex;
  flex-direction: column;
  height: 100%;
  min-height: 0;
}
.docx-editor-vue__toolbar-shell {
  flex: 0 0 auto;
}
.docx-editor-vue__hidden-pm {
  position: absolute;
  width: 1px;
  height: 1px;
  overflow: hidden;
  opacity: 0;
  pointer-events: none;
}
.docx-editor-vue__editor-scroll {
  flex: 1 1 auto;
  min-height: 0;
  overflow: auto;
}
.docx-editor-vue__editor-area {
  position: relative;
  display: flex;
  flex-direction: column;
  align-items: center;
}
.docx-editor-vue__pages-viewport {
  position: relative;
  width: 100%;
  display: flex;
  flex-direction: column;
  align-items: center;
}
.docx-editor-vue__error {
  padding: 12px;
  color: var(--destructive, #b00020);
}
.docx-editor-vue__loading {
  padding: 12px;
  opacity: 0.7;
}
</style>
