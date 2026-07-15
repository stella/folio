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

  Wired composables: pointer interactions (`usePagesPointer`: multi-click,
  drag-select, table quick-insert, hyperlink popup, page indicator), image
  actions + selection overlay (`useImageActions`), right-click menus
  (`useContextMenus`), hyperlink management (`useHyperlinkManagement`), the
  comment/tracked-change sidebar (`useTrackedChanges` + `useCommentSidebarItems`
  inside `UnifiedSidebar`), and the table context toolbar (`TableToolbar`).

  Remaining adapter-specific surfaces:
   - externalPlugins / i18n-locale props are not on `DocxEditorProps`, so the
     host plugin list is empty and the locale defaults to `en`.
   - colorMode prop + useColorMode: dark mode is fixed light here.
   - Header/footer inline editing is not available yet. The title-bar MenuBar is
     wired (File/Format/Insert/Help), including image/table/page-break/TOC.
     The watermark item remains unavailable until its dialog is implemented.
-->
<template>
  <div
    :class="[
      'docx-editor-vue ep-root paged-editor',
      className,
      displayModeClass,
      { 'paged-editor--readonly': readOnly },
    ]"
    :style="style"
  >
    <div class="docx-editor-vue__toolbar-shell">
      <DocxEditorMenuBar
        :show-table-insert="props.showTableInsert !== false"
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
        :show-review-controls="showReviewControls"
        :track-changes-on="trackChangesOn"
        :display-mode="displayMode"
        :read-only="readOnly"
        :comments-sidebar-open="showSidebar"
        :image-context="imageToolbarContext"
        :theme="theme ?? null"
        v-bind="toolbarDynamicProps"
        @find-replace="showFindReplace = true"
        @insert-link="showHyperlink = true"
        @insert-symbol="showInsertSymbol = true"
        @insert-page-break="handleInsertPageBreakAction"
        @page-setup="showPageSetup = true"
        @toggle-outline="outlineSidebar.handleToggleOutline"
        @apply-style="handleApplyStyle"
        @zoom-in="zoomIn"
        @zoom-out="zoomOut"
        @zoom-set="setZoom"
        @toggle-sidebar="showSidebar = !showSidebar"
        @mode-change="setEditorMode"
        @toggle-track-changes="toggleTrackChanges"
        @update:display-mode="setDisplayMode"
        @image-properties="showImageProperties = true"
        @image-wrap-type="handleToolbarImageWrap"
        @image-transform="handleImageTransform"
      >
        <template #table-context>
          <TableToolbar
            :view="editorView"
            :get-commands="getCommands"
            :state-tick="stateTick"
            :theme="theme ?? null"
          />
        </template>
        <template v-if="$slots['toolbar-extra'] || toolbarExtra" #toolbar-extra>
          <slot name="toolbar-extra">
            <VNodeRenderer v-if="toolbarExtra" :node="toolbarExtra" />
          </slot>
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
      :selected-image-pm-pos="selectedImage?.pmPos ?? null"
      :section-properties="currentSectionProps"
      @insert-symbol="handleInsertSymbol"
      @hyperlink-submit="handleHyperlinkSubmit"
      @hyperlink-remove="handleHyperlinkRemove"
      @page-setup-apply="handlePageSetupApply"
    />

    <div v-if="parseError" class="docx-editor-vue__error">{{ parseError }}</div>
    <div v-else-if="showLoadingIndicator" class="docx-editor-vue__loading">
      <slot name="loading-indicator">
        <VNodeRenderer v-if="loadingIndicator" :node="loadingIndicator" />
        <template v-else>{{ t("loadingDocument") }}</template>
      </slot>
    </div>
    <div v-else-if="!isReady && !hasDocumentInput" class="docx-editor-vue__placeholder">
      <slot name="placeholder">
        <VNodeRenderer v-if="placeholder" :node="placeholder" />
        <template v-else>No document loaded</template>
      </slot>
    </div>

    <div ref="hiddenPmRef" class="docx-editor-vue__hidden-pm paged-editor__hidden-pm" />

    <div ref="editorScrollRef" class="docx-editor-vue__editor-scroll" @scroll="handleEditorScroll">
      <div class="docx-editor-vue__editor-area">
        <!-- Horizontal ruler — sticky-top, centered over the page so it scrolls
             horizontally with the document. Mirrors React's DocxEditor ruler
             placement; gated on `rulerVisible && !readOnly`. -->
        <div v-if="rulerVisible && !readOnly" class="docx-editor-vue__ruler-h">
          <HorizontalRuler
            :section-props="currentSectionProps"
            :zoom="zoom"
            :unit="rulerUnit ?? 'inch'"
            :editable="!readOnly"
            :indent-left="paragraphIndent.indentLeft"
            :indent-right="paragraphIndent.indentRight"
            :first-line-indent="paragraphIndent.firstLineIndent"
            :hanging-indent="paragraphIndent.hangingIndent"
            :show-first-line-indent="true"
            :tab-stops="paragraphIndent.tabs"
            @left-margin-change="handleLeftMarginChange"
            @right-margin-change="handleRightMarginChange"
            @indent-left-change="handleIndentLeftChange"
            @indent-right-change="handleIndentRightChange"
            @first-line-indent-change="handleFirstLineIndentChange"
            @tab-stop-remove="handleTabStopRemove"
          />
        </div>

        <!-- Vertical ruler — far-left gutter, does not track page centering
             (word-processor gutter convention). Mirrors React's placement. -->
        <div v-if="rulerVisible && !readOnly" class="docx-editor-vue__ruler-v">
          <VerticalRuler
            :section-props="currentSectionProps"
            :zoom="zoom"
            :unit="rulerUnit ?? 'inch'"
            :editable="!readOnly"
            @top-margin-change="handleTopMarginChange"
            @bottom-margin-change="handleBottomMarginChange"
          />
        </div>

        <div
          ref="pagesViewportRef"
          class="docx-editor-vue__pages-viewport"
          @wheel="handleWheelZoomGated"
          @mousedown="handlePagesMouseDown"
          @mousemove="handlePagesMouseMove"
          @click="handlePagesClick"
          @contextmenu.prevent="handleContextMenu"
        >
          <div
            class="docx-editor-vue__editor-content-wrapper"
            :style="{
              position: 'relative',
              display: 'flex',
              flexDirection: 'column',
              minHeight: '100%',
            }"
          >
            <div
              ref="pagesRef"
              class="docx-editor-vue__pages paged-editor__pages"
              :style="pagesContainerStyle"
            />

            <!-- Decoration origin: the coordinate anchor core's range projection
                 looks up (`[data-testid="selection-overlay"]`) and the container
                 the projected (unscaled) rects are painted into. Sibling of the
                 pages so the projection resolves it via `pagesContainer.parentElement`. -->
            <div
              data-testid="selection-overlay"
              class="docx-editor-vue__decoration-origin"
              aria-hidden="true"
            >
              <AnonymizationRectsOverlay
                :get-view="() => editorView"
                :get-pages-container="() => pagesRef"
                :editor-state="editorState"
                :zoom="zoom"
                :layout="layout"
                :blocks="blocks"
                :measures="measures"
                :on-term-click="props.onAnonymizationTermClick"
                :selected-canonical="props.selectedAnonymizationCanonical"
                :selection-seq="props.anonymizationSelectionSeq"
              />
              <TemplateDirectivesOverlay
                v-if="props.showTemplateDirectives"
                :get-view="() => editorView"
                :get-pages-container="() => pagesRef"
                :editor-state="editorState"
                :zoom="zoom"
                :layout="layout"
                :blocks="blocks"
                :measures="measures"
              />
              <DecorationLayer
                :get-view="() => editorView"
                :get-pages-container="() => pagesRef"
                :zoom="zoom"
                :transaction-version="stateTick"
                :sync-coordinator="syncCoordinator"
              />
            </div>
          </div>

          <ImageSelectionOverlay
            :image-info="selectedImage"
            :zoom="zoom"
            :view="editorView"
            @deselect="selectedImage = null"
            @interact-start="imageInteracting = true"
            @interact-end="imageInteracting = false"
            @context-menu="handleSelectedImageContextMenu"
          />

          <HyperlinkPopup
            :data="hyperlinkPopupData"
            :read-only="readOnly"
            @navigate="handleHyperlinkPopupNavigate"
            @copy="handleHyperlinkPopupCopy"
            @edit="handleHyperlinkPopupEdit"
            @remove="handleHyperlinkPopupRemove"
            @close="hyperlinkPopupData = null"
          />

          <button
            v-if="tableInsertButton"
            type="button"
            class="docx-editor-vue__table-insert-btn"
            :style="{ left: tableInsertButton.x + 'px', top: tableInsertButton.y + 'px' }"
            :aria-label="tableInsertButton.type === 'row' ? 'Insert row' : 'Insert column'"
            @mousedown="handleTableInsertClick"
          >
            +
          </button>

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
            :is-adding-comment="commentLifecycle.isAddingComment.value"
            :add-comment-y-position="commentLifecycle.addCommentYPosition.value"
            :pages-container="pagesRef"
            :page-width-px="pageWidthPx"
            :zoom="zoom"
            :active-item-id="activeSidebarItem"
            @close="showSidebar = false"
            @update:active-item-id="(id: string | null) => (activeSidebarItem = id)"
            @add-comment="commentLifecycle.handleAddComment"
            @cancel-add-comment="commentLifecycle.handleCancelAddComment"
            @comment-reply="commentManagement.handleReply"
            @comment-resolve="commentManagement.handleResolve"
            @comment-unresolve="commentManagement.handleUnresolve"
            @comment-delete="commentManagement.handleDelete"
            @accept-change="commentManagement.handleAcceptChange"
            @reject-change="commentManagement.handleRejectChange"
            @accept-change-by-id="commentManagement.handleAcceptChangeById"
            @reject-change-by-id="commentManagement.handleRejectChangeById"
            @tracked-change-reply="commentManagement.handleTrackedChangeReply"
          />

          <button
            v-if="commentLifecycle.floatingCommentButton.value && !readOnly"
            type="button"
            class="docx-editor-vue__add-comment-btn"
            :style="{
              top: commentLifecycle.floatingCommentButton.value.top + 'px',
              left: commentLifecycle.floatingCommentButton.value.left + 'px',
            }"
            aria-label="Add comment"
            title="Add comment"
            @mousedown.prevent.stop="commentLifecycle.startAddComment"
          >
            <MaterialSymbol name="add_comment" :size="18" />
          </button>
        </div>

        <OutlineToggleButton
          v-if="!showOutline"
          :left-offset="12"
          @toggle="outlineSidebar.handleToggleOutline"
        />

        <PageIndicator
          v-if="scrollPageInfo.totalPages > 1"
          :current-page="scrollPageInfo.currentPage"
          :total-pages="scrollPageInfo.totalPages"
          :visible="scrollPageInfo.visible"
        />

        <DocumentOutline
          :is-open="showOutline"
          :headings="outlineHeadings"
          :get-scroll-container="() => pagesRef"
          @close="showOutline = false"
          @navigate="outlineSidebar.handleOutlineNavigate"
        />
      </div>
    </div>

    <DocxEditorOverlays
      :read-only="readOnly"
      :context-menu="contextMenu"
      :image-context-menu="imageContextMenu"
      :image-context-menu-text-actions="imageContextMenuTextActions"
      :custom-text-menu-items="customTextMenuItems"
      :can-open-image-properties="selectedImage !== null"
      @context-menu-action="handleContextMenuAction"
      @close-context-menu="contextMenu.isOpen = false"
      @image-wrap-select="handleImageWrapSelect"
      @close-image-context-menu="imageContextMenu = null"
      @open-image-properties="showImageProperties = true"
    />
  </div>
</template>

<script setup lang="ts">
import {
  computed,
  onBeforeUnmount,
  onMounted,
  ref,
  shallowRef,
  watch,
  type FunctionalComponent,
  type VNodeChild,
} from "vue";

import {
  extractSelectionState,
  insertImageFromFile,
  insertPageBreakInView,
  insertTableInView,
  insertTableOfContentsInView,
} from "@stll/folio-core/prosemirror";
import { extractSelectionContext } from "@stll/folio-core/prosemirror/plugins/selectionTracker";
import { inspectDocxCompatibility } from "@stll/folio-core/docx/compatibility";
import {
  clearAllCaches,
  resetCanvasContext,
} from "@stll/folio-core/layout-bridge/engine/measuring";
import { onFontsLoaded } from "@stll/folio-core/utils/fontLoader";
import { twipsToPixels } from "@stll/folio-core/paged-layout/sectionGeometry";
import type { Comment } from "@stll/folio-core/types/content";
import type { Document, SectionProperties, Style } from "@stll/folio-core/types/document";
import type { HeadingInfo } from "@stll/folio-core/utils/headingCollector";

import AnonymizationRectsOverlay from "./AnonymizationRectsOverlay.vue";
import CommentMarginMarkers from "./CommentMarginMarkers.vue";
import DecorationLayer from "./DecorationLayer.vue";
import DocumentOutline from "./DocumentOutline.vue";
import DocxEditorDialogs from "./DocxEditor/DocxEditorDialogs.vue";
import DocxEditorMenuBar from "./DocxEditor/DocxEditorMenuBar.vue";
import DocxEditorOverlays from "./DocxEditor/DocxEditorOverlays.vue";
import type { DocxEditorProps, EditorMode } from "./DocxEditor/types";
import {
  isTrackChangesMode,
  toggledTrackChangesMode,
} from "@stll/folio-core/managers/EditorModeManager";
import type { DisplayMode } from "@stll/folio-core/managers/EditorModeManager";
import ImageSelectionOverlay from "./ImageSelectionOverlay.vue";
import OutlineToggleButton from "./OutlineToggleButton.vue";
import PageIndicator from "./PageIndicator.vue";
import TemplateDirectivesOverlay from "./TemplateDirectivesOverlay.vue";
import type { TrackedChangeEntry } from "./sidebar/sidebarUtils";
import Toolbar from "./Toolbar.vue";
import HorizontalRuler from "./ui/HorizontalRuler.vue";
import HyperlinkPopup from "./ui/HyperlinkPopup.vue";
import MaterialSymbol from "./ui/MaterialSymbol.vue";
import TableToolbar from "./ui/TableToolbar.vue";
import VerticalRuler from "./ui/VerticalRuler.vue";
import UnifiedSidebar from "./UnifiedSidebar.vue";

import { useCommentLifecycle } from "../composables/useCommentLifecycle";
import { useCommentManagement } from "../composables/useCommentManagement";
import { useContextMenus } from "../composables/useContextMenus";
import { useDocumentLifecycle } from "../composables/useDocumentLifecycle";
import { useDocxEditor } from "../composables/useDocxEditor";
import { useDocxEditorRefApi } from "../composables/useDocxEditorRefApi";
import { useFormattingActions } from "../composables/useFormattingActions";
import { useHyperlinkManagement } from "../composables/useHyperlinkManagement";
import { useImageActions } from "../composables/useImageActions";
import { useKeyboardShortcuts } from "../composables/useKeyboardShortcuts";
import { useOutlineSidebar } from "../composables/useOutlineSidebar";
import { usePageSetupControls } from "../composables/usePageSetupControls";
import { usePagesPointer } from "../composables/usePagesPointer";
import { useRemoteSelectionSync } from "../composables/useRemoteSelectionSync";
import { useSelectionSync } from "../composables/useSelectionSync";
import { provideDocxPortalClass } from "../composables/usePortalClass";
import { useTableResize } from "../composables/useTableResize";
import { useTrackedChanges } from "../composables/useTrackedChanges";
import { useZoom } from "../composables/useZoom";
import type { FontOption } from "../utils/fontOptions";
import { loadHostFontFaces, removeFontFaces } from "../utils/hostFonts";
import { provideLocale, useTranslation } from "../i18n";
import { provideFolioUI } from "../ui/folio-ui";

const props = withDefaults(defineProps<DocxEditorProps>(), {
  documentBuffer: null,
  document: null,
  showToolbar: true,
  showZoomControl: true,
  showReviewControls: true,
  readOnly: false,
  author: "User",
  mode: "editing",
  initialZoom: 1,
  showOutline: false,
  className: "",
  showTableInsert: true,
});

const emit = defineEmits<{
  (e: "change", doc: Document): void;
  (e: "update:document", doc: Document | null): void;
  (e: "error", error: Error): void;
  (e: "ready"): void;
  (e: "rename", name: string): void;
  (e: "menu-action", action: string): void;
  (e: "mode-change", mode: EditorMode): void;
  (e: "comments-change", comments: Comment[]): void;
}>();

// PORT-BLOCKED: no i18n/locale prop on the fork's DocxEditorProps yet — default to `en`.
provideLocale();
// Provide the consumer's UI-injection overrides (or the built-in defaults) to
// the chrome subtree. Descendant chrome (Toolbar, FormattingBar) resolves the
// injected primitives via `useFolioUI`, so `components.ColorPicker` etc. take
// effect. Provided once at setup; the prop is not expected to change identity.
provideFolioUI(props.components);
// PORT-BLOCKED: no colorMode prop; share a fixed-light token scope with teleported chrome.
const isDark = ref(false);
provideDocxPortalClass(isDark);

const editorMode = ref<EditorMode>(props.mode);
const readOnly = computed(() => props.readOnly || editorMode.value === "viewing");

// Review controls (mirrors React's `useEditorMode` display state + toggle):
//  - `showReviewControls` gates the toolbar's track-changes toggle + markup
//    display-mode selector (documented default: true).
//  - `trackChangesOn` derives from the editing mode (suggesting === on).
//  - `displayMode` is independent local state controlling how tracked changes
//    render; the host applies the `folio-root--<mode>` class, exactly as React.
// Documented default (true) is supplied via `withDefaults`; Vue's boolean-prop
// casting resolves an absent boolean to `false`, so the default must live there
// rather than a `?? true` fallback (which the cast would pre-empt).
const showReviewControls = computed(() => props.showReviewControls);
const trackChangesOn = computed(() => isTrackChangesMode(editorMode.value));
const displayMode = ref<DisplayMode>("all-markup");

function toggleTrackChanges(): void {
  setEditorMode(toggledTrackChangesMode(editorMode.value));
}

function setDisplayMode(mode: DisplayMode): void {
  displayMode.value = mode;
}

// Root modifier class driving the tracked-change display CSS (mirrors React's
// `folio-root--<mode>`). `all-markup` is the default render, so it adds no class.
const displayModeClass = computed(() =>
  displayMode.value === "all-markup" ? "" : `folio-root--${displayMode.value}`,
);

const { t } = useTranslation();

// Stable functional component that renders a host-supplied VNodeChild
// (placeholder / loadingIndicator). Declaring it once — rather than minting a
// fresh `() => node` wrapper per render — keeps a constant component identity,
// so Vue patches the node in place instead of remounting the subtree on every
// reactive tick (each `stateTick` bump). `props: ["node"]` consumes `node` as a
// prop so it is not forwarded as a stray DOM attribute.
const VNodeRenderer: FunctionalComponent<{ node: VNodeChild }> = ({ node }) => node;
VNodeRenderer.props = ["node"];

// Whether the host supplied a document to load. Drives the loading-vs-placeholder
// split while `isReady` is false (mirrors React's history.state gate).
const hasDocumentInput = computed(() => props.documentBuffer != null || props.document != null);

// True while a document is currently present (painted) — the Vue analogue of
// React's truthy `history.state`, which tracks the *current* document, not a
// permanent "ever rendered" latch. Gates `preserveDocumentWhileLoading`: a swap
// tears the PM view down (isReady flips false) but the previous pages stay
// painted in `pagesRef` until the new layout runs, so suppressing the loading
// interstitial keeps the prior document visible until the next paint replaces
// it. Set true once a document paints; cleared on explicit unload (below) so a
// later reload shows the loading indicator again, exactly as React does when
// `history.state` returns to null.
const hasRenderedDocumentOnce = ref(false);

// Mirror React's `history.state` going null on unload: when the host clears the
// document, drop the "currently present" flag so the next load is not treated
// as a seamless swap (which would suppress the loading indicator and blank the
// screen — no prior paint survives an unload).
watch(hasDocumentInput, (hasInput) => {
  if (!hasInput) {
    hasRenderedDocumentOnce.value = false;
  }
});

// ---- Template refs (paint targets) --------------------------------------
const hiddenPmRef = ref<HTMLElement | null>(null);
const pagesRef = ref<HTMLElement | null>(null);
const pagesViewportRef = ref<HTMLElement | null>(null);
const editorScrollRef = ref<HTMLElement | null>(null);

// Ruler visibility — driven by the `showRuler` prop (default off). Renders the
// horizontal + vertical rulers, gated additionally on `!readOnly` (matching
// React, which hides the rulers in read-only mode).
const rulerVisible = computed(() => props.showRuler ?? false);

// ---- Chrome state -------------------------------------------------------
const stateTick = ref(0);
const showFindReplace = ref(false);
const showHyperlink = ref(false);
// The keyboard-shortcuts dialog is not available yet, so F1 / Ctrl+/ only
// preserves the intended open state for the future dialog surface.
const showKeyboardShortcuts = ref(false);
const showInsertSymbol = ref(false);
const showImageProperties = ref(false);
const showPageSetup = ref(false);
const showOutline = ref(props.showOutline);
const showSidebar = ref(false);
const activeSidebarItem = ref<string | null>(null);
const bookmarks = shallowRef<{ name: string; label?: string }[]>([]);

// Populated by `useOutlineSidebar` — collected lazily on outline open and
// re-collected on doc changes while the panel stays open (see below).
const outlineHeadings = shallowRef<HeadingInfo[]>([]);

const {
  zoom,
  zoomPercent,
  isMinZoom,
  isMaxZoom,
  setZoom,
  zoomIn,
  zoomOut,
  handleWheel: handleZoomWheel,
  handleKeyDown: handleZoomKeyDown,
  ZOOM_PRESETS,
} = useZoom(props.initialZoom);

// Ctrl/Cmd+wheel + trackpad-pinch zoom, gated on the `enableWheelZoom` prop
// (default enabled, matching React). Consulted here so the flag is honored.
function handleWheelZoomGated(event: WheelEvent): void {
  if (props.enableWheelZoom === false) {
    return;
  }
  handleZoomWheel(event);
}

// Global keyboard shortcuts (F1 keyboard-shortcuts dialog, Ctrl+=/-/0 zoom,
// Ctrl+F/H find-replace, Ctrl+K hyperlink, Ctrl+/ shortcuts toggle). Installs
// its own window-level listener on mount and tears it down on unmount, so no
// further wiring is needed here. Mirrors React's Cmd/Ctrl+F find-open path
// (`useKeyboardShortcuts` in `packages/react/src/components/hooks`), extended
// with the Vue adapter's additional shortcuts.
useKeyboardShortcuts({
  showKeyboardShortcuts,
  showFindReplace,
  showHyperlink,
  handleZoomKeyDown,
});

// ---- Pipeline -----------------------------------------------------------
const {
  editor,
  editorView,
  editorState,
  remoteSelections,
  isReady,
  isDirty,
  parseError,
  layout,
  blocks,
  measures,
  syncCoordinator,
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
  password: () => props.password,
  documentKey: () => props.documentKey,
  externalPlugins: [],
  collaboration: () => props.collaboration,
  // Anonymization highlights + template directives are driven by the overlay
  // components below; these thread the plugin callbacks and the directive gate.
  onAnonymizationMatchesChange: (matches) => props.onAnonymizationMatchesChange?.(matches),
  showTemplateDirectives: () => props.showTemplateDirectives,
  onSlashMenuChange: (state) => props.onSlashMenuChange?.(state),
  onSlashMenuKeyAction: (action) => props.onSlashMenuKeyAction?.(action) ?? false,
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
    // Selected plain text on every selection-bearing transaction. Atom inline
    // nodes (tab, hard_break) collapse to a single space; empty when collapsed.
    // Mirrors React's PagedEditor `onSelectionTextChange` (textBetween(" ", " ")).
    if (props.onSelectionTextChange) {
      const { from, to } = state.selection;
      const text = from === to ? "" : state.doc.textBetween(from, to, " ", " ");
      props.onSelectionTextChange({ from, to, text });
    }
  },
  onEditorViewReady: (view) => props.onEditorViewReady?.(view),
  onReadOnlyEditAttempt: () => props.onReadonlyEditAttempt?.(),
  // Selective-save feature flags + tripwire observability, mirroring React's
  // DocxEditor save path. Read fresh on each save() so prop updates are honored.
  featureFlags: () => props.featureFlags,
  // Pass through only when the host actually provided a callback so it stays
  // `undefined` otherwise, mirroring React where the tripwire comparison is
  // gated on `onSelectiveSaveTripwire` being defined (an absent callback must
  // skip the expensive compareSelectiveVsFull, not run it into a no-op).
  onSelectiveSaveTripwire: props.onSelectiveSaveTripwire
    ? (result) => props.onSelectiveSaveTripwire?.(result)
    : undefined,
});

// Show the loading interstitial while a document is loading, EXCEPT when the host
// opted into `preserveDocumentWhileLoading` and a prior document is already
// painted — then the swap is seamless (previous paint stays up). Mirrors React's
// `status === "loading" && (!preserveDocumentWhileLoading || !history.state)`.
const showLoadingIndicator = computed(
  () =>
    !isReady.value &&
    hasDocumentInput.value &&
    !(props.preserveDocumentWhileLoading === true && hasRenderedDocumentOnce.value),
);

// Host custom font faces (the `fonts` prop). Register on mount + whenever the
// prop identity changes, then re-measure so metrics reflect the new faces — a
// face registered after the first layout would otherwise keep its fallback
// metrics until the next edit. The cleanup unregisters the faces and re-measures
// their absence. Mirrors React PagedEditor's hostFonts effect (same best-effort
// FontFace path + font-ready re-layout).
function remeasureForFontChange(): void {
  const view = editorView.value;
  if (!view) {
    return;
  }
  resetCanvasContext();
  clearAllCaches();
  reLayout();
}

watch(
  () => props.fonts,
  (fonts, _prev, onCleanup) => {
    if (!fonts || fonts.length === 0) {
      return;
    }
    let cancelled = false;
    let registered: FontFace[] = [];
    void loadHostFontFaces(fonts).then((faces) => {
      if (cancelled) {
        removeFontFaces(faces);
        return undefined;
      }
      registered = faces;
      if (faces.length === 0) {
        return undefined;
      }
      remeasureForFontChange();
      return undefined;
    });
    onCleanup(() => {
      cancelled = true;
      if (registered.length === 0) {
        return;
      }
      removeFontFaces(registered);
      remeasureForFontChange();
    });
  },
  { immediate: true },
);

// ---- Feature composables (order: image → hyperlink → pointer → context) --
const {
  selectedImage,
  imageInteracting,
  imageToolbarContext,
  handleToolbarImageWrap,
  handleImageTransform,
} = useImageActions({ editorView, zoom, stateTick, getCommands });

const selectionSync = useSelectionSync({
  editorView,
  pagesRef,
  zoom,
  selectedImage,
  syncCoordinator,
  imageInteracting,
});

useRemoteSelectionSync({
  pagesRef,
  remoteSelections,
  syncCoordinator,
  zoom,
});

const {
  hyperlinkPopupData,
  handleHyperlinkSubmit,
  handleHyperlinkRemove,
  handleHyperlinkPopupNavigate,
  handleHyperlinkPopupEdit,
  handleHyperlinkPopupRemove,
} = useHyperlinkManagement({ editorView, getCommands });

const tableResize = useTableResize();

// Pages-area pointer gestures. Header/footer double-click is intentionally
// omitted: the fork has no inline HF editor or persistent HF PMs, so the HF
// `emit` (change) path and its handlers stay unused here.
const {
  tableInsertButton,
  scrollPageInfo,
  resolvePos,
  setPmSelection,
  scrollVisiblePositionIntoView,
  handlePagesMouseDown,
  handlePagesMouseMove,
  handlePagesClick,
  handleTableInsertClick,
} = usePagesPointer({
  editorView,
  pagesRef,
  pagesViewportRef,
  selectedImage,
  imageInteracting,
  hyperlinkPopupData,
  readOnly,
  zoom,
  layout,
  tableResize: {
    tryStartResize: tableResize.tryStartResize,
    isResizing: computed(() => tableResize.isResizing()),
  },
  getCommands,
  getDocument,
  reLayout,
  emit: () => {},
  clearOverlay: selectionSync.clearOverlay,
});

// Document Outline: heading collection + navigate-to-heading, driven off the
// visible pages viewport (the hidden PM is off-screen — see
// `scrollVisiblePositionIntoView`'s doc comment). `extractCommentsAndChanges`
// is unused here: DocxEditor.vue's own sidebar toggle (below) already reads
// comments reactively off `commentManagement.comments`, so only the
// outline-specific handlers from this composable are wired.
const outlineSidebar = useOutlineSidebar({
  editorView,
  showOutline,
  showSidebar,
  outlineHeadings,
  activeSidebarItem,
  extractCommentsAndChanges: () => {},
  scrollToVisiblePosition: scrollVisiblePositionIntoView,
});

const {
  contextMenu,
  imageContextMenu,
  imageContextMenuTextActions,
  customTextMenuItems,
  handleContextMenu,
  handleSelectedImageContextMenu,
  handleImageWrapSelect,
  handleContextMenuAction,
} = useContextMenus({
  editorView,
  selectedImage,
  zoom,
  showImageProperties,
  getCommands,
  clearOverlay: selectionSync.clearOverlay,
  setPmSelection,
  resolvePos,
  // Host `custom:` entries perform edits, so drop them in read-only mode —
  // mirrors React, whose context-menu builder returns only copy/selectAll when
  // readOnly and never surfaces the custom list.
  customContextMenuItems: () => (readOnly.value ? undefined : props.customContextMenuItems),
  onCustomContextAction: (id, range) => props.onCustomContextAction?.(id, range),
});

// Tracked-change sidebar cards: derive entries from the live editor state via
// core's extractTrackedChanges, re-running whenever a transaction bumps
// stateTick. `entries` is core's TrackedChangeEntry, which the sidebar's
// (re-exported) type now matches exactly, so no reconciliation is needed.
const trackedChangesResult = useTrackedChanges(editorView, stateTick);
const trackedChanges = computed<TrackedChangeEntry[]>(() => trackedChangesResult.value.entries);

// Comment thread state + mutation handlers (add/reply/resolve/accept/reject),
// wired to core's comment ops. `onCommentsChange` fans out to the host prop and
// the Vue-idiomatic `comments-change` emit.
const commentManagement = useCommentManagement({
  editorView,
  getDocument,
  author: () => props.author,
  commentsProp: () => props.comments,
  onCommentsChange: (next) => {
    props.onCommentsChange?.(next);
    emit("comments-change", next);
  },
  reLayout,
});
const comments = commentManagement.comments;

// Interactive add-comment flow: floating button, pending highlight, submit.
const commentLifecycle = useCommentLifecycle({
  editorView,
  pagesViewport: pagesViewportRef,
  pagesContainer: pagesRef,
  readOnly,
  createComment: commentManagement.createComment,
  pushComment: commentManagement.pushComment,
  reLayout,
  showSidebar,
  setActiveSidebarItem: (id) => {
    activeSidebarItem.value = id;
  },
});

function handleHyperlinkPopupCopy(href: string): void {
  void navigator.clipboard?.writeText(href);
  hyperlinkPopupData.value = null;
}

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

const pageWidthPx = computed(
  () => twipsToPixels(currentSectionProps.value?.pageWidth ?? 12240) * zoom.value,
);

const resolvedCommentIds = computed(() => {
  const ids = new Set<number>();
  for (const comment of comments.value) if (comment.done) ids.add(comment.id);
  return ids;
});

const pagesContainerStyle = computed(() => ({
  transform: zoom.value === 1 ? undefined : `scale(${zoom.value})`,
  transformOrigin: "top center",
  transition: "transform 0.2s ease",
}));

// ---- Feature handlers ----------------------------------------------------
const {
  handleApplyStyle,
  handleInsertSymbol,
  handleInsertSectionBreakNextPage,
  handleInsertSectionBreakContinuous,
} = useFormattingActions({
  editorView,
  getDocument,
});

// Insert flow: prefer the host `onInsert*` prop when provided, otherwise fall
// back to the core view-level helpers (mirrors the contract documented on
// DocxEditorProps). Both the title-bar MenuBar and the toolbar drive these.
function handleInsertImageAction(): void {
  if (props.onInsertImage) {
    props.onInsertImage();
    return;
  }
  const view = editorView.value;
  if (!view) {
    return;
  }
  // No host handler: open the OS file picker and insert directly, matching
  // React's `insertImageFromFile` flow (no intermediate dialog).
  const input = document.createElement("input");
  input.type = "file";
  input.accept = "image/*";
  input.style.display = "none";
  input.addEventListener("change", () => {
    const file = input.files?.[0];
    if (file) {
      void insertImageFromFile(view, file, () => view.focus());
    }
    // Some browsers (iOS Safari, restricted iframes) require the input to be
    // connected to the DOM for `.click()` to open the picker; mirror React's
    // mounted hidden input by attaching it, then clean up after selection.
    input.remove();
  });
  document.body.appendChild(input);
  input.click();
}

function handleInsertTableAction(rows: number, cols: number): void {
  if (props.onInsertTable) {
    props.onInsertTable(rows, cols);
    return;
  }
  const view = editorView.value;
  if (!view) {
    return;
  }
  insertTableInView(view, rows, cols);
  view.focus();
}

function handleInsertPageBreakAction(): void {
  if (props.onInsertPageBreak) {
    props.onInsertPageBreak();
    return;
  }
  const view = editorView.value;
  if (!view) {
    return;
  }
  insertPageBreakInView(view);
  view.focus();
}

function handleInsertTOCAction(): void {
  if (props.onInsertTOC) {
    props.onInsertTOC();
    return;
  }
  const view = editorView.value;
  if (!view) {
    return;
  }
  insertTableOfContentsInView(view);
  view.focus();
}

// Page-setup + ruler edit handlers. The margin / indent / tab-stop handlers feed
// the horizontal + vertical rulers (showRuler); the page-setup dialog reuses
// handlePageSetupApply. Margin / page-size writes mutate section properties
// directly (no PM transaction), so they notify the host through `onChange` here;
// indent / tab-stop edits dispatch PM commands and already notify via the
// transaction pipeline.
const {
  handlePageSetupApply,
  handleLeftMarginChange,
  handleRightMarginChange,
  handleTopMarginChange,
  handleBottomMarginChange,
  handleIndentLeftChange,
  handleIndentRightChange,
  handleFirstLineIndentChange,
  handleTabStopRemove,
} = usePageSetupControls({
  editorView,
  getDocument,
  readOnly,
  stateTick,
  reLayout,
  onChange: (doc) => {
    props.onChange?.(doc);
    emit("change", doc);
    emit("update:document", doc);
  },
});

// Paragraph indent snapshot for the horizontal ruler's indent handles. Derived
// from core's extractSelectionContext, re-run on every selection/doc tick.
// Mirrors React's `state.paragraph*` ruler inputs.
const paragraphIndent = computed(() => {
  void stateTick.value;
  const view = editorView.value;
  const pf = view ? extractSelectionContext(view.state).paragraphFormatting : {};
  return {
    indentLeft: pf.indentLeft ?? 0,
    indentRight: pf.indentRight ?? 0,
    firstLineIndent: pf.indentFirstLine ?? 0,
    hangingIndent: pf.hangingIndent ?? false,
    tabs: pf.tabs ?? null,
  };
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
      outlineSidebar.handleToggleOutline();
      break;
    case "toggleSidebar":
      showSidebar.value = !showSidebar.value;
      break;
    case "insertImage":
      handleInsertImageAction();
      break;
    case "insertPageBreak":
      handleInsertPageBreakAction();
      break;
    case "insertSectionBreakNextPage":
      handleInsertSectionBreakNextPage();
      break;
    case "insertSectionBreakContinuous":
      handleInsertSectionBreakContinuous();
      break;
    case "insertTOC":
      handleInsertTOCAction();
      break;
    default:
      break;
  }
  emit("menu-action", action);
}

// Insert > Table (menu grid picker) routes through the shared insert handler so
// it honors the host `onInsertTable` prop or falls back to the core helper.
function handleMenuTableInsert(rows: number, cols: number): void {
  handleInsertTableAction(rows, cols);
}

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
    // Host-driven open (as opposed to a user click through
    // `handleToggleOutline`) still needs headings collected.
    outlineSidebar.recomputeHeadingsIfOpen();
  },
);

// Scroll-offset bridge on the document scroll container. `initialScrollTop` is
// applied once per load (rAF-deferred so the first paint settles first);
// `onScrollTopChange` fires the raw scrollTop on every scroll.
const initialScrollAppliedRef = ref(false);

function handleEditorScroll(event: Event): void {
  const target = event.currentTarget;
  if (target instanceof HTMLElement) {
    props.onScrollTopChange?.(target.scrollTop);
  }
}

function applyInitialScrollTop(): void {
  if (props.initialScrollTop === undefined || initialScrollAppliedRef.value) {
    return;
  }
  initialScrollAppliedRef.value = true;
  const top = props.initialScrollTop;
  requestAnimationFrame(() => {
    if (editorScrollRef.value) {
      editorScrollRef.value.scrollTop = top;
    }
  });
}

watch(isReady, (ready) => {
  if (!ready) {
    // Document unloaded/swapped: re-arm the one-shot so `initialScrollTop` is
    // applied to the next document too, not just the first.
    initialScrollAppliedRef.value = false;
    return;
  }
  // Mark that a document has painted, so a later `preserveDocumentWhileLoading`
  // swap keeps the prior pages visible instead of flashing the loading state.
  hasRenderedDocumentOnce.value = true;

  // A document just (re)painted: re-collect headings if the outline panel is
  // already open — covers the initial-mount `showOutline` prop and a document
  // swap while the panel stays open (an outline toggle can't fire in either
  // case, since it was already open before the view existed).
  outlineSidebar.recomputeHeadingsIfOpen();

  // Populate the sidebar from comments embedded in the loaded document
  // (uncontrolled only); controlled hosts own the array via `comments`.
  commentManagement.seedFromDocument();

  // Surface the parsed document's editing-compatibility report (once per load).
  const doc = getDocument();
  if (doc) {
    props.onCompatibilityChange?.(inspectDocxCompatibility(doc));
  }

  // Auto-open the review sidebar once per load when the document arrives with
  // comments or tracked changes (default on). `sidebarAutoOpenedRef` is reset by
  // useDocumentLifecycle on every swap. Mirrors React's autoOpenReviewSidebar.
  if (
    props.autoOpenReviewSidebar !== false &&
    !sidebarAutoOpenedRef.value &&
    (comments.value.length > 0 || trackedChanges.value.length > 0)
  ) {
    sidebarAutoOpenedRef.value = true;
    showSidebar.value = true;
  }

  applyInitialScrollTop();
  emit("ready");
});

// Recompute chrome state on every selection / doc / layout event (Requirement 3).
onMounted(() => {
  const offSelection = editor.on("selectionChange", () => {
    stateTick.value++;
    commentLifecycle.updateFloatingButton();
  });
  const offDoc = editor.on("docChange", () => {
    stateTick.value++;
    outlineSidebar.recomputeHeadingsIfOpen();
  });
  const offLayout = editor.on("layoutComplete", () => {
    stateTick.value++;
  });
  // Global mousemove/mouseup listeners for table column/row/edge resize.
  const cleanupTableResize = tableResize.install();
  // Ping the host whenever a (bundled / embedded / system) font finishes loading
  // so it can re-render. Core's fontLoader is a module-level pub/sub shared with
  // React; the font-name array is ignored, matching React's DocxEditor.
  const offFontsLoaded = onFontsLoaded(() => props.onFontsLoaded?.());
  onBeforeUnmount(() => {
    offSelection();
    offDoc();
    offLayout();
    cleanupTableResize();
    offFontsLoaded();
  });
});

// ---- Imperative ref surface ---------------------------------------------
const { exposed } = useDocxEditorRefApi({
  editor,
  editorView,
  // Doc-dirty flag (any doc-changing transaction; cleared on save + load) and
  // comment-list dirty flag together back `hasPendingChanges`, mirroring React's
  // ParagraphChangeTracker signals OR-ed with commentsDirtyRef. The ref API's
  // save wrapper clears the comment flag; useDocxEditor clears the doc flag.
  isDirty,
  commentsDirty: commentManagement.commentsDirty,
  clearCommentsDirty: commentManagement.clearCommentsDirty,
  layout,
  pagesRef,
  pagesViewportRef,
  zoom,
  scrollVisiblePositionIntoView,
  author: () => props.author,
  // Mirror React's applyAIEditOperations comment closure: mint the comment,
  // append it to the thread list, and hand back its id for the tracked-change
  // mark that references it.
  createAIEditComment: (text) => {
    const comment = commentManagement.createComment(text);
    commentManagement.pushComment(comment);
    return comment.id;
  },
  getComments: () => commentManagement.comments.value,
  setComments: commentManagement.setComments,
  focus,
  getDocument,
  setZoom,
  save,
  loadDocument,
  loadDocumentBuffer: loadBuffer,
  onPrint: props.onPrint,
  onSave: props.onSave,
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
/* Coordinate anchor + paint surface for the decoration overlays (anonymization,
   template directives). Covers the content wrapper so its top-left matches the
   projection origin; unscaled (the pages carry the zoom transform, the overlays
   multiply coordinates back by zoom themselves). */
.docx-editor-vue__decoration-origin {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  height: 100%;
  pointer-events: none;
  z-index: 1;
}
@keyframes folio-caret-blink {
  0%,
  45% {
    opacity: 1;
  }
  50%,
  95% {
    opacity: 0;
  }
  100% {
    opacity: 1;
  }
}
.docx-editor-vue__table-insert-btn {
  position: absolute;
  z-index: 6;
  width: 20px;
  height: 20px;
  transform: translate(-50%, -50%);
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0;
  border: 1px solid var(--doc-primary, #1a73e8);
  border-radius: 50%;
  background: var(--doc-surface, #fff);
  color: var(--doc-primary, #1a73e8);
  font-size: 15px;
  line-height: 1;
  cursor: pointer;
}
.docx-editor-vue__table-insert-btn:hover {
  background: var(--doc-primary, #1a73e8);
  color: var(--doc-on-primary, #fff);
}
.docx-editor-vue__add-comment-btn {
  position: absolute;
  z-index: 50;
  transform: translate(-50%, -50%);
  width: 28px;
  height: 28px;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 0;
  border: 1px solid var(--doc-border, #dadce0);
  border-radius: 6px;
  background: var(--doc-page, #fff);
  color: var(--doc-text-muted, #5f6368);
  cursor: pointer;
}
.docx-editor-vue__add-comment-btn:hover {
  background: var(--doc-shadow-subtle, #f1f3f4);
  color: var(--doc-primary, #1a73e8);
}
.docx-editor-vue__error {
  padding: 12px;
  color: var(--destructive, #b00020);
}
.docx-editor-vue__loading,
.docx-editor-vue__placeholder {
  padding: 12px;
  opacity: 0.7;
}
/* Horizontal ruler — sticky at the top of the scroll container, centered over
   the page so it tracks horizontal scrolling. Mirrors React's ruler shell. */
.docx-editor-vue__ruler-h {
  position: sticky;
  top: 0;
  z-index: 20;
  display: flex;
  justify-content: center;
  flex-shrink: 0;
  padding-block: 4px;
  padding-inline: 20px;
  background-color: hsl(var(--muted));
}
/* Vertical ruler — far-left gutter of the editor area; deliberately does not
   track page centering (word-processor gutter convention). */
.docx-editor-vue__ruler-v {
  position: absolute;
  left: 0;
  top: 0;
  z-index: 20;
}
</style>

<!--
  Track-changes display modes. Non-scoped: the `.docx-insertion` /
  `.docx-deletion` spans are painted by ProseMirror's `toDOM` (with inline
  author-color styles), so they never carry this SFC's scope attribute; these
  overrides must be global to win against the inline styles via `!important`.
  Ported verbatim from packages/react/src/styles/editor.css so both adapters
  render each display mode identically.

  All Markup: default — per-author colored underline / strikethrough (inline)
  No Markup: hide change styling, show final result
  Original: hide insertions, show deletions as normal text
  Simple Markup: hide inline marks, show clean text
-->
<style>
/* --- No Markup: show accepted result without change marks --- */
.folio-root--no-markup .docx-insertion {
  color: inherit !important;
  text-decoration: none !important;
  background: none !important;
}
.folio-root--no-markup .docx-deletion {
  display: none !important;
}

/* --- Original: show pre-change state --- */
.folio-root--original .docx-insertion {
  display: none !important;
}
.folio-root--original .docx-deletion {
  color: inherit !important;
  text-decoration: none !important;
  background: none !important;
}

/* --- Simple Markup: clean text (hide inline marks) --- */
.folio-root--simple-markup .docx-insertion {
  color: inherit !important;
  text-decoration: none !important;
  background: none !important;
}
.folio-root--simple-markup .docx-deletion {
  display: none !important;
}
</style>
