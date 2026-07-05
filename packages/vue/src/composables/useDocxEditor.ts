/**
 * useDocxEditor — Vue composable for the folio DOCX editor lifecycle.
 *
 * The load-bearing pipeline of the Vue adapter:
 *   DOCX / Document  →  ProseMirror state  →  layout  →  DOM paint.
 *
 * This is the Vue equivalent of the React `PagedEditor` + `HiddenProseMirror`
 * pair, RE-ENGINEERED against our fork's headless controller (`controller/*`)
 * rather than upstream's `editor` subsystem. The controller owns:
 *   - the off-screen EditorView lifecycle (`createHiddenEditorManager`), and
 *   - the compute-only layout pass (`controller/layoutPipeline#runLayoutPipeline`).
 *
 * This composable is the FRAMEWORK ADAPTER: it supplies the manager its ref
 * getters + callbacks (deciding WHEN to act), assembles the layout-pipeline deps
 * from Vue reactive state on every run, coalesces layout through the shared
 * scheduler, and exposes the reactive surface + a controller handle that
 * `DocxEditor.vue` subscribes to (mirroring how React's PagedEditor exposes its
 * imperative ref).
 *
 * Note on scroll suppression: upstream's Vue port ported a `stripScrollFlag`
 * helper because it created the EditorView itself. Our fork does NOT — the
 * hidden-editor manager owns `dispatchTransaction` and already installs
 * `handleScrollToSelection: suppressHiddenEditorScrollToSelection` on the view,
 * so no local scroll-flag stripping is needed here.
 */

import { onScopeDispose, ref, shallowRef, toValue, watch } from "vue";
import type { MaybeRefOrGetter, Ref } from "vue";

import type { EditorState, Plugin, Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { createFolioEditor } from "@stll/folio-core/controller/folioEditor";
import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import { createFolioEditorEmitter } from "@stll/folio-core/controller/folioEditorEvents";
import {
  createHiddenEditorManager,
  createHiddenEditorState,
} from "@stll/folio-core/controller/hiddenEditorManager";
import type { HiddenEditorManager } from "@stll/folio-core/controller/hiddenEditorManager";
import { runLayoutPipeline as runLayoutPipelineCompute } from "@stll/folio-core/controller/layoutPipeline";
import type { LayoutOutcome, LayoutRunOptions } from "@stll/folio-core/controller/layoutPipeline";
import {
  browserClock,
  createLayoutScheduler,
} from "@stll/folio-core/controller/layoutScheduler";
import type { LayoutScheduler } from "@stll/folio-core/controller/layoutScheduler";
import { createLayoutSession } from "@stll/folio-core/controller/layoutSession";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { getFootnoteText } from "@stll/folio-core/docx/footnoteParser";
import type {
  ConvertHeaderFooterOptions,
  HeaderFooterMetrics,
} from "@stll/folio-core/layout-bridge/convert/headerFooterLayout";
import {
  convertHeaderFooterPmDocToContent,
  convertHeaderFooterToContent,
} from "@stll/folio-core/layout-bridge/convert/headerFooterLayout";
import type { ColumnLayout } from "@stll/folio-core/layout-engine";
import type {
  FootnoteContent,
  HeaderFooterContent,
  Layout,
  PageHeaderFooterRefs,
} from "@stll/folio-core/layout-engine/types";
import { LayoutPainter } from "@stll/folio-core/layout-painter";
import type { FootnoteRenderItem } from "@stll/folio-core/layout-painter/renderPage";
import { LayoutSelectionGate } from "@stll/folio-core/paged-layout/LayoutSelectionGate";
import {
  getMargins,
  getPageSize,
  twipsToPixels,
} from "@stll/folio-core/paged-layout/sectionGeometry";
import { getTransactionDirtyRange } from "@stll/folio-core/paged-layout/transactionDirtyRange";
import { fromProseDoc } from "@stll/folio-core/prosemirror/conversion/fromProseDoc";
import { ExtensionManager } from "@stll/folio-core/prosemirror/extensions/ExtensionManager";
import { createStarterKit } from "@stll/folio-core/prosemirror/extensions/StarterKit";
import type { CommandMap } from "@stll/folio-core/prosemirror/extensions/types";
import {
  createSuggestionModePlugin,
  setSuggestionMode,
} from "@stll/folio-core/prosemirror/plugins/suggestionMode";
import type { TemplatePreviewEntry } from "@stll/folio-core/prosemirror/plugins/templatePreviewValues";
import type { Footnote } from "@stll/folio-core/types/content";
import type {
  Document,
  HeaderFooter,
  SectionProperties,
} from "@stll/folio-core/types/document";
import type { DocxInput } from "@stll/folio-core/utils/docxInput";

// ============================================================================
// CONSTANTS (mirror React PagedEditor.tsx)
// ============================================================================

const DEFAULT_PAGE_GAP = 24;
/** Quiet-window debounce before an interactive layout pass. */
const TRANSACTION_LAYOUT_DEBOUNCE_MS = 32;
/** Upper bound for how long visible layout can trail the hidden editor. */
const TRANSACTION_LAYOUT_MAX_DELAY_MS = 96;
/** Delay before converting PM state back to the Folio document model. */
const DOCUMENT_CHANGE_NOTIFY_DELAY = 250;

// Stable empty fallback for the template-preview entries dep. A fresh `[]` per
// run would defeat the pipeline's identity checks.
const EMPTY_TEMPLATE_PREVIEW_ENTRIES: readonly TemplatePreviewEntry[] = [];

// The layout pipeline threads the hidden header/footer PM handle opaquely. The
// persistent-HF-PM surface (one off-screen EditorView per rId) is not wired in
// this composable yet — see the PORT-BLOCKED note on `syncHfPMs` below — so we
// pass `null`, which makes `renderHfFromContentOrPm` fall back to painting HF
// content straight from the document blocks.
type HfPmsHandle = { getView: (rId: string) => EditorView | null } | null;

// ============================================================================
// HELPERS (ported from React PagedEditor.tsx — pure, framework-neutral)
// ============================================================================

/** Column geometry from `<w:cols>` (twips → px). */
function getColumns(sectionProps: SectionProperties | null | undefined): ColumnLayout | undefined {
  const count = sectionProps?.columnCount ?? 1;
  if (count <= 1) {
    return undefined;
  }
  const gap = twipsToPixels(sectionProps?.columnSpace ?? 720);
  const cols: ColumnLayout = {
    count,
    gap,
    equalWidth: sectionProps?.equalWidth ?? true,
  };
  if (sectionProps?.separator !== undefined) {
    cols.separator = sectionProps.separator;
  }
  return cols;
}

function getHeaderFooterRefsFromSectionProperties(props: SectionProperties): PageHeaderFooterRefs {
  const refs: PageHeaderFooterRefs = {};
  if (props.titlePg !== undefined) {
    refs.titlePg = props.titlePg;
  }
  for (const ref of props.headerReferences ?? []) {
    if (ref.type === "default") {
      refs.headerDefault = ref.rId;
    } else if (ref.type === "first") {
      refs.headerFirst = ref.rId;
    } else {
      refs.headerEven = ref.rId;
    }
  }
  for (const ref of props.footerReferences ?? []) {
    if (ref.type === "default") {
      refs.footerDefault = ref.rId;
    } else if (ref.type === "first") {
      refs.footerFirst = ref.rId;
    } else {
      refs.footerEven = ref.rId;
    }
  }
  return refs;
}

function getSectionHeaderFooterRefs(
  documentModel: Document | null,
): PageHeaderFooterRefs[] | undefined {
  const body = documentModel?.package.document;
  if (!body) {
    return undefined;
  }
  const sections = body.sections;
  if (sections && sections.length > 0) {
    return sections.map((section) => getHeaderFooterRefsFromSectionProperties(section.properties));
  }
  const finalProps = body.finalSectionProperties;
  if (!finalProps) {
    return undefined;
  }
  return [getHeaderFooterRefsFromSectionProperties(finalProps)];
}

/** One resolved header/footer slot: its `HeaderFooter` and relationship id. */
type ResolvedHfSlot = { hf: HeaderFooter | null; rId: string | null };
type ResolvedHeaderFooters = {
  header: ResolvedHfSlot;
  footer: ResolvedHfSlot;
  firstHeader: ResolvedHfSlot;
  firstFooter: ResolvedHfSlot;
};

const EMPTY_HF_SLOT: ResolvedHfSlot = { hf: null, rId: null };

/**
 * Resolve the header/footer parts a section references into concrete
 * `HeaderFooter` objects + rIds. Minimal single-section resolution: the
 * default part (all pages, or pages 2+ under titlePg) and the first-page part.
 * The full HF-editing surface (persistent PMs, even/odd, per-section swaps)
 * lives in React's `useHeaderFooterEditor` and is not ported here.
 */
function resolveHeaderFooters(
  documentModel: Document | null,
  sectionProps: SectionProperties | null,
): ResolvedHeaderFooters {
  const pkg = documentModel?.package;
  if (!pkg || !sectionProps) {
    return {
      header: EMPTY_HF_SLOT,
      footer: EMPTY_HF_SLOT,
      firstHeader: EMPTY_HF_SLOT,
      firstFooter: EMPTY_HF_SLOT,
    };
  }
  const headers = pkg.headers;
  const footers = pkg.footers;
  const lookupHeader = (rId: string | null): HeaderFooter | null =>
    rId ? headers?.get(rId) ?? null : null;
  const lookupFooter = (rId: string | null): HeaderFooter | null =>
    rId ? footers?.get(rId) ?? null : null;

  let headerRId: string | null = null;
  let firstHeaderRId: string | null = null;
  for (const ref of sectionProps.headerReferences ?? []) {
    if (ref.type === "default") {
      headerRId = ref.rId;
    } else if (ref.type === "first") {
      firstHeaderRId = ref.rId;
    }
  }
  let footerRId: string | null = null;
  let firstFooterRId: string | null = null;
  for (const ref of sectionProps.footerReferences ?? []) {
    if (ref.type === "default") {
      footerRId = ref.rId;
    } else if (ref.type === "first") {
      firstFooterRId = ref.rId;
    }
  }

  return {
    header: { hf: lookupHeader(headerRId), rId: headerRId },
    footer: { hf: lookupFooter(footerRId), rId: footerRId },
    firstHeader: { hf: lookupHeader(firstHeaderRId), rId: firstHeaderRId },
    firstFooter: { hf: lookupFooter(firstFooterRId), rId: firstFooterRId },
  };
}

/**
 * Source HeaderFooterContent for the painter from a persistent hidden HF
 * EditorView when one exists (keeps the painter in lockstep with live PM
 * edits), otherwise straight from the HeaderFooter document blocks.
 */
function renderHfFromContentOrPm(
  hf: HeaderFooter | null | undefined,
  rId: string | null | undefined,
  hfPMs: HfPmsHandle,
  contentWidth: number,
  metrics: HeaderFooterMetrics,
  options: ConvertHeaderFooterOptions,
): HeaderFooterContent | undefined {
  if (!hf) {
    return undefined;
  }
  const optsWithRId: ConvertHeaderFooterOptions = rId ? { ...options, rId } : options;
  const view = rId ? hfPMs?.getView(rId) : null;
  if (view) {
    return convertHeaderFooterPmDocToContent(view.state.doc, contentWidth, metrics, optsWithRId);
  }
  return convertHeaderFooterToContent(hf, contentWidth, metrics, optsWithRId);
}

function renderHeaderFooterContentByRId(
  source: Map<string, HeaderFooter> | undefined,
  hfPMs: HfPmsHandle,
  contentWidth: number,
  metrics: HeaderFooterMetrics,
  options: ConvertHeaderFooterOptions,
): Map<string, HeaderFooterContent> | undefined {
  if (!source || source.size === 0) {
    return undefined;
  }
  const rendered = new Map<string, HeaderFooterContent>();
  for (const [rId, hf] of source) {
    const content = renderHfFromContentOrPm(hf, rId, hfPMs, contentWidth, metrics, options);
    if (content) {
      rendered.set(rId, content);
    }
  }
  return rendered.size > 0 ? rendered : undefined;
}

/** Build per-page footnote render items from the page → footnote-id mapping. */
function buildFootnoteRenderItems(
  pageFootnoteMap: Map<number, number[]>,
  footnoteContentMap: Map<number, FootnoteContent>,
  doc: Document | null,
): Map<number, FootnoteRenderItem[]> {
  const result = new Map<number, FootnoteRenderItem[]>();
  if (!doc || !doc.package.footnotes) {
    return result;
  }

  const fnLookup = new Map<number, Footnote>();
  for (const fn of doc.package.footnotes) {
    if (fn.noteType && fn.noteType !== "normal") {
      continue;
    }
    fnLookup.set(fn.id, fn);
  }

  for (const [pageNumber, footnoteIds] of pageFootnoteMap) {
    const items: FootnoteRenderItem[] = [];
    for (const fnId of footnoteIds) {
      const fn = fnLookup.get(fnId);
      if (!fn) {
        continue;
      }
      const content = footnoteContentMap.get(fnId);
      const displayNum = content?.displayNumber ?? 0;
      items.push({
        displayNumber: String(displayNum),
        text: getFootnoteText(fn),
        ...(content
          ? {
              content: {
                blocks: content.blocks,
                measures: content.measures,
                height: content.height,
              },
            }
          : {}),
      });
    }
    if (items.length > 0) {
      result.set(pageNumber, items);
    }
  }
  return result;
}

const VALID_HIGHLIGHT_COLORS = new Set([
  "black",
  "blue",
  "cyan",
  "darkBlue",
  "darkCyan",
  "darkGray",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "green",
  "lightGray",
  "magenta",
  "none",
  "red",
  "white",
  "yellow",
]);

/** Diagnostic string listing highlight marks the painter would reject. */
function describeInvalidHighlightMarks(doc: EditorState["doc"]): string {
  const invalidHighlights: string[] = [];
  const visit = (node: EditorState["doc"], path: string): void => {
    for (const [index, mark] of node.marks.entries()) {
      if (
        mark.type.name === "highlight" &&
        !VALID_HIGHLIGHT_COLORS.has(String(mark.attrs["color"]))
      ) {
        invalidHighlights.push(`${path}.marks[${index}]=${JSON.stringify(mark.attrs)}`);
      }
    }
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    node.forEach((child, _offset, index) => {
      visit(child, `${path}.content[${index}]`);
    });
  };
  visit(doc, "doc");
  return invalidHighlights.join("; ");
}

/** Whether the browser's document FontFaceSet has settled. */
function documentFontsAreLoaded(): boolean {
  if (typeof document === "undefined" || !("fonts" in document)) {
    return true;
  }
  return document.fonts.status === "loaded";
}

// ============================================================================
// COMPOSABLE
// ============================================================================

export interface UseDocxEditorOptions {
  /** Container element that hosts the off-screen ProseMirror editor. */
  hiddenContainer: Ref<HTMLElement | null>;
  /** Container element the paginated pages are painted into. */
  pagesContainer: Ref<HTMLElement | null>;
  /** Whether the editor is read-only. Reactive. */
  readOnly?: MaybeRefOrGetter<boolean>;
  /** Gap between pages in pixels. */
  pageGap?: number;
  /**
   * Stable identity of the loaded document (same across internal edits, distinct
   * per loaded file). Threaded to the hidden-editor manager's external-load
   * detection. Reactive.
   */
  documentKey?: MaybeRefOrGetter<string | undefined>;
  /**
   * Editor mode. `'suggesting'` activates the mounted suggestion-mode plugin so
   * typed text becomes tracked changes. Reactive — flip at runtime.
   */
  editorMode?: MaybeRefOrGetter<"editing" | "suggesting" | "viewing">;
  /** Author name attached to tracked changes minted in suggesting mode. */
  author?: MaybeRefOrGetter<string>;
  /** External ProseMirror plugins supplied by the host app. */
  externalPlugins?: Plugin[];
  /** Callback on document change (debounced PM → Document model conversion). */
  onChange?: (doc: Document) => void;
  /** Callback on parse / layout / lifecycle error. */
  onError?: (error: Error) => void;
  /** Callback on every selection-bearing transaction. */
  onSelectionUpdate?: (state: EditorState) => void;
  /** Callback when the hidden EditorView is created or torn down. */
  onEditorViewReady?: (view: EditorView | null) => void;
  /** Callback when a read-only user action would mutate the document. */
  onReadOnlyEditAttempt?: () => void;
}

export interface UseDocxEditorReturn {
  /** The headless controller (imperative API + layout access + events). */
  editor: FolioEditor;
  /** Off-screen ProseMirror EditorView, or null before mount. */
  editorView: Ref<EditorView | null>;
  /** Latest editor state, updated on each transaction. */
  editorState: Ref<EditorState | null>;
  /** True once the hidden view is mounted and a document is loaded. */
  isReady: Ref<boolean>;
  /** Last parse error message, or null if the most recent load succeeded. */
  parseError: Ref<string | null>;
  /** Computed page layout, or null before first paint. */
  layout: Ref<Layout | null>;
  /** Load a DOCX from a binary buffer / Blob / File. */
  loadBuffer: (buffer: DocxInput) => Promise<void>;
  /** Load an already-parsed `Document` directly. */
  loadDocument: (doc: Document) => void;
  /** Serialize the current document to a DOCX blob. */
  save: () => Promise<Blob | null>;
  /** Snapshot the current document model. */
  getDocument: () => Document | null;
  /** Publish a fresh Document object (e.g. after HF materialisation). */
  setDocument: (doc: Document) => void;
  /** Access the extension command map for invoking marks / nodes / features. */
  getCommands: () => CommandMap;
  /** Focus the hidden ProseMirror view. */
  focus: () => void;
  /** Force a re-layout without a doc change (e.g. after page-setup edits). */
  reLayout: () => void;
  /** Destroy the editor view and clean up listeners. */
  destroy: () => void;
}

export function useDocxEditor(options: UseDocxEditorOptions): UseDocxEditorReturn {
  const {
    hiddenContainer,
    pagesContainer,
    readOnly = false,
    pageGap = DEFAULT_PAGE_GAP,
    documentKey,
    editorMode,
    author,
    externalPlugins = [],
    onChange,
    onError,
    onSelectionUpdate,
    onEditorViewReady,
    onReadOnlyEditAttempt,
  } = options;

  // ---- Reactive state -----------------------------------------------------
  // `docModel` (not `document`) so the global `document` stays reachable for the
  // FontFaceSet probe + DOM host lookups.
  const docModel = shallowRef<Document | null>(null);
  const editorView = shallowRef<EditorView | null>(null);
  const editorState = shallowRef<EditorState | null>(null);
  const layout = shallowRef<Layout | null>(null);
  const isReady = ref(false);
  const parseError = ref<string | null>(null);

  // ---- Long-lived controller singletons -----------------------------------
  // One ExtensionManager owns the schema + plugins + commands for the body view.
  const extensionManager = new ExtensionManager(createStarterKit());
  extensionManager.buildSchema();
  extensionManager.initializeRuntime();

  // Suggestion-mode plugin registered inactive; `setSuggestionMode` toggles it
  // via plugin meta. Reused across state rebuilds (PM plugins are stateless
  // definitions). Prepended to the host's external plugins so the hidden-editor
  // manager installs it on every (re)created state.
  const suggestionPlugin = createSuggestionModePlugin(false);

  const emitter = createFolioEditorEmitter();
  const syncCoordinator = new LayoutSelectionGate();
  const session = createLayoutSession();
  const painter = new LayoutPainter({ pageGap, showShadow: true });

  // ---- Layout pipeline ----------------------------------------------------

  function applyOutcome(outcome: LayoutOutcome): void {
    if (outcome.layout) {
      layout.value = outcome.layout;
    }
    if (outcome.blockLookup) {
      painter.setBlockLookup(outcome.blockLookup);
    }
    if (outcome.layout) {
      emitter.emit("layoutComplete", outcome.layout);
    }
  }

  /**
   * Assemble the compute deps from current reactive state and run the shared
   * layout pipeline (blocks → measure → layout → paint). All values are read
   * fresh here, so there is no stale-closure risk (the pipeline treats them as
   * plain inputs, matching the React adapter).
   */
  function runLayoutPipeline(state: EditorState, runOptions: LayoutRunOptions = {}): void {
    const container = pagesContainer.value;
    // The pipeline paints into an HTMLDivElement; narrow without a cast so a
    // non-div host (or a not-yet-mounted ref) simply computes without painting.
    const paintTarget = container instanceof HTMLDivElement ? container : null;

    const model = docModel.value;
    const body = model?.package.document;
    // Lead geometry from the first section; trailing section falls back to the
    // body's final section properties. Mirrors the React split.
    const sectionProps =
      body?.sections?.[0]?.properties ?? body?.finalSectionProperties ?? null;

    const pageSize = getPageSize(sectionProps);
    const margins = getMargins(sectionProps);
    const columns = getColumns(sectionProps);
    const contentWidth = pageSize.w - margins.left - margins.right;
    const styles = model?.package.styles ?? null;
    const theme = model?.package.theme ?? null;
    const defaultTabStop = model?.package.settings?.defaultTabStop;
    const hf = resolveHeaderFooters(model, sectionProps);

    try {
      const outcome = runLayoutPipelineCompute<HfPmsHandle>(
        {
          contentWidth,
          columns,
          pageSize,
          margins,
          pageGap,
          syncCoordinator,
          headerContent: hf.header.hf,
          footerContent: hf.footer.hf,
          firstPageHeaderContent: hf.firstHeader.hf,
          firstPageFooterContent: hf.firstFooter.hf,
          headerContentRId: hf.header.rId,
          footerContentRId: hf.footer.rId,
          firstPageHeaderContentRId: hf.firstHeader.rId,
          firstPageFooterContentRId: hf.firstFooter.rId,
          sectionHeaderFooterRefs: getSectionHeaderFooterRefs(model),
          theme,
          sectionProperties: sectionProps,
          document: model,
          defaultTabStop,
          styles,
          layout: layout.value,
          hfPMs: null,
          painter,
          pagesContainer: paintTarget,
          session,
          renderHfFromContentOrPm,
          renderHeaderFooterContentByRId,
          documentFontsAreLoaded,
          buildFootnoteRenderItems,
          describeInvalidHighlightMarks,
          emptyTemplatePreviewEntries: EMPTY_TEMPLATE_PREVIEW_ENTRIES,
        },
        state,
        runOptions,
      );
      applyOutcome(outcome);
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      onError?.(error);
    }
  }

  // rAF-coalescing scheduler (shared with React via core). A burst of
  // keystrokes lays out once per frame instead of synchronously per keystroke.
  const scheduler: LayoutScheduler<EditorState> = createLayoutScheduler<EditorState>({
    runLayout: (state, runOptions) => runLayoutPipeline(state, runOptions),
    debounceMs: TRANSACTION_LAYOUT_DEBOUNCE_MS,
    maxDelayMs: TRANSACTION_LAYOUT_MAX_DELAY_MS,
    clock: browserClock,
  });

  // ---- Debounced PM → Document writeback ----------------------------------

  let docChangeTimer: number | null = null;

  function flushDocumentChangeNotification(): void {
    if (docChangeTimer !== null) {
      window.clearTimeout(docChangeTimer);
      docChangeTimer = null;
    }
    const view = editorView.value;
    const base = docModel.value;
    if (!view || !base) {
      return;
    }
    try {
      const updated = fromProseDoc(view.state.doc, base);
      docModel.value = updated;
      onChange?.(updated);
      emitter.emit("docChange", updated);
    } catch (err) {
      onError?.(err instanceof Error ? err : new Error(String(err)));
    }
  }

  function scheduleDocumentChangeNotification(): void {
    if (docChangeTimer !== null) {
      window.clearTimeout(docChangeTimer);
    }
    docChangeTimer = window.setTimeout(() => {
      docChangeTimer = null;
      flushDocumentChangeNotification();
    }, DOCUMENT_CHANGE_NOTIFY_DELAY);
  }

  // ---- Hidden-editor manager ----------------------------------------------
  // The off-screen EditorView lifecycle (create/destroy, editorProps, external-
  // document sync) lives in the framework-agnostic manager. This composable
  // supplies the ref getters + callbacks and drives it through its methods —
  // the Vue equivalent of React's HiddenProseMirror.

  const manager: HiddenEditorManager = createHiddenEditorManager({
    getHost: () => hiddenContainer.value,
    getDocument: () => docModel.value,
    getStyles: () => docModel.value?.package.styles ?? null,
    getExtensionManager: () => extensionManager,
    getExternalPlugins: () => [suggestionPlugin, ...externalPlugins],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: () => toValue(readOnly),
    getDocumentKey: () => toValue(documentKey),
    getDocumentContext: () => docModel.value,
    onTransaction: handleTransaction,
    onSelectionChange: (state) => {
      editorState.value = state;
      onSelectionUpdate?.(state);
      emitter.emit("selectionChange", {
        from: state.selection.from,
        to: state.selection.to,
      });
    },
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => onReadOnlyEditAttempt?.(),
    onEditorViewReady: handleEditorViewReady,
    onEditorViewDestroy: () => {
      editorView.value = null;
      isReady.value = false;
      onEditorViewReady?.(null);
    },
    onRemoteSelectionsChange: () => {
      // Collaboration remote selections are not wired in the Vue adapter yet.
    },
  });

  function handleTransaction(transaction: Transaction, newState: EditorState): void {
    editorState.value = newState;
    if (transaction.docChanged) {
      syncCoordinator.incrementStateSeq();
      scheduler.schedule(newState, getTransactionDirtyRange(transaction));
      scheduleDocumentChangeNotification();
    }
    syncCoordinator.requestRender();
  }

  function handleEditorViewReady(view: EditorView): void {
    editorView.value = view;
    editorState.value = view.state;
    isReady.value = true;

    // Initial layout for the freshly mounted view.
    runLayoutPipeline(view.state, { reason: "initial" });
    syncCoordinator.requestRender();

    // Apply the current editor mode to the mounted suggestion plugin.
    syncSuggestionMode(view);

    onEditorViewReady?.(view);

    // Auto-focus so the user can type immediately (unless read-only). rAF lets
    // the first paint settle first. Mirrors PagedEditor.handleEditorViewReady.
    if (!toValue(readOnly)) {
      requestAnimationFrame(() => {
        if (editorView.value === view) {
          view.focus();
        }
      });
    }
  }

  // ---- Headless controller handle -----------------------------------------

  const editor: FolioEditor = createFolioEditor({
    getEditorApi: () => manager.api,
    getLayout: () => layout.value,
    runLayout: (state, runOptions) => runLayoutPipeline(state, runOptions),
    emitter,
  });

  // ---- Suggestion-mode sync -----------------------------------------------

  function syncSuggestionMode(view: EditorView): void {
    const active = toValue(editorMode) === "suggesting";
    setSuggestionMode(active, view.state, view.dispatch, toValue(author));
  }

  watch(
    [() => toValue(editorMode), () => toValue(author)],
    () => {
      const view = editorView.value;
      if (view) {
        syncSuggestionMode(view);
      }
    },
  );

  // Keep the hidden view's editable() ARIA state in sync with readOnly.
  watch(
    () => toValue(readOnly),
    () => manager.syncEditable(),
  );

  // ---- View lifecycle -----------------------------------------------------

  function mountView(): void {
    // Fresh boot: request creation once the host is present. The manager gates
    // creation on the host, so this is safe to call before the ref resolves.
    manager.ensureView();
  }

  function remountForNewDocument(): void {
    // A new document is a truly external swap. Rather than rely on the manager's
    // metadata-signature heuristic (which can skip a same-metadata reload), tear
    // the view down and rebuild it so the seed state + initial layout run fresh.
    scheduler.dispose();
    manager.destroyView();
    manager.ensureView();
    if (!manager.getView()) {
      // Host not mounted yet: paint from a precomputed state so the pages show
      // before first interaction, then let mountView create the real view.
      paintFromPrecomputedState();
    }
  }

  /**
   * Paint an initial layout from a precomputed hidden-editor state before the
   * EditorView mounts, so the document is visible on cold boot without waiting
   * for the host element. Mirrors React's pre-hidden initial-layout effect.
   */
  function paintFromPrecomputedState(): void {
    const model = docModel.value;
    if (!model) {
      return;
    }
    try {
      const initialState = createHiddenEditorState({
        document: model,
        styles: model.package.styles ?? null,
        manager: extensionManager,
        externalPlugins: [suggestionPlugin, ...externalPlugins],
        collaborationModules: null,
        reason: "mount",
      });
      editorState.value = initialState;
      runLayoutPipeline(initialState, { reason: "initial" });
    } catch (err) {
      onError?.(err instanceof Error ? err : new Error(String(err)));
    }
  }

  // Recreate / repaint whenever the host mounts or the document identity flips.
  watch(
    [hiddenContainer, pagesContainer],
    () => {
      if (hiddenContainer.value && docModel.value && !manager.getView()) {
        mountView();
      }
    },
    { flush: "post" },
  );

  // ---- Public loading API -------------------------------------------------

  async function loadBuffer(buffer: DocxInput): Promise<void> {
    parseError.value = null;
    isReady.value = false;
    try {
      const doc = await parseDocx(buffer);
      docModel.value = doc;
      remountForNewDocument();
    } catch (err) {
      const error = err instanceof Error ? err : new Error(String(err));
      parseError.value = error.message;
      onError?.(error);
    }
  }

  function loadDocument(doc: Document): void {
    parseError.value = null;
    docModel.value = doc;
    remountForNewDocument();
  }

  // ---- Public API ---------------------------------------------------------

  async function save(): Promise<Blob | null> {
    const view = editorView.value;
    const base = docModel.value;
    if (!view || !base) {
      return null;
    }
    const { repackDocx, createDocx } = await import("@stll/folio-core/docx/rezip");
    const updatedDoc = fromProseDoc(view.state.doc, base);
    const buffer = updatedDoc.originalBuffer
      ? await repackDocx(updatedDoc)
      : await createDocx(updatedDoc);
    return new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
  }

  function getDocument(): Document | null {
    return docModel.value;
  }

  function setDocument(doc: Document): void {
    docModel.value = doc;
  }

  function getCommands(): CommandMap {
    return extensionManager.getCommands();
  }

  function focus(): void {
    editorView.value?.focus();
  }

  function reLayout(): void {
    const view = editorView.value;
    if (view) {
      runLayoutPipeline(view.state, { reason: "manual" });
    }
  }

  function destroy(): void {
    scheduler.dispose();
    if (docChangeTimer !== null) {
      window.clearTimeout(docChangeTimer);
      docChangeTimer = null;
    }
    manager.destroyView();
    extensionManager.destroy();
    editorState.value = null;
    layout.value = null;
    docModel.value = null;
    isReady.value = false;
  }

  onScopeDispose(destroy);

  return {
    editor,
    editorView,
    editorState,
    isReady,
    parseError,
    layout,
    loadBuffer,
    loadDocument,
    save,
    getDocument,
    setDocument,
    getCommands,
    focus,
    reLayout,
    destroy,
  };
}
