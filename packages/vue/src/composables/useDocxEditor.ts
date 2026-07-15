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
import { loadCollaborationModules } from "@stll/folio-core/controller/collaborationModules";
import {
  collectRemoteSelections,
  createHiddenEditorManager,
  createHiddenEditorState,
} from "@stll/folio-core/controller/hiddenEditorManager";
import type {
  CollaborationModules,
  HiddenEditorManager,
  HiddenProseMirrorCollaboration,
  HiddenProseMirrorRemoteSelection,
} from "@stll/folio-core/controller/hiddenEditorManager";
import { runLayoutPipeline as runLayoutPipelineCompute } from "@stll/folio-core/controller/layoutPipeline";
import type { LayoutOutcome, LayoutRunOptions } from "@stll/folio-core/controller/layoutPipeline";
import { browserClock, createLayoutScheduler } from "@stll/folio-core/controller/layoutScheduler";
import type { LayoutScheduler } from "@stll/folio-core/controller/layoutScheduler";
import { createLayoutSession } from "@stll/folio-core/controller/layoutSession";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { getFootnoteText } from "@stll/folio-core/docx/footnoteParser";
import type { FolioSelectiveSaveFlags } from "@stll/folio-core/docx/selectiveSaveFlags";
import type { TripwireResult } from "@stll/folio-core/docx/selectiveSaveTripwire";
import type {
  ConvertHeaderFooterOptions,
  HeaderFooterMetrics,
} from "@stll/folio-core/layout-bridge/convert/headerFooterLayout";
import {
  convertHeaderFooterPmDocToContent,
  convertHeaderFooterToContent,
} from "@stll/folio-core/layout-bridge/convert/headerFooterLayout";
import { resolveSectionHeaderFooterRefs, type ColumnLayout } from "@stll/folio-core/layout-engine";
import type {
  FlowBlock,
  FootnoteContent,
  HeaderFooterContent,
  Layout,
  Measure,
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
import {
  getChangedParagraphIds,
  hasStructuralChanges,
  hasUntrackedChanges,
} from "@stll/folio-core/prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { createStarterKit } from "@stll/folio-core/prosemirror/extensions/StarterKit";
import type { CommandMap } from "@stll/folio-core/prosemirror/extensions/types";
import {
  createSuggestionModePlugin,
  setSuggestionMode,
} from "@stll/folio-core/prosemirror/plugins/suggestionMode";
import type { AnonymizationMatch } from "@stll/folio-core/prosemirror/plugins/anonymizationDecorations";
import { createAnonymizationDecorationsPlugin } from "@stll/folio-core/prosemirror/plugins/anonymizationDecorations";
import { createTemplateDirectivesPlugin } from "@stll/folio-core/prosemirror/plugins/templateDirectives";
import { createTemplatePreviewValuesPlugin } from "@stll/folio-core/prosemirror/plugins/templatePreviewValues";
import type { TemplatePreviewEntry } from "@stll/folio-core/prosemirror/plugins/templatePreviewValues";
import type {
  TemplateSlashMenuKeyAction,
  TemplateSlashMenuState,
} from "@stll/folio-core/prosemirror/plugins/templateSlashMenu";
import { templateSlashMenuPlugin } from "@stll/folio-core/prosemirror/plugins/templateSlashMenu";
import type { Footnote } from "@stll/folio-core/types/content";
import type { Document, HeaderFooter, SectionProperties } from "@stll/folio-core/types/document";
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
    rId ? (headers?.get(rId) ?? null) : null;
  const lookupFooter = (rId: string | null): HeaderFooter | null =>
    rId ? (footers?.get(rId) ?? null) : null;

  let headerRId: string | null = null;
  let firstHeaderRId: string | null = null;
  for (const hfRef of sectionProps.headerReferences ?? []) {
    if (hfRef.type === "default") {
      headerRId = hfRef.rId;
    } else if (hfRef.type === "first") {
      firstHeaderRId = hfRef.rId;
    }
  }
  let footerRId: string | null = null;
  let firstFooterRId: string | null = null;
  for (const hfRef of sectionProps.footerReferences ?? []) {
    if (hfRef.type === "default") {
      footerRId = hfRef.rId;
    } else if (hfRef.type === "first") {
      firstFooterRId = hfRef.rId;
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

export type UseDocxEditorOptions = {
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
  /** Password for Agile-encrypted .docx files (Office 2010+). Reactive. */
  password?: MaybeRefOrGetter<string | undefined>;
  /**
   * Editor mode. `'suggesting'` activates the mounted suggestion-mode plugin so
   * typed text becomes tracked changes. Reactive — flip at runtime.
   */
  editorMode?: MaybeRefOrGetter<"editing" | "suggesting" | "viewing">;
  /** Author name attached to tracked changes minted in suggesting mode. */
  author?: MaybeRefOrGetter<string>;
  /** External ProseMirror plugins supplied by the host app. */
  externalPlugins?: Plugin[];
  /** Reactive Yjs collaboration owner and the ProseMirror binding plugins. */
  collaboration?: MaybeRefOrGetter<UseDocxEditorCollaboration | undefined>;
  /**
   * Fires with the anonymization decoration plugin's current match list on
   * mount, term push, and every doc edit. The anonymization plugin is always
   * installed (inert until terms are pushed via `setAnonymizationTermsMeta`),
   * mirroring the React adapter.
   */
  onAnonymizationMatchesChange?: (matches: readonly AnonymizationMatch[]) => void;
  /**
   * Install the template-directive + slash-menu plugins so `{{...}}` markers are
   * scanned and the `/` trigger is active. Reactive: toggling reconfigures the
   * live view's plugin set. Off for ordinary documents; on for the template
   * editor. Mirrors React's `showTemplateDirectives` plugin gate.
   */
  showTemplateDirectives?: MaybeRefOrGetter<boolean | undefined>;
  /** Fires when the slash-menu trigger opens, its query changes, or it closes. */
  onSlashMenuChange?: (state: TemplateSlashMenuState) => void;
  /** Resolves a navigation/commit key while the slash menu is open. */
  onSlashMenuKeyAction?: (action: TemplateSlashMenuKeyAction) => boolean;
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
  /**
   * Operational flags for the save path. Selective save and its tripwire mode
   * are off by default; hosts opt in once their rollout pipeline is ready.
   * Reactive — read fresh on each `save()`. Mirrors React's `featureFlags`.
   */
  featureFlags?: MaybeRefOrGetter<FolioSelectiveSaveFlags | undefined>;
  /**
   * Fired with the structured selective-vs-full comparison after a save when
   * {@link FolioSelectiveSaveFlags.selectiveSaveTripwire} is on. Observability
   * only — never blocks or poisons the save. Mirrors React's callback.
   */
  onSelectiveSaveTripwire?: ((result: TripwireResult) => void) | undefined;
};

export type UseDocxEditorCollaboration = HiddenProseMirrorCollaboration & {
  plugins?: Plugin[] | undefined;
};

export type UseDocxEditorReturn = {
  /** The headless controller (imperative API + layout access + events). */
  editor: FolioEditor;
  /** Off-screen ProseMirror EditorView, or null before mount. */
  editorView: Ref<EditorView | null>;
  /** Latest editor state, updated on each transaction. */
  editorState: Ref<EditorState | null>;
  /** Remote collaborative cursors decoded into body ProseMirror positions. */
  remoteSelections: Ref<HiddenProseMirrorRemoteSelection[]>;
  /** True once the hidden view is mounted and a document is loaded. */
  isReady: Ref<boolean>;
  /**
   * True when the live PM state has doc edits not yet serialized by `save()`.
   * Set on every doc-changing transaction, cleared on a successful save and on
   * document load/swap. Backs the ref API's `hasPendingChanges` (OR-ed with the
   * comment-list dirty flag).
   */
  isDirty: Ref<boolean>;
  /** Last parse error message, or null if the most recent load succeeded. */
  parseError: Ref<string | null>;
  /** Computed page layout, or null before first paint. */
  layout: Ref<Layout | null>;
  /** Flow blocks from the latest layout pass — the range-projection fallback. */
  blocks: Ref<FlowBlock[]>;
  /** Measures from the latest layout pass — the range-projection fallback. */
  measures: Ref<Measure[]>;
  /** Gate that schedules overlay work only after the matching layout paint. */
  syncCoordinator: LayoutSelectionGate;
  /** Load a DOCX from a binary buffer / Blob / File. */
  loadBuffer: (buffer: DocxInput) => Promise<void>;
  /** Load an already-parsed `Document` directly. */
  loadDocument: (doc: Document) => void;
  /** Serialize the current document to a DOCX blob. */
  save: (options?: { selective?: boolean }) => Promise<Blob | null>;
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
};

export function useDocxEditor(options: UseDocxEditorOptions): UseDocxEditorReturn {
  const {
    hiddenContainer,
    pagesContainer,
    readOnly = false,
    pageGap = DEFAULT_PAGE_GAP,
    documentKey,
    password,
    editorMode,
    author,
    externalPlugins = [],
    collaboration,
    onAnonymizationMatchesChange,
    showTemplateDirectives,
    onSlashMenuChange,
    onSlashMenuKeyAction,
    onChange,
    onError,
    onSelectionUpdate,
    onEditorViewReady,
    onReadOnlyEditAttempt,
    featureFlags,
    onSelectiveSaveTripwire,
  } = options;

  // ---- Reactive state -----------------------------------------------------
  // `docModel` (not `document`) so the global `document` stays reachable for the
  // FontFaceSet probe + DOM host lookups.
  const docModel = shallowRef<Document | null>(null);
  const editorView = shallowRef<EditorView | null>(null);
  const editorState = shallowRef<EditorState | null>(null);
  const collaborationModules = shallowRef<CollaborationModules | null>(null);
  const remoteSelections = shallowRef<HiddenProseMirrorRemoteSelection[]>([]);
  const layout = shallowRef<Layout | null>(null);
  // Latest flow blocks + measures — the range-projection fallback (off-screen /
  // not-yet-painted ranges) needs them; the primary DOM-rect path does not.
  const blocks = shallowRef<FlowBlock[]>([]);
  const measures = shallowRef<Measure[]>([]);
  const isReady = ref(false);
  const parseError = ref<string | null>(null);
  // "Live PM state has edits not yet serialized" flag, the Vue analogue of the
  // signals React derives from the ParagraphChangeTracker in `hasPendingChanges`.
  // Set on every doc-changing transaction (the same signal that drives the
  // debounced writeback/`onChange`) and cleared on a successful `save()` and on
  // document load/swap. Vue keeps the change tracker uncleared across saves (it
  // is the selective-save baseline; see the featureFlags parity note), so this
  // flag — not the tracker — is what surfaces pending edits to the ref API.
  const isDirty = ref(false);

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

  // Anonymization-term highlights. Always installed (inert until the host pushes
  // terms via `setAnonymizationTermsMeta`), matching React. The callback is read
  // through the stable closure below so the plugin identity never changes.
  const anonymizationPlugin = createAnonymizationDecorationsPlugin({
    onMatchesChange: (matches) => onAnonymizationMatchesChange?.(matches),
  });
  // Template-directive scan + slash-menu trigger. Installed only when
  // `showTemplateDirectives` is on (see `buildExternalPlugins` /
  // `syncTemplatePlugins`). Preview plugin is inert until values are pushed.
  const templateDirectivesPlugin = createTemplateDirectivesPlugin();
  const templateSlashMenu = templateSlashMenuPlugin({
    onChange: (state) => onSlashMenuChange?.(state),
    onKeyAction: (action) => onSlashMenuKeyAction?.(action) ?? false,
  });
  const templatePreviewPlugin = createTemplatePreviewValuesPlugin();

  const templatePluginsEnabled = (): boolean => toValue(showTemplateDirectives) === true;

  // Plugin list installed on every (re)created hidden-editor state. Anonymization
  // + preview are always present; the template directive/slash-menu pair is gated
  // on `showTemplateDirectives`. The host's `externalPlugins` lead, matching the
  // React adapter's ordering (host plugins before feature decorations).
  function buildExternalPlugins(): Plugin[] {
    return [
      suggestionPlugin,
      ...externalPlugins,
      ...(toValue(collaboration)?.plugins ?? []),
      anonymizationPlugin,
      ...(templatePluginsEnabled() ? [templateDirectivesPlugin, templateSlashMenu] : []),
      templatePreviewPlugin,
    ];
  }

  const emitter = createFolioEditorEmitter();
  const syncCoordinator = new LayoutSelectionGate();
  const session = createLayoutSession();
  const painter = new LayoutPainter({ pageGap, showShadow: true });

  // ---- Layout pipeline ----------------------------------------------------

  function applyOutcome(outcome: LayoutOutcome): void {
    if (outcome.layout) {
      layout.value = outcome.layout;
    }
    if (outcome.blocks) {
      blocks.value = outcome.blocks;
    }
    if (outcome.measures) {
      measures.value = outcome.measures;
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
    const sectionProps = body?.sections?.[0]?.properties ?? body?.finalSectionProperties ?? null;

    const pageSize = getPageSize(sectionProps);
    const margins = getMargins(sectionProps);
    const columns = getColumns(sectionProps);
    const contentWidth = pageSize.w - margins.left - margins.right;
    const styles = model?.package.styles ?? null;
    const theme = model?.package.theme ?? null;
    const defaultTabStop = model?.package.settings?.defaultTabStop;
    const documentSettings = model?.package.settings;
    const mirrorMargins =
      documentSettings !== undefined &&
      "mirrorMargins" in documentSettings &&
      documentSettings.mirrorMargins === true;
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
          sectionHeaderFooterRefs: resolveSectionHeaderFooterRefs(model),
          theme,
          sectionProperties: sectionProps,
          document: model,
          defaultTabStop,
          mirrorMargins,
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
    getExternalPlugins: buildExternalPlugins,
    getCollaboration: () => toValue(collaboration),
    getCollaborationModules: () => collaborationModules.value,
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
    onRemoteSelectionsChange: (selections) => {
      remoteSelections.value = selections;
    },
  });

  const publishRemoteSelections = (): void => {
    const awareness = toValue(collaboration)?.awareness;
    const modules = collaborationModules.value;
    const view = manager.getView();
    remoteSelections.value =
      awareness && modules && view ? collectRemoteSelections(view.state, awareness, modules) : [];
  };

  watch(
    () => toValue(collaboration),
    (activeCollaboration, _previousCollaboration, onCleanup) => {
      remoteSelections.value = [];
      if (!activeCollaboration) {
        collaborationModules.value = null;
        manager.retryViewCreation();
        manager.syncExternalDocument();
        return;
      }

      let cancelled = false;
      onCleanup(() => {
        cancelled = true;
      });
      collaborationModules.value = null;
      void loadCollaborationModules().then(
        (modules) => {
          if (cancelled) {
            return undefined;
          }
          collaborationModules.value = modules;
          manager.retryViewCreation();
          manager.syncExternalDocument();
          publishRemoteSelections();
          return undefined;
        },
        (error: unknown) => {
          if (cancelled) {
            return undefined;
          }
          collaborationModules.value = null;
          onError?.(error instanceof Error ? error : new Error(String(error)));
          return undefined;
        },
      );
    },
    { immediate: true },
  );

  watch(
    () => ({
      awareness: toValue(collaboration)?.awareness,
      modules: collaborationModules.value,
      view: editorView.value,
    }),
    ({ awareness, modules, view }, _previous, onCleanup) => {
      if (!awareness || !modules || !view) {
        remoteSelections.value = [];
        return;
      }

      awareness.on("change", publishRemoteSelections);
      publishRemoteSelections();
      onCleanup(() => {
        awareness.off("change", publishRemoteSelections);
        remoteSelections.value = [];
      });
    },
    { flush: "post" },
  );

  function handleTransaction(transaction: Transaction, newState: EditorState): void {
    editorState.value = newState;
    if (transaction.docChanged) {
      isDirty.value = true;
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

  watch([() => toValue(editorMode), () => toValue(author)], () => {
    const view = editorView.value;
    if (view) {
      syncSuggestionMode(view);
    }
  });

  // Keep the hidden view's editable() ARIA state in sync with readOnly.
  watch(
    () => toValue(readOnly),
    () => manager.syncEditable(),
  );

  // Live-toggle the template directive/slash-menu pair without tearing the view
  // down: add or remove the two known plugin instances via `reconfigure`, so
  // history and selection survive the toggle. Mirrors React swapping its
  // `editorPlugins` array when `showTemplateDirectives` flips.
  function syncTemplatePlugins(view: EditorView): void {
    const enabled = templatePluginsEnabled();
    const installed = view.state.plugins.includes(templateDirectivesPlugin);
    if (enabled === installed) {
      return;
    }
    const plugins = view.state.plugins.filter(
      (plugin) => plugin !== templateDirectivesPlugin && plugin !== templateSlashMenu,
    );
    if (enabled) {
      // Insert at the same position as `buildExternalPlugins` (right before the
      // preview plugin) so ProseMirror prop precedence is identical whether the
      // pair was toggled live or installed on a fresh/reloaded view.
      const previewIndex = plugins.indexOf(templatePreviewPlugin);
      const insertAt = previewIndex === -1 ? plugins.length : previewIndex;
      plugins.splice(insertAt, 0, templateDirectivesPlugin, templateSlashMenu);
    }
    view.updateState(view.state.reconfigure({ plugins }));
  }

  watch(
    () => templatePluginsEnabled(),
    () => {
      const view = editorView.value;
      if (view) {
        syncTemplatePlugins(view);
      }
    },
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
    // The freshly loaded document has no unsaved edits yet (mirrors React
    // clearing its dirty signals on document reset).
    isDirty.value = false;
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
    if (!model || toValue(collaboration)) {
      return;
    }
    try {
      const initialState = createHiddenEditorState({
        document: model,
        styles: model.package.styles ?? null,
        manager: extensionManager,
        externalPlugins: buildExternalPlugins(),
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
      const doc = await parseDocx(buffer, { password: toValue(password) });
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

  async function save(saveOptions?: { selective?: boolean }): Promise<Blob | null> {
    const view = editorView.value;
    const base = docModel.value;
    if (!view || !base) {
      return null;
    }

    // Snapshot the editor state once, before any awaited dynamic imports, so
    // the document we serialize and the selective-save change signals are all
    // read from the same snapshot. A later `view.state` read could pick up an
    // edit that landed mid-save and diff against the wrong baseline.
    const state = view.state;

    const { resolveSelectiveSaveFlags } = await import("@stll/folio-core/docx/selectiveSaveFlags");
    const flags = resolveSelectiveSaveFlags(toValue(featureFlags));

    const updatedDoc = fromProseDoc(state.doc, base);
    const baselineBuffer = updatedDoc.originalBuffer ?? null;

    // The tripwire observes the selective path independently of the user-visible
    // save mode. Only `useSelectiveForSave` is allowed to pick the returned bytes.
    const useSelectiveForSave = flags.selectiveSave && saveOptions?.selective !== false;
    const shouldAttemptSelective = useSelectiveForSave || flags.selectiveSaveTripwire;

    const { repackDocx, createDocx } = await import("@stll/folio-core/docx/rezip");

    let selectiveBuffer: ArrayBuffer | null = null;
    if (shouldAttemptSelective && baselineBuffer) {
      const { attemptSelectiveSave } = await import("@stll/folio-core/docx/selectiveSave");
      selectiveBuffer = await attemptSelectiveSave(updatedDoc, baselineBuffer, {
        changedParaIds: getChangedParagraphIds(state),
        structuralChange: hasStructuralChanges(state),
        hasUntrackedChanges: hasUntrackedChanges(state),
        maxBytes: flags.selectiveSaveMaxBytes,
      });
    }

    let buffer: ArrayBuffer | null = useSelectiveForSave ? selectiveBuffer : null;
    let fullBuffer: ArrayBuffer | null = null;

    if (!buffer) {
      // No original buffer means a from-scratch document — build one via createDocx.
      fullBuffer = baselineBuffer ? await repackDocx(updatedDoc) : await createDocx(updatedDoc);
      buffer = fullBuffer;
    } else if (flags.selectiveSaveTripwire) {
      try {
        fullBuffer = await repackDocx(updatedDoc);
      } catch {
        // Tripwire-only full repack failures must never poison a successful
        // selective save.
      }
    }

    if (flags.selectiveSaveTripwire && fullBuffer && onSelectiveSaveTripwire) {
      // The comparison itself never blocks the save path.
      try {
        const { compareSelectiveVsFull } =
          await import("@stll/folio-core/docx/selectiveSaveTripwire");
        onSelectiveSaveTripwire(await compareSelectiveVsFull(selectiveBuffer, fullBuffer));
      } catch {
        // Comparison failures must never poison the save path.
      }
    }

    // The bytes are produced: the live doc state is now serialized, so no doc
    // edits are pending. Mirrors React's handleSave clearing the change tracker.
    // (The tracker itself is intentionally left uncleared — it is the
    // selective-save baseline; see the featureFlags parity note.) The comment-
    // list dirty flag lives in useCommentManagement and is cleared by the ref
    // API's save wrapper, the single public save entry point.
    isDirty.value = false;

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
    remoteSelections.value = [];
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
    setDocument,
    getCommands,
    focus,
    reLayout,
    destroy,
  };
}
