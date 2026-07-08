import type { CSSProperties, VNodeChild } from "vue";

import type { Plugin } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import type { XmlFragment } from "yjs";

import type {
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
  FolioCommentAnchor,
  FolioReviewChange,
} from "@stll/folio-core/ai-edits";
import type {
  ContentControlFilter,
  SetContentControlContentInput,
  SetContentControlValueInput,
} from "@stll/folio-core/content-controls";
import type { FolioEditor } from "@stll/folio-core/controller/folioEditor";
import type { DocxCompatibility } from "@stll/folio-core/docx/compatibility";
import type { FolioSelectiveSaveFlags } from "@stll/folio-core/docx/selectiveSaveFlags";
import type { TripwireResult } from "@stll/folio-core/docx/selectiveSaveTripwire";
import type { SelectionState, TableContextInfo } from "@stll/folio-core/prosemirror";
import type { AnonymizationMatch } from "@stll/folio-core/prosemirror/plugins/anonymizationDecorations";
import type {
  TemplateSlashMenuKeyAction,
  TemplateSlashMenuState,
} from "@stll/folio-core/prosemirror/plugins/templateSlashMenu";
import type { Document, ParagraphAlignment, SdtProperties, TabStop, Theme } from "@stll/folio-core/types/document";
import type { Comment } from "@stll/folio-core/types/content";
import type { DocxInput } from "@stll/folio-core/utils/docxInput";
import type { ScrollToParaIdOptions } from "@stll/folio-core/paged-layout/paragraphFlash";
// `EditorMode` is owned by core's `EditorModeManager`.
import type { EditorMode } from "@stll/folio-core/managers/EditorModeManager";
// Re-exported so chrome SFCs (e.g. EditingModeDropdown.vue) can import the mode
// union from the adapter's own type surface rather than reaching into core.
export type { EditorMode };
import type { DocumentLoadState } from "@stll/folio-core/managers/DocumentLoaderManager";

import type { FolioUIComponents } from "../../ui/folio-ui";
import type { FontOption } from "../../utils/fontOptions";
import type { PagedEditorRef } from "./pagedEditorRef";

/**
 * Declarative description of one custom font face the host registers with the
 * editor (the `fonts` prop). Multiple entries can share `family` for weights.
 *
 * TODO(vue): move to fontLoader when it lands (mirrors React's
 * `packages/react/src/paged-editor/hostFonts` `FontDefinition`).
 */
export type FontDefinition = {
  /** CSS `font-family` name to expose; match the family your documents reference. */
  family: string;
  /** URL to the font file (woff2, woff, ttf, or otf). */
  src: string;
  /** CSS `font-weight` for this face (a number like `700`, or a keyword). Defaults to `normal`. */
  weight?: number | string;
};

/** Image context surfaced to the toolbar when the cursor is on an image node. */
export type ImageContextInfo = {
  pos: number;
  wrapType: string;
  displayMode: string;
  cssFloat: string | null;
  transform: string | null;
  alt: string | null;
  borderWidth: number | null;
  borderColor: string | null;
  borderStyle: string | null;
};

/** Tracked-change context surfaced to the contextual review toolbar. */
export type ActiveTrackedChangeInfo = {
  type: "insertion" | "deletion";
  author: string;
  date: string | null;
  from: number;
  to: number;
};

export type DocxEditorProps = {
  /** Document data — ArrayBuffer, Uint8Array, Blob, or File */
  documentBuffer?: DocxInput | null;
  /** Password for Agile-encrypted .docx files (Office 2010+). */
  password?: string | undefined;
  /** Pre-parsed document (alternative to documentBuffer) */
  document?: Document | null;
  /**
   * Stable identity of the loaded document (same across internal edits, distinct
   * per loaded file). Distinguishes a genuine external load from an edited
   * document round-tripped back through state. Falls back to a document-metadata
   * signature when omitted.
   */
  documentKey?: string;
  /** Callback when document is saved */
  onSave?: (buffer: ArrayBuffer) => void;
  /** Author name used for comments and track changes */
  author?: string;
  /** Callback when document changes */
  onChange?: (document: Document) => void;
  /** Callback when selection changes */
  onSelectionChange?: (state: SelectionState | null) => void;
  /**
   * Fired with the resolved PM positions and selected plain text on every
   * selection-bearing transaction. Atom inline nodes (tab, hard_break) collapse
   * to a single space; empty string when the selection is collapsed.
   */
  onSelectionTextChange?: (selection: { from: number; to: number; text: string }) => void;
  /** Callback on error */
  onError?: (error: Error) => void;
  /** Callback when fonts are loaded */
  onFontsLoaded?: () => void;
  /**
   * Custom families shown in the toolbar's font-family dropdown. Strings render
   * as plain names; pass `FontOption` objects for a CSS fallback chain and
   * category grouping. Omit for folio's defaults; an empty array renders an
   * empty (but enabled) dropdown. Pass a stable reference.
   */
  fontFamilies?: ReadonlyArray<string | FontOption>;
  /**
   * Custom font faces the host registers with the browser so runs render in
   * them. Each entry injects one face; multiple entries can share `family` for
   * different weights. Registration is best-effort (a malformed entry is
   * skipped). Match `family` to the applied `fontFamilies` name. Pass a stable
   * reference.
   */
  fonts?: ReadonlyArray<FontDefinition>;
  /** Theme for styling */
  theme?: Theme | null;
  /** Whether to show toolbar (default: true) */
  showToolbar?: boolean;
  /** Whether to show zoom control (default: true) */
  showZoomControl?: boolean;
  /**
   * Whether to show the review controls — the track-changes toggle and the
   * markup display-mode selector (default: true). Turn off for plain-markdown
   * editing, where tracked changes and markup views are meaningless.
   */
  showReviewControls?: boolean;
  /**
   * Whether the page header/footer can be edited — double-click-to-edit and the
   * hover hints (default: true). Turn off for plain-markdown editing, which has
   * no running header/footer.
   */
  showHeaderFooterEditing?: boolean;
  /** Whether to show page margin guides/boundaries (default: false) */
  showMarginGuides?: boolean;
  /** Color for margin guides (default: '#c0c0c0') */
  marginGuideColor?: string;
  /**
   * Whether the rulers start visible (default: false). Seeds the toolbar's
   * ruler toggle; users show/hide the rulers at runtime (like Word's
   * View > Ruler).
   */
  showRuler?: boolean;
  /** Measurement unit shown on the rulers (default: 'inch'). */
  rulerUnit?: "inch" | "cm";
  /** Initial zoom level (default: 1.0) */
  initialZoom?: number;
  /** Whether Ctrl/Cmd+wheel and trackpad-pinch zoom are enabled (default: true) */
  enableWheelZoom?: boolean;
  /** Whether the editor is read-only. When true, hides toolbar and rulers */
  readOnly?: boolean;
  /** Whether comments/tracked changes should auto-open the review sidebar (default: true) */
  autoOpenReviewSidebar?: boolean;
  /**
   * Override folio's built-in chrome UI primitives (Button, …) with a
   * consumer's design-system components. Omitted entries fall back to folio's
   * defaults, so standalone folio works without any injection.
   */
  components?: Partial<FolioUIComponents>;
  /** Custom toolbar actions */
  toolbarExtra?: VNodeChild;
  /** Additional CSS class name */
  className?: string;
  /** Additional inline styles */
  style?: CSSProperties;
  /** Placeholder when no document */
  placeholder?: VNodeChild;
  /** Loading indicator */
  loadingIndicator?: VNodeChild;
  /** Keep the current parsed document visible while a new buffer is loading. */
  preserveDocumentWhileLoading?: boolean;
  /** Initial scroll offset for the editor's document scroll container. */
  initialScrollTop?: number;
  /** Callback when the editor's document scroll container scrolls. */
  onScrollTopChange?: (scrollTop: number) => void;
  /** Whether to show the document outline sidebar (default: false) */
  showOutline?: boolean;
  /** Whether to show print button in toolbar (default: true) */
  showPrintButton?: boolean;
  /** Callback when print is triggered */
  onPrint?: () => void;
  /**
   * Insert controls are opt-in: each renders only when its handler is provided
   * (and, for tables, `showTableInsert` is not false). Omit a handler to keep
   * that control out of the toolbar. Wire to the view-level helpers exported
   * from the package (`insertTableInView`, `insertPageBreakInView`,
   * `insertTableOfContentsInView`, `insertImageFromFile`).
   */
  onInsertImage?: (() => void) | undefined;
  /** Insert a `rows × columns` table at the current selection. */
  onInsertTable?: ((rows: number, columns: number) => void) | undefined;
  /** Whether the Insert Table control is shown when `onInsertTable` is set (default: true). */
  showTableInsert?: boolean | undefined;
  /** Insert a page break at the current selection. */
  onInsertPageBreak?: (() => void) | undefined;
  /** Insert a table of contents generated from the document's headings. */
  onInsertTOC?: (() => void) | undefined;
  /** Callback when content is copied */
  onCopy?: () => void;
  /** Callback when content is cut */
  onCut?: () => void;
  /** Callback when content is pasted */
  onPaste?: () => void;
  /** Editor mode: 'editing' (direct edits), 'suggesting' (track changes), or 'viewing' (read-only). Default: 'editing' */
  mode?: EditorMode;
  /** Callback when the editing mode changes */
  onModeChange?: (mode: EditorMode) => void;
  /** Callback when a readonly user action would mutate the document. */
  onReadonlyEditAttempt?: () => void;
  /**
   * Controlled comments array. When provided, the editor reads comment thread
   * metadata from this prop and emits every change through `onCommentsChange`.
   * Use with collaboration backends (Yjs, etc.); PM carries range markers,
   * thread metadata lives outside the doc.
   */
  comments?: Comment[];
  /** Fires whenever the comments array changes (controlled and uncontrolled). */
  onCommentsChange?: (comments: Comment[]) => void;
  /** Callback with the parsed document's editing compatibility report. */
  onCompatibilityChange?: (compatibility: DocxCompatibility) => void;
  /**
   * Fires when the live ProseMirror view is captured (or torn down). The host
   * wires this to drive the AI suggestion overlay (decoration meta, apply,
   * scroll-to) from outside the editor.
   */
  onEditorViewReady?: (view: EditorView | null) => void;
  /** Yjs-backed collaboration owner. Experimental and opt-in. */
  collaboration?: DocxEditorCollaboration | undefined;
  /**
   * Fires with the current anonymization match list every time the decoration
   * plugin recomputes it (mount, term push, doc edit, async DOCX load). The
   * host mirrors per-document counts and "matching workspace terms" to the
   * inspector facet without polling the plugin state.
   */
  onAnonymizationMatchesChange?: (matches: readonly AnonymizationMatch[]) => void;
  /**
   * Fires when the user clicks an anonymization highlight in the rendered
   * document. Hosts push the selection into a sidebar bridge so the inspector
   * facet can scroll/flash the matching row.
   */
  onAnonymizationTermClick?: ((canonical: string, label: string) => void) | undefined;
  /**
   * Canonical to mark as selected in the rendered document. The first matching
   * rect scrolls into view whenever `anonymizationSelectionSeq` increments (so
   * repeated sidebar clicks of the same term re-trigger the scroll).
   */
  selectedAnonymizationCanonical?: string | null | undefined;
  /** Monotonic counter from the bridge store; drives the re-scroll. */
  anonymizationSelectionSeq?: number | undefined;
  /**
   * Render legal-template markers ({{field}}, {{@clause:..}}, {{#if}}/{{#each}})
   * as rich widgets on the page instead of raw text. Off for ordinary
   * documents; on for the template editor.
   */
  showTemplateDirectives?: boolean | undefined;
  /**
   * Fires when the template editor's `/` slash-command trigger opens, its typed
   * query changes, or it closes. Only active alongside `showTemplateDirectives`.
   * The host renders the floating menu and inserts the chosen marker; the plugin
   * owns only the trigger state.
   */
  onSlashMenuChange?: ((state: TemplateSlashMenuState) => void) | undefined;
  /**
   * Resolves a navigation/commit key (Up / Down / Enter) while the slash menu is
   * open. Return `true` to let the editor swallow the key. Escape is handled
   * inside the plugin and never reaches this callback.
   */
  onSlashMenuKeyAction?: ((action: TemplateSlashMenuKeyAction) => boolean) | undefined;
  /**
   * Host-provided text context-menu entries, appended after the built-ins.
   * `requiresSelection` items only render when text is selected. Selecting one
   * fires `onCustomContextAction` with the item id and the selection range
   * captured when the menu opened (right-click can collapse the live PM
   * selection, so the captured range is the reliable one).
   */
  customContextMenuItems?:
    | readonly {
        id: string;
        label: string;
        requiresSelection?: boolean;
        icon?: VNodeChild;
      }[]
    | undefined;
  onCustomContextAction?:
    | ((id: string, selectionRange: { from: number; to: number }) => void)
    | undefined;
  /**
   * Operational flags for save-path features. Selective save and its tripwire
   * mode are OFF by default; hosts opt in once their rollout pipeline is ready.
   */
  featureFlags?: FolioSelectiveSaveFlags;
  /**
   * Receives the structured comparison between selective save and full repack
   * whenever {@link FolioSelectiveSaveFlags.selectiveSaveTripwire} is enabled.
   * Saves are never blocked by the tripwire; this is observability only.
   */
  onSelectiveSaveTripwire?: (result: TripwireResult) => void;
};

export type { FolioSelectiveSaveFlags, TripwireResult };

export type DocxEditorCollaboration = {
  awareness?:
    | {
        clientID: number;
        getStates: () => Map<number, unknown>;
        off: (event: "change" | "update", handler: () => void) => void;
        on: (event: "change" | "update", handler: () => void) => void;
      }
    | undefined;
  onSeeded?: (() => void) | undefined;
  plugins?: Plugin[] | undefined;
  shouldSeed?: boolean | undefined;
  yXmlFragment: XmlFragment;
};

/**
 * Imperative handle exposed by the DocxEditor component.
 */
export type DocxEditorRef = {
  /** Get the current document */
  getDocument: () => Document | null;
  /** Whether the live ProseMirror state has edits that have not been serialized. */
  hasPendingChanges: () => boolean;
  /** Get the editor ref */
  getEditorRef: () => PagedEditorRef | null;
  /** The headless editor controller (Seam 6), or null before the editor mounts. */
  getEditor: () => FolioEditor | null;
  /** Save the document to buffer. Pass { selective: false } to force full repack. */
  save: (options?: { selective?: boolean }) => Promise<ArrayBuffer | null>;
  /** Set zoom level */
  setZoom: (zoom: number) => void;
  /** Get current zoom level */
  getZoom: () => number;
  /** Focus the editor */
  focus: () => void;
  /** Get current page number */
  getCurrentPage: () => number;
  /** Get total page count */
  getTotalPages: () => number;
  /** Scroll to a specific page */
  scrollToPage: (pageNumber: number) => void;
  /**
   * Scroll the paginated view to the paragraph with the given Word `w14:paraId`.
   * Pass `options.highlight` to briefly flash it in a custom color.
   */
  scrollToParaId: (paraId: string, options?: ScrollToParaIdOptions) => boolean;
  /** Open print preview */
  openPrintPreview: () => void;
  /** Print the document directly */
  print: () => void;
  /** Load a pre-parsed document programmatically */
  loadDocument: (doc: Document) => void;
  /** Load a DOCX buffer programmatically (ArrayBuffer, Uint8Array, Blob, or File) */
  loadDocumentBuffer: (buffer: DocxInput) => Promise<void>;
  /**
   * Force-create the underlying editor view if it has been deferred. Surfaces
   * that need a live view before the user interacts with the document body
   * (notably the AI chat composer, which gates "send" on snapshot availability)
   * must call this once on mount with `focus: false`; otherwise the perf
   * deferral keeps the view null and they wait forever.
   */
  ensureEditorView: (options?: { focus?: boolean }) => void;
  /** Create the block snapshot that an external AI editor should reference. */
  createAIEditSnapshot: () => FolioAIEditSnapshot | null;
  /** Apply AI-authored operations against a previously created block snapshot. */
  applyAIEditOperations: (options: {
    snapshot: FolioAIEditSnapshot;
    operations: FolioAIEditOperation[];
    mode?: FolioAIEditApplyMode;
    author?: string;
  }) => FolioAIEditApplyResult;
  /**
   * Accept the tracked-change marks belonging to a previously applied AI edit.
   * Pass a single id for inserts/standalone deletions, or the full id list
   * (`applied.revisionIds`) for a replace, which has separate ids for its
   * deletion and insertion sides. Returns whether a matching range was found.
   */
  acceptAIEditOperation: (revisionIds: number | readonly number[]) => boolean;
  /**
   * Reject the tracked-change marks belonging to a previously applied AI edit.
   * Same id semantics as `acceptAIEditOperation`.
   */
  rejectAIEditOperation: (revisionIds: number | readonly number[]) => boolean;
  /**
   * Undo / redo the last history-recorded edit IN PLACE — e.g. an
   * `acceptAIEditOperation` / `rejectAIEditOperation` resolution — without
   * reloading the document, so the view and scroll position are preserved.
   * Routes to whichever surface is active: the inline header/footer editor when
   * one is being edited, otherwise the document body. Only edits that produce
   * document changes are recorded; selection-only changes carry no steps and are
   * never undone. Returns `false` when there is nothing to undo / redo.
   */
  undo: () => boolean;
  redo: () => boolean;
  /**
   * Scroll the editor viewport so the tracked-change marks belonging to the
   * given `revisionIds` come into view, and select them. No-op when none of the
   * revisions are present.
   */
  scrollToAIEditOperation: (revisionIds: number | readonly number[]) => boolean;
  /**
   * Scroll the editor viewport so the block referenced by `blockId` is in view,
   * and place the selection inside it. Returns `false` when the block can't be
   * resolved on the live document (e.g. it was edited away).
   *
   * When `snapshot` is supplied, ids resolve against it — required for review
   * panel pending-suggestion navigation, because block ids are sequential and a
   * freshly recomputed snapshot would re-number blocks after a structural
   * accept. Without `snapshot`, falls back to a fresh-from-live-doc snapshot.
   */
  scrollToBlock: (blockId: string, snapshot?: FolioAIEditSnapshot) => boolean;

  // -------------------------------------------------------------------------
  // Read surface (agents, review tooling)
  // -------------------------------------------------------------------------

  /**
   * The tracked changes (insertions and deletions) present in the live
   * document, read from the current ProseMirror `state.doc`. Empty array
   * before the editor view mounts.
   */
  getTrackedChanges: () => FolioReviewChange[];
  /**
   * The comment anchors present in the live document, read from the current
   * ProseMirror `state.doc`. Empty array before the editor view mounts.
   */
  getCommentAnchors: () => FolioCommentAnchor[];
  /**
   * The plain text of the current selection, or `""` when the editor has no
   * view or the selection is collapsed.
   */
  getSelectionText: () => string;
  /**
   * The plain text of a single rendered page (1-based), joining each of its
   * layout fragments with a newline. Returns `null` when the page number is
   * out of range or the layout hasn't been computed yet.
   */
  getPageText: (page: number) => string | null;

  // -------------------------------------------------------------------------
  // Block-level content controls (w:sdt)
  // -------------------------------------------------------------------------

  /**
   * List every block-level content control in the live document, filtered by
   * tag / alias / id / sdtType. Returns the modeled SdtProperties projection
   * plus a synthetic identifier so callers can address each control by insertion
   * order even if tag/alias collide.
   *
   * IMPORTANT: the returned `pmPos` values are valid only for the document state
   * at fetch time. After any mutation that changes earlier content's size, later
   * controls shift and stored `pmPos` values become stale. For batch-fill flows
   * over duplicate-tag controls, either refetch via `getContentControls()` after
   * each mutation, or address by stable `{ tag }` / `{ id }` / `{ alias }`.
   */
  getContentControls: (filter?: ContentControlFilter) => {
    properties: SdtProperties;
    /** Index in the PM document tree, outer→inner. */
    path: number[];
    /**
     * PM position of the control's open token in the SOURCE document state.
     * Disambiguates duplicates within one snapshot, but does not survive
     * mutations that resize earlier content — see `getContentControls`.
     */
    pmPos: number;
  }[];
  /**
   * Scroll to the first content control matching the filter and place the caret
   * inside it. Returns false when no match is present.
   */
  scrollToContentControl: (filter: ContentControlFilter) => boolean;
  /**
   * Replace a control's inner content with plain text or pre-built block
   * content. Locked controls throw `ContentControlLockedError` unless
   * `{ force: true }` is passed. Writes go through a normal undoable transaction.
   */
  setContentControlContent: (
    filter: ContentControlFilter,
    input: SetContentControlContentInput,
    options?: { force?: boolean },
  ) => boolean;
  /**
   * Set a structured value on a typed control (dropdown / checkbox / date).
   * Throws `ContentControlTypeError` when the input shape does not match the
   * control's sdtType. Writes go through a normal undoable transaction.
   */
  setContentControlValue: (
    filter: ContentControlFilter,
    input: SetContentControlValueInput,
    options?: { force?: boolean },
  ) => boolean;
  /**
   * Remove a control. With `{ keepContent: true }` the inner blocks survive
   * in-place; otherwise the wrapper and its children are dropped. Refuses to
   * unwrap a `w15:repeatingSection` (would orphan its row items) unless
   * `{ force: true }` is passed.
   */
  removeContentControl: (
    filter: ContentControlFilter,
    options?: { keepContent?: boolean; force?: boolean },
  ) => boolean;
};

/**
 * Selection formatting snapshot for the toolbar.
 *
 * TODO(vue): mirror the React toolbar's `SelectionFormatting` exactly once the
 * Vue toolbar lands; `EditorState` is internal, so fidelity here is secondary to
 * `DocxEditorProps` / `DocxEditorRef`.
 */
export type SelectionFormatting = {
  /** Whether selected text is bold */
  bold?: boolean | undefined;
  /** Whether selected text is italic */
  italic?: boolean | undefined;
  /** Whether selected text is underlined */
  underline?: boolean | undefined;
  /** Whether selected text has strikethrough */
  strike?: boolean | undefined;
  /** Whether selected text is superscript */
  superscript?: boolean | undefined;
  /** Whether selected text is subscript */
  subscript?: boolean | undefined;
  /** Font family of selected text */
  fontFamily?: string | undefined;
  /** Font size of selected text (in half-points) */
  fontSize?: number | undefined;
  /** Text color */
  color?: string | undefined;
  /** Highlight color */
  highlight?: string | undefined;
  /** Paragraph alignment */
  alignment?: ParagraphAlignment | undefined;
  /** List state of the current paragraph */
  listState?: ListState | undefined;
  /** Line spacing in twips (OOXML value, 240 = single spacing) */
  lineSpacing?: number | undefined;
  /** Paragraph style ID */
  styleId?: string | undefined;
  /** Paragraph left indentation in twips */
  indentLeft?: number | undefined;
  /** Whether the paragraph is RTL (bidi) */
  bidi?: boolean | undefined;
};

/**
 * List state of the current paragraph.
 *
 * TODO(vue): source from the Vue list toolbar (mirrors React's
 * `components/ui/ListButtons` `ListState`) when it lands.
 */
export type ListState = {
  type: string;
  level: number;
  isInList: boolean;
  numId?: number;
};

/** Aggregated internal state held by DocxEditor's top-level reducer slot. */
export type EditorState = {
  documentLoad: DocumentLoadState;
  /** Current selection formatting for toolbar */
  selectionFormatting: SelectionFormatting;
  /** Paragraph indent data for ruler */
  paragraphIndentLeft: number;
  paragraphIndentRight: number;
  paragraphFirstLineIndent: number;
  paragraphHangingIndent: boolean;
  paragraphTabs: TabStop[] | null;
  /** ProseMirror table context (for showing table toolbar) */
  pmTableContext: TableContextInfo | null;
  /** Image context when cursor is on an image node */
  pmImageContext: ImageContextInfo | null;
  /** Active tracked change at cursor (for contextual toolbar) */
  activeTrackedChange: ActiveTrackedChangeInfo | null;
};
