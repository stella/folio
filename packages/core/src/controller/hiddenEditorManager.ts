import { OP_STORIES } from "@stll/docx-core/ops";
/**
 * Hidden-editor view lifecycle manager
 *
 * Framework-agnostic owner of the off-screen ProseMirror EditorView. Holds the
 * view plus the bookkeeping that decides when an incoming document is a truly
 * external change (vs. an internal edit echoed back through props), builds the
 * editor state and `editorProps`, and composes the slice-2a imperative API so
 * the API and the lifecycle share one view owner. The React adapter
 * (`HiddenProseMirror`) keeps the input refs and its effects (which decide
 * *when* to act) and drives this manager through the methods below.
 */

import { panic, Result, TaggedError } from "better-result";
import type { EditorState, Plugin, Transaction } from "prosemirror-state";
import { EditorState as PMEditorState } from "prosemirror-state";
import type { DirectEditorProps } from "prosemirror-view";
import { EditorView } from "prosemirror-view";
import type * as YProseMirror from "y-prosemirror";
import type { Doc as YDoc, XmlFragment } from "yjs";
import type * as Yjs from "yjs";

import {
  recordHiddenEditorPhase,
  recordHiddenEditorStateCreate,
  type HiddenEditorStateReason,
} from "../layout-engine/layoutInstrumentation";
import { suppressHiddenEditorScrollToSelection } from "../paged-layout/hiddenEditorScroll";
import { isReadOnlyEditKey } from "../paged-layout/readOnlyEditAttempt";
import { toProseDoc, createEmptyDoc } from "../prosemirror/conversion";
import type { ExtensionManager } from "../prosemirror/extensions/ExtensionManager";
import { ensureBaseDirectionInState } from "../prosemirror/extensions/features/AutoBidiDetectionExtension";
import {
  ensureParaIdsInDoc,
  ensureParaIdsInState,
} from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import { createDocumentStylesPlugin } from "../prosemirror/plugins/documentStyles";
import { createDocumentNumberingPlugin } from "../prosemirror/plugins/documentNumbering";
import { schema } from "../prosemirror/schema";
import { createTextInputPlugin } from "../prosemirror/textInput";
import {
  applyAttrSchemaMigrations,
  proseDocumentParagraphSourceContract,
  readYjsAttrSchemaVersion,
  readYjsParagraphSourceContract,
  withParagraphSourceContract,
  writeYjsDocumentMetadata,
} from "../prosemirror/yjsDocumentMetadata";
import type { Document, StyleDefinitions } from "../types/document";
import type { RemoteSelection } from "../types/remote-selection";
import type { EditorMode } from "../managers/EditorModeManager";
import { createHiddenEditorApi, type HiddenEditorApi } from "./hiddenEditorApi";
import { createCanonicalInputBoundary } from "./canonicalInput";
import {
  createCanonicalSession,
  publishCanonicalProjection,
  type CanonicalSession,
  type CanonicalSessionMode,
  type CanonicalCommit,
} from "./canonicalSession";
import {
  createParagraphChangeTrackerPlugin,
  markPackageChange,
} from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";

export class CanonicalSessionRefusalError extends TaggedError("CanonicalSessionRefusalError")<{
  message: string;
}> {}

type EditorSession =
  | { type: "prosemirror" }
  | { type: "canonical"; session: CanonicalSession }
  | { type: "refused"; reason: string; documentIdentity: string | null };

// Initial-load normalization. `appendTransaction` does not fire for the seed
// document, so the paraId allocator and RTL base-direction detection are
// applied imperatively to freshly created states.
const normalizeSeedState = (state: EditorState): EditorState =>
  ensureBaseDirectionInState(ensureParaIdsInState(state));

const normalizeLocalSeedState = (state: EditorState): EditorState =>
  ensureBaseDirectionInState(state);

type YProseMirrorModule = typeof YProseMirror;
type YjsModule = typeof Yjs;

export type CollaborationModules = {
  yProseMirror: YProseMirrorModule;
  yjs: YjsModule;
};

type CollaborationAwareness = {
  clientID: number;
  getStates: () => Map<number, unknown>;
  off: (event: "change" | "update", handler: () => void) => void;
  on: (event: "change" | "update", handler: () => void) => void;
};

export type HiddenProseMirrorCollaboration = {
  awareness?: CollaborationAwareness | undefined;
  onSeeded?: (() => void) | undefined;
  shouldSeed?: boolean | undefined;
  yXmlFragment: XmlFragment;
};

export type HiddenProseMirrorRemoteSelection = RemoteSelection;

type YSyncState = {
  binding: {
    mapping: Parameters<YProseMirrorModule["relativePositionToAbsolutePosition"]>[3];
  };
  doc: YDoc;
  type: XmlFragment;
};

type AwarenessCursor = {
  anchor: Record<string, unknown>;
  head: Record<string, unknown>;
};

type AwarenessUser = {
  color: string;
  name: string;
};

const isObjectRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null;

/**
 * A remote collaborator's awareness `user.color` is read from another
 * peer's Yjs awareness state — effectively untrusted network input — and
 * gets fed straight into `element.style.background`/`backgroundColor` (see
 * `RemoteSelectionOverlay`). Restrict it to a plain 6-digit hex color before
 * accepting it so a crafted value (e.g. a CSS `url(...)` payload) can't ride
 * along as a CSS-injection / beacon vector; anything else falls back to the
 * default selection color below.
 */
const AWARENESS_HEX_COLOR_PATTERN = /^#[0-9A-Fa-f]{6}$/u;

const isYSyncState = (value: unknown, yjs: YjsModule): value is YSyncState => {
  if (!isObjectRecord(value)) {
    return false;
  }
  const binding = value["binding"];
  return (
    value["doc"] instanceof yjs.Doc &&
    value["type"] instanceof yjs.XmlFragment &&
    isObjectRecord(binding) &&
    binding["mapping"] instanceof Map
  );
};

const readAwarenessCursor = (state: unknown): AwarenessCursor | null => {
  if (!isObjectRecord(state)) {
    return null;
  }
  const cursor = state["cursor"];
  if (!isObjectRecord(cursor)) {
    return null;
  }
  const anchor = cursor["anchor"];
  const head = cursor["head"];
  if (!isObjectRecord(anchor) || !isObjectRecord(head)) {
    return null;
  }
  return { anchor, head };
};

const readAwarenessUser = (state: unknown, clientId: number): AwarenessUser => {
  if (!isObjectRecord(state) || !isObjectRecord(state["user"])) {
    return {
      color: "var(--doc-image-selection)",
      name: `User ${clientId}`,
    };
  }

  const user = state["user"];
  const rawColor = user["color"];
  const color =
    typeof rawColor === "string" && AWARENESS_HEX_COLOR_PATTERN.test(rawColor)
      ? rawColor
      : "var(--doc-image-selection)";
  return {
    color,
    name: typeof user["name"] === "string" ? user["name"] : `User ${clientId}`,
  };
};

export const collectRemoteSelections = (
  state: EditorState,
  awareness: CollaborationAwareness,
  collaborationModules: CollaborationModules,
): HiddenProseMirrorRemoteSelection[] => {
  const syncState: unknown = collaborationModules.yProseMirror.ySyncPluginKey.getState(state);
  if (!isYSyncState(syncState, collaborationModules.yjs)) {
    return [];
  }

  const selections: HiddenProseMirrorRemoteSelection[] = [];
  for (const [clientId, remoteState] of awareness.getStates()) {
    if (clientId === awareness.clientID) {
      continue;
    }

    const cursor = readAwarenessCursor(remoteState);
    if (cursor === null) {
      continue;
    }

    const anchor = collaborationModules.yProseMirror.relativePositionToAbsolutePosition(
      syncState.doc,
      syncState.type,
      collaborationModules.yjs.createRelativePositionFromJSON(cursor.anchor),
      syncState.binding.mapping,
    );
    const head = collaborationModules.yProseMirror.relativePositionToAbsolutePosition(
      syncState.doc,
      syncState.type,
      collaborationModules.yjs.createRelativePositionFromJSON(cursor.head),
      syncState.binding.mapping,
    );
    if (anchor === null || head === null) {
      continue;
    }

    const user = readAwarenessUser(remoteState, clientId);
    selections.push({
      anchor,
      clientId,
      color: user.color,
      head,
      name: user.name,
    });
  }

  return selections;
};

const HIDDEN_EDITOR_ATTRIBUTES = {
  "aria-label": "Document content",
  "aria-multiline": "true",
  autocapitalize: "off",
  autocomplete: "off",
  autocorrect: "off",
  role: "textbox",
  spellcheck: "false",
  translate: "no",
};

/**
 * Create ProseMirror state from document
 *
 * When an ExtensionManager is provided, it supplies the schema and plugins.
 * Otherwise falls back to the default singleton schema with no extension plugins.
 */
export type CreateHiddenEditorStateOptions = {
  document: Document | null;
  styles?: StyleDefinitions | null | undefined;
  manager?: ExtensionManager | undefined;
  externalPlugins?: Plugin[] | undefined;
  collaboration?: HiddenProseMirrorCollaboration | undefined;
  collaborationModules?: CollaborationModules | null | undefined;
  reason?: HiddenEditorStateReason | undefined;
};

export function createHiddenEditorState(options: CreateHiddenEditorStateOptions): EditorState {
  const {
    document,
    styles,
    manager,
    externalPlugins = [],
    collaboration,
    collaborationModules,
    reason = "mount",
  } = options;
  recordHiddenEditorStateCreate(reason);

  const activeSchema = manager?.getSchema() ?? schema;
  let localDoc = createEmptyDoc();
  if (document) {
    const startedAt = performance.now();
    localDoc =
      styles === undefined || styles === null
        ? toProseDoc(document)
        : toProseDoc(document, { styles });
    localDoc = ensureParaIdsInDoc(localDoc);
    recordHiddenEditorPhase(reason, "to-prose-doc", performance.now() - startedAt);
  } else {
    localDoc = ensureParaIdsInDoc(localDoc);
  }

  // Expose the document's styles to style-aware commands (e.g. the Enter
  // handler's `w:next` switch from heading to body text). Same resolver for
  // collab and non-collab paths.
  const styleResolverPlugin = createDocumentStylesPlugin(styles ?? document?.package.styles);
  const numberingPlugin = createDocumentNumberingPlugin(document?.package.numbering);
  const plugins: Plugin[] = [
    ...externalPlugins,
    ...(manager?.getPlugins() ?? [createTextInputPlugin()]),
    styleResolverPlugin,
    numberingPlugin,
  ];

  if (collaboration) {
    if (!collaborationModules) {
      panic("Collaboration modules must be loaded before creating collaborative editor state.");
    }

    if (collaboration.shouldSeed && collaboration.yXmlFragment.length === 0) {
      const seedState = normalizeLocalSeedState(
        PMEditorState.create({
          doc: localDoc,
          schema: activeSchema,
          plugins,
        }),
      );
      collaborationModules.yProseMirror.prosemirrorToYXmlFragment(
        seedState.doc,
        collaboration.yXmlFragment,
      );
      const collaborationDocument = collaboration.yXmlFragment.doc;
      if (!collaborationDocument) {
        panic("A collaboration fragment must belong to a Yjs document before seeding.");
      }
      writeYjsDocumentMetadata(collaborationDocument, seedState.doc);
      collaboration.onSeeded?.();
    }

    const collaborationDocument = collaboration.yXmlFragment.doc;
    if (!collaborationDocument) {
      panic("A collaboration fragment must belong to a Yjs document before loading.");
    }
    // Gate before any node is built: y-prosemirror copies unknown attr values
    // into the node verbatim, drops unknown keys, and deletes an element it
    // cannot build. A snapshot this build cannot read must never reach it.
    const attrSchemaVersion = readYjsAttrSchemaVersion(collaborationDocument);
    if (attrSchemaVersion.isErr()) {
      throw attrSchemaVersion.error;
    }
    applyAttrSchemaMigrations(
      collaborationDocument,
      collaboration.yXmlFragment,
      attrSchemaVersion.value,
    );
    let { doc } = collaborationModules.yProseMirror.initProseMirrorDoc(
      collaboration.yXmlFragment,
      activeSchema,
    );
    const collaborationContract = readYjsParagraphSourceContract(collaborationDocument);
    const localContract = proseDocumentParagraphSourceContract(localDoc);
    if (localContract && collaborationContract !== localContract) {
      panic("The collaboration state belongs to a different paragraph-property source.");
    }
    if (collaborationContract) {
      doc = withParagraphSourceContract(doc, collaborationContract);
    }

    const initializedState = normalizeSeedState(
      PMEditorState.create({
        doc,
        schema: activeSchema,
        plugins,
      }),
    );
    if (!initializedState.doc.eq(doc)) {
      collaborationModules.yProseMirror.prosemirrorToYXmlFragment(
        initializedState.doc,
        collaboration.yXmlFragment,
      );
      ({ doc } = collaborationModules.yProseMirror.initProseMirrorDoc(
        collaboration.yXmlFragment,
        activeSchema,
      ));
      if (collaborationContract) {
        doc = withParagraphSourceContract(doc, collaborationContract);
      }
    }

    const startedAt = performance.now();
    const state = normalizeSeedState(
      PMEditorState.create({
        doc,
        schema: activeSchema,
        plugins,
      }),
    );
    recordHiddenEditorPhase(reason, "editor-state", performance.now() - startedAt);
    return state;
  }

  const startedAt = performance.now();
  const state = normalizeLocalSeedState(
    PMEditorState.create({
      doc: localDoc,
      schema: activeSchema,
      plugins,
    }),
  );
  recordHiddenEditorPhase(reason, "editor-state", performance.now() - startedAt);
  return state;
}

function syncHiddenEditorAccessibility(view: EditorView, readOnly: boolean): void {
  const { dom } = view;
  if (!dom.hasAttribute("tabindex")) {
    dom.tabIndex = 0;
  }
  dom.setAttribute("aria-readonly", readOnly ? "true" : "false");
}

export type HiddenEditorManagerDeps = {
  getHost: () => HTMLElement | null;
  getDocument: () => Document | null;
  getStyles: () => StyleDefinitions | null | undefined;
  getExtensionManager: () => ExtensionManager | undefined;
  getExternalPlugins: () => Plugin[];
  getCollaboration: () => HiddenProseMirrorCollaboration | undefined;
  getCollaborationModules: () => CollaborationModules | null;
  getPrecomputedInitialState: () => EditorState | null | undefined;
  getReadOnly: () => boolean;
  getExperimentalSession?: () => "canonical" | undefined;
  getEditingMode?: () => EditorMode;
  getSuggestionAuthor?: () => string;
  onSessionRefusal?: (reason: string) => void;
  /**
   * Identity of the loaded document as tracked by the adapter's loader: the
   * same value across internal edits (so typing does not trigger an external
   * re-sync) and a distinct value per load.
   */
  getDocumentIdentity: () => string;
  /** Document context for the API's `getDocument` (PM state -> Document). */
  getDocumentContext: () => Document | null;
  onTransaction: (update: HiddenEditorTransactionUpdate) => void;
  onSelectionChange: (state: EditorState) => void;
  onKeyDown: (view: EditorView, event: KeyboardEvent) => boolean;
  onCopy?: () => void;
  onCut?: () => void;
  onPaste?: () => void;
  onReadOnlyEditAttempt: () => void;
  onEditorViewReady: (view: EditorView) => void;
  onEditorViewDestroy: () => void;
  onRemoteSelectionsChange: (selections: HiddenProseMirrorRemoteSelection[]) => void;
};

export type HiddenEditorTransactionUpdate = {
  transactions: readonly Transaction[];
  newState: EditorState;
  docChanged: boolean;
};

type PreventableDomEvent = { preventDefault: () => void };

export const createHiddenEditorClipboardHandlers = (
  deps: Pick<
    HiddenEditorManagerDeps,
    "getReadOnly" | "onCopy" | "onCut" | "onPaste" | "onReadOnlyEditAttempt"
  >,
) => ({
  copy: () => {
    deps.onCopy?.();
    return false;
  },
  cut: (_view: unknown, event: PreventableDomEvent) => {
    if (!deps.getReadOnly()) {
      deps.onCut?.();
      return false;
    }
    deps.onReadOnlyEditAttempt();
    event.preventDefault();
    return true;
  },
  paste: (_view: unknown, event: PreventableDomEvent) => {
    if (!deps.getReadOnly()) {
      deps.onPaste?.();
      return false;
    }
    deps.onReadOnlyEditAttempt();
    event.preventDefault();
    return true;
  },
});

export type HiddenEditorManager = {
  /** Request the view (sets the requested flag, then attempts creation). */
  ensureView: () => void;
  /** Re-attempt a previously-requested-but-deferred creation (no-op otherwise). */
  retryViewCreation: () => void;
  isViewRequested: () => boolean;
  destroyView: () => void;
  syncExternalDocument: () => void;
  syncEditable: () => void;
  getView: () => EditorView | null;
  isInitialized: () => boolean;
  api: HiddenEditorApi;
};

export const createHiddenEditorManager = (deps: HiddenEditorManagerDeps): HiddenEditorManager => {
  let view: EditorView | null = null;
  let editorSession: EditorSession = { type: "prosemirror" };
  let modeOverride: CanonicalSessionMode | null = null;
  const syncCanonicalMode = (): void => {
    if (editorSession.type !== "canonical") return;
    editorSession.session.setMode(
      modeOverride ??
        (deps.getEditingMode?.() === "suggesting"
          ? { type: "suggesting", author: deps.getSuggestionAuthor?.() ?? "User" }
          : { type: "editing" }),
    );
  };
  const refuse = (reason: string): void => {
    if (deps.onSessionRefusal) deps.onSessionRefusal(reason);
    else throw new CanonicalSessionRefusalError({ message: reason });
  };
  const seedSession = (document: Document | null): boolean => {
    if (deps.getExperimentalSession?.() !== "canonical") {
      editorSession = { type: "prosemirror" };
      return true;
    }
    const documentIdentity = deps.getDocumentIdentity();
    if (editorSession.type === "refused" && editorSession.documentIdentity === documentIdentity) {
      return false;
    }
    if (deps.getCollaboration() || !document) {
      const reason = deps.getCollaboration()
        ? "Canonical sessions do not support collaboration."
        : "Canonical sessions require a loaded Document.";
      editorSession = { type: "refused", reason, documentIdentity };
      refuse(reason);
      return false;
    }
    const result = createCanonicalSession(document, deps.getStyles());
    if (result.isErr()) {
      editorSession = { type: "refused", reason: result.error.message, documentIdentity };
      refuse(result.error.message);
      return false;
    }
    editorSession = { type: "canonical", session: result.value };
    modeOverride = null;
    syncCanonicalMode();
    return true;
  };
  const canonicalState = (session: CanonicalSession) =>
    PMEditorState.create({
      doc: session.projection.doc,
      plugins: [
        createParagraphChangeTrackerPlugin(),
        createDocumentStylesPlugin(deps.getStyles() ?? session.document.package.styles),
        createDocumentNumberingPlugin(session.document.package.numbering),
      ],
    });
  const publishCommit = (commit: CanonicalCommit): boolean => {
    if (!view || deps.getReadOnly() || editorSession.type !== "canonical") return false;
    const session = editorSession.session;
    const result = publishCanonicalProjection({ state: view.state, commit, session });
    if (result.isErr()) {
      refuse(result.error.message);
      return false;
    }
    const staged = result.value;
    view.updateState(staged.state);
    deps.onTransaction({
      transactions: staged.transactions,
      newState: staged.state,
      docChanged: true,
    });
    deps.onSelectionChange(staged.state);
    return true;
  };
  const history = (direction: "undo" | "redo"): boolean => {
    if (!view || deps.getReadOnly() || editorSession.type !== "canonical") return false;
    const session = editorSession.session;
    if (direction === "undo" ? !session.canUndo : !session.canRedo) return false;
    const prepared =
      direction === "undo" ? session.prepareUndo(view.state) : session.prepareRedo(view.state);
    if (prepared.isErr()) {
      refuse(prepared.error.message);
      return false;
    }
    return publishCommit(prepared.value);
  };
  const canonicalInputLifecycle = {
    breakUndoGroup: () => {
      if (editorSession.type === "canonical") editorSession.session.breakUndoGroup();
    },
    beginComposition: () => {
      if (deps.getReadOnly() || editorSession.type !== "canonical") return false;
      const begun = editorSession.session.beginComposition();
      if (begun.isErr()) {
        refuse(begun.error.message);
        return false;
      }
      return true;
    },
    endComposition: () => {
      if (editorSession.type === "canonical") editorSession.session.endComposition();
    },
  };
  const input = createCanonicalInputBoundary({
    ...canonicalInputLifecycle,
    replace: (intent) => {
      if (!view || deps.getReadOnly() || editorSession.type !== "canonical") return;
      syncCanonicalMode();
      const prepared = editorSession.session.prepareReplace(view.state, intent);
      if (prepared.isErr()) refuse(prepared.error.message);
      else publishCommit(prepared.value);
    },
    split: () => {
      if (!view || deps.getReadOnly() || editorSession.type !== "canonical") return;
      syncCanonicalMode();
      const prepared = editorSession.session.prepareSplit(view.state);
      if (prepared.isErr()) refuse(prepared.error.message);
      else publishCommit(prepared.value);
    },
    join: (direction) => {
      if (!view || deps.getReadOnly() || editorSession.type !== "canonical") return;
      syncCanonicalMode();
      const prepared = editorSession.session.prepareJoin(view.state, direction);
      if (prepared.isErr()) refuse(prepared.error.message);
      else publishCommit(prepared.value);
    },
    undo: () => history("undo"),
    redo: () => history("redo"),
    refuse,
  });
  let isDestroying = false;
  // The React adapter requests the view explicitly (first interaction) or
  // eagerly (collaboration); creation only proceeds once requested.
  let requested = false;
  // Track if we've initialized - first render needs to set up state.
  let isInitialized = false;
  // Track the document identity to detect truly external changes vs changes
  // that originated from editing (which get passed back through props).
  let lastDocumentId: string | null = null;
  let lastCollaborationFragment: XmlFragment | null = null;

  const tryCreate = (): void => {
    if (!requested || view !== null || isDestroying) {
      return;
    }
    const host = deps.getHost();
    if (!host) {
      return;
    }
    const collaboration = deps.getCollaboration();
    const collaborationModules = deps.getCollaborationModules();
    if (collaboration && !collaborationModules && deps.getExperimentalSession?.() !== "canonical") {
      return;
    }

    const precomputedInitialState = deps.getPrecomputedInitialState();
    const document = deps.getDocument();
    const styles = deps.getStyles();
    const extensionManager = deps.getExtensionManager();
    const externalPlugins = deps.getExternalPlugins();

    if (!seedSession(document)) return;
    input.reset();
    const initialState = (() => {
      if (editorSession.type === "canonical") return canonicalState(editorSession.session);
      if (precomputedInitialState && !collaboration) return precomputedInitialState;
      return createHiddenEditorState({
        document,
        styles,
        manager: extensionManager,
        externalPlugins,
        collaboration,
        collaborationModules,
        reason: "mount",
      });
    })();

    const editorProps: DirectEditorProps = {
      state: initialState,
      attributes: HIDDEN_EDITOR_ATTRIBUTES,
      editable: () => !deps.getReadOnly() && editorSession.type !== "refused",
      dispatchTransaction: (transaction: Transaction) => {
        if (!view || isDestroying) {
          return;
        }

        if (deps.getReadOnly() && transaction.docChanged) {
          deps.onReadOnlyEditAttempt();
          return;
        }

        if (editorSession.type === "refused") return;
        if (editorSession.type === "canonical" && transaction.docChanged) {
          if (input.acceptComposition(view, transaction)) return;
          if (!input.commitNativeProposal(view, transaction)) {
            input.refuseNativeMutation(view);
          }
          // PM's DOM observer dirties the view before dispatching a native
          // proposal. Repaint from the committed state even on refusal.
          view.updateState(view.state);
          return;
        }
        if (editorSession.type === "canonical" && transaction.selectionSet && !input.isComposing)
          editorSession.session.breakUndoGroup();
        const applied = view.state.applyTransaction(transaction);
        if (editorSession.type === "canonical" && !applied.state.doc.eq(view.state.doc)) {
          refuse("A plugin attempted an unclassified canonical document mutation.");
          return;
        }
        view.updateState(applied.state);

        const docChanged = applied.transactions.some(
          (appliedTransaction) => appliedTransaction.docChanged,
        );
        const selectionChanged = applied.transactions.some(
          (appliedTransaction) => appliedTransaction.selectionSet || appliedTransaction.docChanged,
        );

        // Notify about transaction.
        deps.onTransaction({
          transactions: applied.transactions,
          newState: applied.state,
          docChanged,
        });

        // Notify about selection changes.
        if (selectionChanged) {
          deps.onSelectionChange(applied.state);
        }

        const currentCollaboration = deps.getCollaboration();
        const currentCollaborationModules = deps.getCollaborationModules();
        if (currentCollaboration?.awareness && currentCollaborationModules) {
          deps.onRemoteSelectionsChange(
            collectRemoteSelections(
              applied.state,
              currentCollaboration.awareness,
              currentCollaborationModules,
            ),
          );
        }
      },
      // Intercept key events before ProseMirror processes them
      handleKeyDown: (pmView: EditorView, event: KeyboardEvent): boolean => {
        if (deps.getReadOnly() && isReadOnlyEditKey(event)) {
          deps.onReadOnlyEditAttempt();
          event.preventDefault();
          return true;
        }

        if (editorSession.type === "canonical" && input.handleKeyDown(pmView, event)) return true;
        return deps.onKeyDown(pmView, event);
      },
      handleTextInput: (pmView, from, to, text) =>
        editorSession.type === "canonical" ? input.handleTextInput(pmView, from, to, text) : false,
      handleScrollToSelection: suppressHiddenEditorScrollToSelection,
      // Prevent focus handling from interfering with visual layer
      handleDOMEvents: {
        focus: () => false,
        blur: (pmView) =>
          editorSession.type === "canonical" ? input.handleDOMEvents.blur(pmView) : false,
        ...createHiddenEditorClipboardHandlers(deps),
        mousedown: (pmView) =>
          editorSession.type === "canonical" ? input.handleDOMEvents.mousedown(pmView) : false,
        compositionstart: (pmView) =>
          editorSession.type === "canonical"
            ? input.handleDOMEvents.compositionstart(pmView)
            : false,
        compositionend: (pmView) =>
          editorSession.type === "canonical" ? input.handleDOMEvents.compositionend(pmView) : false,
        input: (pmView) =>
          editorSession.type === "canonical" ? input.handleDOMEvents.input(pmView) : false,
        paste: (pmView, event) =>
          editorSession.type === "canonical"
            ? input.handleDOMEvents.paste(pmView, event)
            : createHiddenEditorClipboardHandlers(deps).paste(pmView, event),
        cut: (pmView, event) =>
          editorSession.type === "canonical"
            ? input.handleDOMEvents.cut(pmView, event)
            : createHiddenEditorClipboardHandlers(deps).cut(pmView, event),
        beforeinput: (_view, event) => {
          if (editorSession.type === "canonical" && !deps.getReadOnly())
            return input.handleDOMEvents.beforeinput(_view, event);
          if (!deps.getReadOnly()) {
            return false;
          }
          deps.onReadOnlyEditAttempt();
          event.preventDefault();
          return true;
        },
        drop: (_view, event) => {
          if (editorSession.type === "canonical") return input.handleDOMEvents.drop(_view, event);
          if (!deps.getReadOnly()) {
            return false;
          }
          deps.onReadOnlyEditAttempt();
          event.preventDefault();
          return true;
        },
      },
    };

    const viewStartedAt = performance.now();
    view = new EditorView(host, editorProps);
    recordHiddenEditorPhase("mount", "editor-view", performance.now() - viewStartedAt);
    syncHiddenEditorAccessibility(view, deps.getReadOnly());
    isInitialized = true;
    lastDocumentId = deps.getDocumentIdentity();
    lastCollaborationFragment = collaboration?.yXmlFragment ?? null;

    // Notify that view is ready.
    deps.onEditorViewReady(view);
  };

  const ensureView = (): void => {
    requested = true;
    tryCreate();
  };

  // Completes a previously-requested creation once a gate clears (e.g. the
  // async collaboration modules finish loading); a no-op until requested.
  const retryViewCreation = (): void => {
    tryCreate();
  };

  const isViewRequested = (): boolean => requested;

  const destroyView = (): void => {
    if (view && !isDestroying) {
      isDestroying = true;

      deps.onEditorViewDestroy();

      view.destroy();
      view = null;
      input.reset();
      isDestroying = false;
    }
  };

  // Update state when document changes externally (e.g., loading a new file).
  // This should NOT run when the document prop changes due to internal edits
  // being passed back through the parent component's state.
  const syncExternalDocument = (): void => {
    if (!view || isDestroying) {
      return;
    }
    const collaboration = deps.getCollaboration();
    const collaborationModules = deps.getCollaborationModules();
    if (collaboration && !collaborationModules && deps.getExperimentalSession?.() !== "canonical") {
      return;
    }

    const document = deps.getDocument();
    const currentDocId = deps.getDocumentIdentity();
    const currentCollaborationFragment = collaboration?.yXmlFragment ?? null;
    const collaborationSourceChanged = currentCollaborationFragment !== lastCollaborationFragment;

    const sessionChanged =
      (editorSession.type !== "prosemirror") !== (deps.getExperimentalSession?.() === "canonical");
    if (collaboration && !collaborationSourceChanged && !sessionChanged) {
      return;
    }

    // Skip if this is the same document (likely passed back after internal edit)
    // Only reset state if:
    // 1. Not yet initialized (first mount)
    // 2. Document identity changed (truly external change like loading a new file)
    // 3. Collaboration starts/stops or switches sessions
    if (
      isInitialized &&
      currentDocId === lastDocumentId &&
      !collaborationSourceChanged &&
      !sessionChanged
    ) {
      return;
    }

    if (!seedSession(document)) {
      destroyView();
      return;
    }

    input.reset();

    // Update tracking state
    isInitialized = true;
    lastDocumentId = currentDocId;
    lastCollaborationFragment = currentCollaborationFragment;

    // Create new state from document
    const newState =
      editorSession.type === "canonical"
        ? canonicalState(editorSession.session)
        : createHiddenEditorState({
            document,
            styles: deps.getStyles(),
            manager: deps.getExtensionManager(),
            externalPlugins: deps.getExternalPlugins(),
            collaboration,
            collaborationModules,
            reason: "external-document",
          });
    const updateStartedAt = performance.now();
    view.updateState(newState);
    recordHiddenEditorPhase(
      "external-document",
      "update-state",
      performance.now() - updateStartedAt,
    );
    syncHiddenEditorAccessibility(view, deps.getReadOnly());

    deps.onSelectionChange(newState);
  };

  const syncEditable = (): void => {
    if (!view) {
      return;
    }
    // EditorView calls editable() dynamically; ARIA state needs explicit sync.
    syncHiddenEditorAccessibility(view, deps.getReadOnly());
  };

  const publishStoryCommit = (storyView: EditorView, commit: CanonicalCommit): boolean => {
    if (!view || editorSession.type !== "canonical") return false;
    if (deps.getEditingMode?.() === "suggesting") {
      refuse("Suggesting is unavailable in the experimental canonical session.");
      return false;
    }
    const session = editorSession.session;
    const bodyTransaction = view.state.tr;
    if (!bodyTransaction.doc.eq(commit.bodyProjection.doc))
      bodyTransaction.replaceWith(
        0,
        bodyTransaction.doc.content.size,
        commit.bodyProjection.doc.content,
      );
    markPackageChange(bodyTransaction);
    bodyTransaction.setMeta("addToHistory", false);
    const bodyStaged = Result.try(() => view?.state.applyTransaction(bodyTransaction));
    if (
      bodyStaged.isErr() ||
      !bodyStaged.value ||
      !bodyStaged.value.transactions.includes(bodyTransaction) ||
      !bodyStaged.value.state.doc.eq(commit.bodyProjection.doc) ||
      bodyStaged.value.transactions.some(
        (transaction) => transaction !== bodyTransaction && transaction.docChanged,
      )
    ) {
      refuse("A plugin refused or changed the canonical body projection.");
      return false;
    }
    const staged = publishCanonicalProjection({ state: storyView.state, commit, session });
    if (staged.isErr()) {
      refuse(staged.error.message);
      return false;
    }
    storyView.updateState(staged.value.state);
    view.updateState(bodyStaged.value.state);
    deps.onTransaction({
      transactions: staged.value.transactions,
      newState: view.state,
      docChanged: true,
    });
    deps.onSelectionChange(view.state);
    return true;
  };

  const api = createHiddenEditorApi({
    getView: () => view,
    getDocumentContext: () => (editorSession.type === "refused" ? null : deps.getDocumentContext()),
    getCanonicalDocument: () =>
      editorSession.type === "canonical" ? editorSession.session.document : null,
    setCanonicalMode: (mode) => {
      if (editorSession.type !== "canonical") return false;
      modeOverride = mode;
      syncCanonicalMode();
      return true;
    },
    resolveCanonicalRevisions: (revisionIds, resolution) => {
      if (!view || deps.getReadOnly() || editorSession.type !== "canonical") return false;
      const prepared = editorSession.session.prepareResolve(view.state, {
        revisionIds,
        resolution,
      });
      if (prepared.isErr()) {
        refuse(prepared.error.message);
        return false;
      }
      return publishCommit(prepared.value);
    },
    getCanonicalHistory: () =>
      editorSession.type === "canonical"
        ? {
            undo: () => history("undo"),
            redo: () => history("redo"),
            canUndo: () => editorSession.type === "canonical" && editorSession.session.canUndo,
            canRedo: () => editorSession.type === "canonical" && editorSession.session.canRedo,
          }
        : null,
    canonicalOperations: {
      updateCanonicalInputLifecycle: (action) => {
        if (editorSession.type !== "canonical") return false;
        switch (action) {
          case "beginComposition":
            return canonicalInputLifecycle.beginComposition();
          case "endComposition":
            canonicalInputLifecycle.endComposition();
            return true;
          case "breakUndoGroup":
            canonicalInputLifecycle.breakUndoGroup();
            return true;
        }
      },
      getCanonicalStorySelection: (story) => {
        if (editorSession.type !== "canonical") return null;
        const session = editorSession.session;
        const selection = session.selection;
        if (!selection) return null;
        const projection = session.projectStory(story);
        if (projection.isErr()) return null;
        const anchor = projection.value.positionAt(selection.anchor);
        const head = projection.value.positionAt(selection.head);
        return anchor.isOk() && head.isOk() ? { anchor: anchor.value, head: head.value } : null;
      },
      applyCanonicalStoryHistory: ({ view: storyView, story, direction }) => {
        if (story === OP_STORIES.MAIN && storyView === view) return history(direction);
        if (!view || editorSession.type !== "canonical" || deps.getReadOnly()) return false;
        const session = editorSession.session;
        if (direction === "undo" ? !session.canUndo : !session.canRedo) return false;
        const prepared =
          direction === "undo"
            ? session.prepareUndo(storyView.state, story)
            : session.prepareRedo(storyView.state, story);
        if (prepared.isErr()) {
          refuse(prepared.error.message);
          return false;
        }
        return publishStoryCommit(storyView, prepared.value);
      },
      applyCanonicalOperations: (ops) => {
        ensureView();
        if (!view || editorSession.type !== "canonical" || deps.getReadOnly()) return false;
        const prepared = editorSession.session.prepareOperations(view.state, ops);
        if (prepared.isErr()) {
          refuse(prepared.error.message);
          return false;
        }
        return publishCommit(prepared.value);
      },
      getCanonicalStoryProjection: (story) => {
        if (deps.getExperimentalSession?.() === "canonical") ensureView();
        if (editorSession.type !== "canonical" || editorSession.session.isComposing) return null;
        const result = editorSession.session.projectStory(story);
        if (result.isErr()) {
          refuse(result.error.message);
          return null;
        }
        return result.value.doc;
      },
      replaceCanonicalStoryText: ({ view: storyView, story, intent }) => {
        if (!view || editorSession.type !== "canonical" || deps.getReadOnly()) return false;
        const session = editorSession.session;
        const prepared = session.prepareReplace(storyView.state, { ...intent, story });
        if (prepared.isErr()) {
          refuse(prepared.error.message);
          return false;
        }
        return publishStoryCommit(storyView, prepared.value);
      },
    },
    isDestroying: () => isDestroying,
    ensureView,
    isViewRequested,
  });

  return {
    ensureView,
    retryViewCreation,
    isViewRequested,
    destroyView,
    syncExternalDocument,
    syncEditable,
    getView: () => view,
    isInitialized: () => isInitialized,
    api,
  };
};
