import type { CanonicalGap } from "@stll/folio-core/types/canonicalCapabilities";
/**
 * HiddenProseMirror Component
 *
 * Off-screen ProseMirror instance that owns all keyboard input and state
 * while the paginated layout engine handles visual output. Responsibilities:
 *
 * - Keyboard input handling
 * - Selection state management
 * - Accessibility (semantic document structure for screen readers)
 * - ProseMirror transaction processing
 *
 * Visibility approach: The editor is moved off-viewport with position:fixed
 * and rendered transparent so it can still receive focus and remain part of
 * the accessibility tree. Content width is kept in sync with the document
 * so that ProseMirror's internal measurements stay valid.
 */

import {
  useRef,
  useEffect,
  useLayoutEffect,
  useCallback,
  useImperativeHandle,
  useState,
  forwardRef,
} from "react";
import type { CSSProperties } from "react";

import { panic } from "better-result";
import type { Plugin, EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import {
  collectRemoteSelections,
  createHiddenEditorManager,
  type CollaborationModules,
  type HiddenEditorManager,
  type HiddenEditorTransactionUpdate,
  type HiddenProseMirrorCollaboration,
  type HiddenProseMirrorRemoteSelection,
} from "@stll/folio-core/controller/hiddenEditorManager";
import { loadCollaborationModules } from "@stll/folio-core/controller/collaborationModules";
import type { ExtensionManager } from "@stll/folio-core/prosemirror/extensions/ExtensionManager";
import type { Document, Theme, StyleDefinitions } from "@stll/folio-core/types/document";
import type { HiddenEditorApi } from "@stll/folio-core/controller/hiddenEditorApi";
// Import ProseMirror CSS
import "prosemirror-view/style/prosemirror.css";

import "../styles/prosemirror-layer.css";

export type {
  HiddenProseMirrorCollaboration,
  HiddenProseMirrorRemoteSelection,
} from "@stll/folio-core/controller/hiddenEditorManager";
export { createHiddenEditorState } from "@stll/folio-core/controller/hiddenEditorManager";

type CollaborationLoad =
  | { status: "disabled" | "loading" }
  | { status: "ready"; modules: CollaborationModules }
  | { status: "failed"; error: unknown };

const EMPTY_EXTERNAL_PLUGINS: Plugin[] = [];

// ============================================================================
// TYPES
// ============================================================================

export type HiddenProseMirrorProps = {
  /** The document to edit */
  document: Document | null;
  experimentalSession?: "canonical";
  suggestionModeActive?: boolean;
  suggestionAuthor?: string;
  onSessionRefusal?: (reason: string, gap: CanonicalGap, error?: Error) => void;
  /**
   * Identity of the loaded document (same across internal edits, distinct per
   * load); a change means an external load and resets the editor state.
   */
  documentIdentity: string;
  /** Document styles for style resolution */
  styles?: StyleDefinitions | null;
  /** Theme for styling */
  theme?: Theme | null;
  /** Width in pixels (should match document content width) */
  widthPx?: number;
  /** Whether the editor is read-only */
  readOnly?: boolean;
  /** Callback when document changes via transaction */
  onTransaction?: (update: HiddenEditorTransactionUpdate) => void;
  /** Callback when selection changes */
  onSelectionChange?: (state: EditorState) => void;
  /** External ProseMirror plugins */
  externalPlugins?: Plugin[];
  /** Yjs-backed collaboration document owner. */
  collaboration?: HiddenProseMirrorCollaboration | undefined;
  onRemoteSelectionsChange?: ((selections: HiddenProseMirrorRemoteSelection[]) => void) | undefined;
  /** Extension manager for plugins/schema/commands (optional — falls back to default) */
  extensionManager?: ExtensionManager;
  /** Callback when EditorView is ready */
  onEditorViewReady?: (view: EditorView) => void;
  /** Initial state already built by the parent for pre-view layout. */
  precomputedInitialState?: EditorState | null;
  /** Callback when EditorView is destroyed */
  onEditorViewDestroy?: () => void;
  /** Intercept key events before ProseMirror processes them. Return true to prevent PM handling. */
  onKeyDown?: (view: EditorView, event: KeyboardEvent) => boolean;
  /** Fires when the editor receives a copy event. */
  onCopy?: () => void;
  /** Fires when an editable editor receives a cut event. */
  onCut?: () => void;
  /** Fires when an editable editor receives a paste event. */
  onPaste?: () => void;
  /** Callback when a readonly user action would mutate the document. */
  onReadOnlyEditAttempt?: () => void;
};

export type HiddenProseMirrorRef = HiddenEditorApi & {
  /** Internal composition availability for deferred document notifications. */
  isCanonicalComposing: () => boolean;
  /** Get the off-screen host element. */
  getHostElement: () => HTMLElement | null;
};

// ============================================================================
// STYLES
// ============================================================================

/**
 * Hidden wrapper styles - visually hidden, focus-safe scroll isolation.
 *
 * The focused ProseMirror contenteditable can ask the browser to reveal its
 * caret. Keeping it inside a tiny overflow-hidden fixed wrapper confines that
 * native scroll work to the wrapper instead of the visible document viewport.
 */
const HIDDEN_WRAPPER_STYLES: CSSProperties = {
  position: "fixed",
  left: "-9999px",
  top: "0",
  width: "1px",
  height: "1px",
  overflow: "hidden",
  opacity: 0,
  zIndex: -1,
  pointerEvents: "none",
  contain: "layout paint",
  overflowAnchor: "none",
  // Don't set aria-hidden - the inner editor remains the accessible document.
};

/**
 * Hidden host styles - full document-width PM mount inside the isolated wrapper.
 */
const HIDDEN_HOST_STYLES: CSSProperties = {
  position: "absolute",
  left: "0",
  top: "0",
  userSelect: "none",
  overflowAnchor: "none",
  // Don't use visibility:hidden - the editor must remain focusable.
};

// ============================================================================
// COMPONENT
// ============================================================================

/**
 * HiddenProseMirror - Off-screen ProseMirror editor for keyboard input
 */
export const HiddenProseMirror = forwardRef<HiddenProseMirrorRef, HiddenProseMirrorProps>(
  (props, ref) => {
    const {
      document,
      documentIdentity,
      experimentalSession,
      suggestionModeActive = false,
      suggestionAuthor = "User",
      onSessionRefusal,
      styles,
      theme: _theme,
      widthPx = 612, // Default Letter width at 72dpi
      readOnly = false,
      onTransaction,
      onSelectionChange,
      externalPlugins = EMPTY_EXTERNAL_PLUGINS,
      collaboration,
      extensionManager,
      onEditorViewReady,
      onEditorViewDestroy,
      onKeyDown,
      onCopy,
      onCut,
      onPaste,
      onReadOnlyEditAttempt,
      onRemoteSelectionsChange,
      precomputedInitialState,
    } = props;

    const hasCollaboration = collaboration !== undefined;
    const [collaborationLoad, setCollaborationLoad] = useState<CollaborationLoad>(() => ({
      status: hasCollaboration ? "loading" : "disabled",
    }));
    if (hasCollaboration !== (collaborationLoad.status !== "disabled")) {
      setCollaborationLoad({ status: hasCollaboration ? "loading" : "disabled" });
    }
    const collaborationModules =
      hasCollaboration && collaborationLoad.status === "ready" ? collaborationLoad.modules : null;
    const collaborationModulesError =
      hasCollaboration && collaborationLoad.status === "failed" ? collaborationLoad.error : null;

    // Refs
    const hostRef = useRef<HTMLDivElement>(null);
    // Manager-input refs: the framework-agnostic view manager reads these via
    // accessor functions, so it always sees the latest committed value.
    const readOnlyRef = useRef(readOnly);
    const experimentalSessionRef = useRef(experimentalSession);
    const suggestionModeActiveRef = useRef(suggestionModeActive);
    const suggestionAuthorRef = useRef(suggestionAuthor);
    const onSessionRefusalRef = useRef(onSessionRefusal);
    const documentRef = useRef(document);
    const documentIdentityRef = useRef(documentIdentity);
    const stylesRef = useRef(styles);
    const extensionManagerRef = useRef(extensionManager);
    const externalPluginsRef = useRef(externalPlugins);
    const precomputedInitialStateRef = useRef(precomputedInitialState);
    const collaborationRef = useRef(collaboration);
    const collaborationModulesRef = useRef(collaborationModules);

    // Store callbacks in refs to avoid dependency array issues that cause infinite loops
    // when the parent component passes unstable callback references
    const onTransactionRef = useRef(onTransaction);
    const onSelectionChangeRef = useRef(onSelectionChange);
    const onEditorViewReadyRef = useRef(onEditorViewReady);
    const onEditorViewDestroyRef = useRef(onEditorViewDestroy);
    const onKeyDownRef = useRef(onKeyDown);
    const onCopyRef = useRef(onCopy);
    const onCutRef = useRef(onCut);
    const onPasteRef = useRef(onPaste);
    const onReadOnlyEditAttemptRef = useRef(onReadOnlyEditAttempt);
    const onRemoteSelectionsChangeRef = useRef(onRemoteSelectionsChange);

    // The off-screen EditorView lifecycle (create/destroy, editorProps, and the
    // external-document / editable sync) lives in the framework-agnostic manager;
    // this component keeps the input refs and its effects (which decide *when* to
    // act) and drives the manager through its methods. Created once, like the
    // layout scheduler in PagedEditor.
    const managerRef = useRef<HiddenEditorManager | null>(null);
    useLayoutEffect(() => {
      // Keep refs in sync
      readOnlyRef.current = readOnly;
      experimentalSessionRef.current = experimentalSession;
      suggestionModeActiveRef.current = suggestionModeActive;
      suggestionAuthorRef.current = suggestionAuthor;
      onSessionRefusalRef.current = onSessionRefusal;
      stylesRef.current = styles;
      extensionManagerRef.current = extensionManager;
      externalPluginsRef.current = externalPlugins;
      precomputedInitialStateRef.current = precomputedInitialState;
      onTransactionRef.current = onTransaction;
      onSelectionChangeRef.current = onSelectionChange;
      onEditorViewReadyRef.current = onEditorViewReady;
      onEditorViewDestroyRef.current = onEditorViewDestroy;
      onKeyDownRef.current = onKeyDown;
      onCopyRef.current = onCopy;
      onCutRef.current = onCut;
      onPasteRef.current = onPaste;
      onReadOnlyEditAttemptRef.current = onReadOnlyEditAttempt;
      onRemoteSelectionsChangeRef.current = onRemoteSelectionsChange;
      collaborationRef.current = collaboration;
      collaborationModulesRef.current = collaborationModules;

      // Keep document ref in sync
      documentRef.current = document;
      documentIdentityRef.current = documentIdentity;
      if (managerRef.current === null) {
        managerRef.current = createHiddenEditorManager({
          getHost: () => hostRef.current,
          getDocument: () => documentRef.current,
          getStyles: () => stylesRef.current,
          getExtensionManager: () => extensionManagerRef.current,
          getExternalPlugins: () => externalPluginsRef.current,
          getCollaboration: () => collaborationRef.current,
          getCollaborationModules: () => collaborationModulesRef.current,
          getPrecomputedInitialState: () => precomputedInitialStateRef.current,
          getReadOnly: () => readOnlyRef.current,
          getExperimentalSession: () => experimentalSessionRef.current,
          getEditingMode: () => (suggestionModeActiveRef.current ? "suggesting" : "editing"),
          getSuggestionAuthor: () => suggestionAuthorRef.current,
          onSessionRefusal: (reason, gap, error) =>
            onSessionRefusalRef.current?.(reason, gap, error),
          getDocumentIdentity: () => documentIdentityRef.current,
          getDocumentContext: () => documentRef.current,
          onTransaction: (update) => onTransactionRef.current?.(update),
          onSelectionChange: (state) => onSelectionChangeRef.current?.(state),
          onKeyDown: (view, event) => onKeyDownRef.current?.(view, event) ?? false,
          onCopy: () => onCopyRef.current?.(),
          onCut: () => onCutRef.current?.(),
          onPaste: () => onPasteRef.current?.(),
          onReadOnlyEditAttempt: () => onReadOnlyEditAttemptRef.current?.(),
          onEditorViewReady: (view) => onEditorViewReadyRef.current?.(view),
          onEditorViewDestroy: () => onEditorViewDestroyRef.current?.(),
          onRemoteSelectionsChange: (selections) =>
            onRemoteSelectionsChangeRef.current?.(selections),
        });
      }
    });

    // View creation and external state sync stay passive: ancestor layout
    // effects publish committed callbacks before readiness/selection events.
    // This runs before the awareness subscription below so it can see a view
    // whose deferred collaboration-module load just completed.
    useEffect(() => {
      managerRef.current?.retryViewCreation();
      managerRef.current?.syncExternalDocument();
      managerRef.current?.syncEditable();
    });

    useEffect(() => {
      if (!hasCollaboration) {
        return undefined;
      }

      let cancelled = false;
      void loadCollaborationModules().then(
        (modules) => {
          if (!cancelled) {
            setCollaborationLoad({ status: "ready", modules });
          }
          return undefined;
        },
        (error: unknown) => {
          if (!cancelled) {
            setCollaborationLoad({ status: "failed", error });
          }
          return undefined;
        },
      );

      return () => {
        cancelled = true;
      };
    }, [hasCollaboration]);

    // ========================================================================
    // EditorView Lifecycle
    // ========================================================================

    useEffect(() => {
      const awareness = collaboration?.awareness;
      if (!awareness || !managerRef.current?.getView() || !collaborationModules) {
        onRemoteSelectionsChangeRef.current?.([]);
        return undefined;
      }

      const publishRemoteSelections = () => {
        const view = managerRef.current?.getView();
        if (!view) {
          return;
        }
        onRemoteSelectionsChangeRef.current?.(
          collectRemoteSelections(view.state, awareness, collaborationModules),
        );
      };

      awareness.on("change", publishRemoteSelections);
      publishRemoteSelections();

      return () => {
        awareness.off("change", publishRemoteSelections);
        onRemoteSelectionsChangeRef.current?.([]);
      };
    }, [collaboration, collaborationModules]);

    // Stable wrapper so the unmount effect keeps `destroyView` in its dependency
    // array; the teardown body lives in the manager.
    const destroyView = useCallback(() => {
      managerRef.current?.destroyView();
    }, []);

    useEffect(() => () => destroyView(), [destroyView]);

    // ========================================================================
    // Imperative Handle
    // ========================================================================

    useImperativeHandle(
      ref,
      () => ({
        ...managerRef.current!.api,
        getHostElement: () => hostRef.current,
        isCanonicalComposing: () => managerRef.current?.isCanonicalComposing() ?? false,
      }),
      [],
    );

    if (hasCollaboration && collaborationModulesError) {
      let detail = "unknown error";
      if (collaborationModulesError instanceof Error) {
        detail = collaborationModulesError.message;
      } else if (typeof collaborationModulesError === "string") {
        detail = collaborationModulesError;
      }
      panic(
        `Failed to load collaboration editor modules. Reload the document to retry. Cause: ${detail}`,
      );
    }

    // ========================================================================
    // Render
    // ========================================================================

    return (
      <div className="paged-editor__hidden-pm-wrapper" style={HIDDEN_WRAPPER_STYLES}>
        <div
          ref={hostRef}
          className="paged-editor__hidden-pm"
          style={{
            ...HIDDEN_HOST_STYLES,
            width: widthPx > 0 ? `${widthPx}px` : undefined,
          }}
          // DO NOT set aria-hidden - this editor provides semantic structure
        />
      </div>
    );
  },
);
