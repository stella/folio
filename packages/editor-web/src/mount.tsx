/**
 * Mounts folio's React editor behind {@link FolioEditorHost}. Everything the
 * host needs goes through the host object and the returned handle, so the
 * same bundle serves a VS Code webview and a plain page alike.
 */

import { createRef, useState, useSyncExternalStore } from "react";
import type { RefObject } from "react";
import { flushSync } from "react-dom";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import type { HostShortcut } from "@stll/folio-core/managers/editorShortcuts";
import {
  hasStructuralChanges,
  hasUntrackedChanges,
} from "@stll/folio-core/prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { DocxEditor } from "@stll/folio-react";
import type { DocxEditorRef } from "@stll/folio-react";
import { getFolioMessages } from "@stll/folio-react/messages";
import type { EditorState } from "prosemirror-state";

import type {
  FolioEditorDocument,
  FolioEditorHandle,
  FolioEditorHost,
  FolioEditorSaveStrategy,
} from "./host";
import { createHistoryBridgePlugin } from "./historyBridge";
import type { HistoryChange } from "./historyBridge";
import type { FolioEditingMode } from "./modes";

/**
 * The host keeps the undo stack the user sees, and print means nothing inside
 * an embedded page, so the editor leaves both sets of keys to the host.
 */
const HOST_SHORTCUTS = ["history", "print"] as const satisfies readonly HostShortcut[];

/**
 * How long to wait for the editor to mount a freshly parsed document, polled
 * on timers because animation frames stop while a webview is in the background.
 */
const MOUNT_POLL_MS = 16;
const MOUNT_POLL_LIMIT = 250;

export const mountFolioEditor = (root: HTMLElement, host: FolioEditorHost): FolioEditorHandle => {
  const session = createSession(host);
  const reactRoot = createRoot(root);
  // Synchronous, so the editor's ref is attached before the first load starts.
  flushSync(() => {
    reactRoot.render(<FolioEditorApp session={session} />);
  });
  void session.load(host.init.document);

  return {
    serialize: session.serialize,
    undo: () => session.editor.current?.undo() ?? false,
    redo: () => session.editor.current?.redo() ?? false,
    reload: (document) => {
      void session.load(document);
    },
    setMode: (mode) => session.mode.set(mode),
    unmount: () => {
      session.dispose();
      reactRoot.unmount();
    },
  };
};

type FolioEditorAppProps = { session: Session };

const FolioEditorApp = ({ session }: FolioEditorAppProps) => {
  const mode = useSyncExternalStore(session.mode.subscribe, session.mode.get);
  // One plugin instance for the editor's lifetime, as `plugins` requires.
  const [plugins] = useState(() => [createHistoryBridgePlugin(session.onHistoryChange)]);
  const { author } = session.host.init;
  const [locale] = useState(() => canonicalLocale(session.host.init.locale));

  return (
    <IntlProvider
      locale={locale}
      messages={getFolioMessages(locale)}
      timeZone={Intl.DateTimeFormat().resolvedOptions().timeZone}
    >
      <DocxEditor
        ref={session.editor}
        author={author}
        mode={mode}
        onModeChange={session.onModeChange}
        hostShortcuts={HOST_SHORTCUTS}
        plugins={plugins}
        onEditorViewReady={session.onEditorViewReady}
        onError={session.onError}
        preserveDocumentWhileLoading={true}
        // Headers and footers keep undo stacks of their own, which the host's
        // single stack cannot mirror yet.
        showHeaderFooterEditing={false}
      />
    </IntlProvider>
  );
};

/** The host's locale as `Intl` spells it, or English for one `Intl` rejects. */
const canonicalLocale = (locale: string): string => {
  try {
    return Intl.getCanonicalLocales(locale).at(0) ?? "en";
  } catch {
    return "en";
  }
};

type ModeStore = {
  get: () => FolioEditingMode;
  set: (mode: FolioEditingMode) => void;
  subscribe: (listener: () => void) => () => void;
};

const createModeStore = (initial: FolioEditingMode): ModeStore => {
  let mode = initial;
  const listeners = new Set<() => void>();
  return {
    get: () => mode,
    set: (next) => {
      if (next === mode) return;
      mode = next;
      for (const listener of listeners) listener();
    },
    subscribe: (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
};

type Session = ReturnType<typeof createSession>;

type LoadState =
  | { type: "loading"; generation: number; error: string | null }
  | { type: "loaded"; generation: number };

/** The editor's state outside React: its ref, view, loaded document and dirty flag. */
const createSession = (host: FolioEditorHost) => {
  const editor: RefObject<DocxEditorRef | null> = createRef();
  const mode = createModeStore(host.init.mode);
  let bodyState: (() => EditorState) | null = null;
  let loaded: FolioEditorDocument = host.init.document;
  let loadState: LoadState = { type: "loading", generation: 0, error: null };
  let dirty = false;
  let disposed = false;

  const setDirty = (next: boolean) => {
    if (next === dirty) return;
    dirty = next;
    host.onDirtyChange(next);
  };

  const waitForMountedEditor = async (ref: DocxEditorRef): Promise<void> => {
    for (let poll = 0; poll < MOUNT_POLL_LIMIT && ref.getEditorRef() === null; poll += 1) {
      await new Promise((resolve) => setTimeout(resolve, MOUNT_POLL_MS));
    }
  };

  /** The load numbered `generation`, while it is still the one in flight. */
  const inFlight = (generation: number): Extract<LoadState, { type: "loading" }> | null =>
    !disposed && loadState.type === "loading" && loadState.generation === generation
      ? loadState
      : null;

  const load = async (document: FolioEditorDocument): Promise<void> => {
    const ref = editor.current;
    if (ref === null || disposed) return;
    const generation = loadState.generation + 1;
    loadState = { type: "loading", generation, error: null };
    try {
      await ref.loadDocumentBuffer(document.bytes);
    } catch (error) {
      if (inFlight(generation) !== null) {
        host.onLoadFailed(error instanceof Error ? error.message : "The document did not load.");
      }
      return;
    }
    // A document that does not parse resolves the load and reports through `onError`.
    const parsed = inFlight(generation);
    if (parsed === null) return;
    if (parsed.error !== null) {
      host.onLoadFailed(parsed.error);
      return;
    }
    await waitForMountedEditor(ref);
    if (inFlight(generation) === null) return;
    loadState = { type: "loaded", generation };
    loaded = document;
    // Create the body view now rather than on the first click, so the history
    // bridge and the save path see every edit from the first keystroke.
    ref.ensureEditorView({ focus: false });
    setDirty(false);
    host.onLoaded(document.fileVersion);
  };

  const saveStrategy = (): FolioEditorSaveStrategy => {
    const state = bodyState?.();
    if (state === undefined) return { type: "full-repack", reason: "noBodyView" };
    if (hasStructuralChanges(state)) return { type: "full-repack", reason: "structuralChange" };
    if (hasUntrackedChanges(state)) return { type: "full-repack", reason: "untrackedChange" };
    return { type: "selective-first" };
  };

  const serialize = async () => {
    const ref = editor.current;
    if (ref === null || loadState.type !== "loaded") {
      throw new Error("The editor has no document to serialize.");
    }
    const fileVersion = loaded.fileVersion;
    // Read before saving: a save clears the change tracker it is read from.
    const strategy = saveStrategy();
    const buffer = await ref.save();
    if (buffer === null) {
      throw new Error("folio could not serialize the document.");
    }
    setDirty(ref.hasPendingChanges());
    return { bytes: new Uint8Array(buffer), fileVersion, strategy };
  };

  return {
    host,
    editor,
    mode,
    load,
    serialize,
    onHistoryChange: (change: HistoryChange) => {
      if (loadState.type !== "loaded") return;
      if (change === "newStep") host.onEdit();
      setDirty(editor.current?.hasPendingChanges() ?? false);
    },
    onEditorViewReady: (view: { state: EditorState } | null) => {
      bodyState = view === null ? null : () => view.state;
    },
    onModeChange: (next: FolioEditingMode) => {
      mode.set(next);
      host.onModeChange(next);
    },
    onError: (error: Error) => {
      if (loadState.type === "loading") {
        loadState = { type: "loading", generation: loadState.generation, error: error.message };
        return;
      }
      host.onError(error.message);
    },
    dispose: () => {
      disposed = true;
    },
  };
};
