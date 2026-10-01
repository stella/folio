import { describe, expect, test } from "bun:test";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { panic } from "better-result";
import { createEmptyDocument } from "../utils/createDocument";
import type { EditorState } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import {
  createHiddenEditorManager,
  createHiddenEditorClipboardHandlers,
  type HiddenEditorManagerDeps,
  type HiddenProseMirrorRemoteSelection,
} from "./hiddenEditorManager";

// A real DOM-backed EditorView cannot be constructed in this headless test
// environment (repo convention; see hiddenEditorApi.test.ts). These tests cover
// the lifecycle guards that run before any EditorView exists, plus the
// slice-2a API composition that shares the manager's view owner.

type Spy = { calls: number };

type DepsOverrides = Partial<HiddenEditorManagerDeps>;

const makeDeps = (
  overrides: DepsOverrides = {},
): { deps: HiddenEditorManagerDeps; spies: Record<string, Spy> } => {
  const spies: Record<string, Spy> = {
    onTransaction: { calls: 0 },
    onSelectionChange: { calls: 0 },
    onKeyDown: { calls: 0 },
    onCopy: { calls: 0 },
    onCut: { calls: 0 },
    onPaste: { calls: 0 },
    onReadOnlyEditAttempt: { calls: 0 },
    onEditorViewReady: { calls: 0 },
    onEditorViewDestroy: { calls: 0 },
    onRemoteSelectionsChange: { calls: 0 },
  };

  const deps: HiddenEditorManagerDeps = {
    getHost: () => null,
    getDocument: () => null,
    getStyles: () => null,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: () => false,
    getDocumentIdentity: () => "0",
    getDocumentContext: () => null,
    onTransaction: () => {
      spies["onTransaction"].calls += 1;
    },
    onSelectionChange: (_state: EditorState) => {
      spies["onSelectionChange"].calls += 1;
    },
    onKeyDown: (_view: EditorView, _event: KeyboardEvent) => {
      spies["onKeyDown"].calls += 1;
      return false;
    },
    onCopy: () => {
      spies["onCopy"].calls += 1;
    },
    onCut: () => {
      spies["onCut"].calls += 1;
    },
    onPaste: () => {
      spies["onPaste"].calls += 1;
    },
    onReadOnlyEditAttempt: () => {
      spies["onReadOnlyEditAttempt"].calls += 1;
    },
    onEditorViewReady: (_view: EditorView) => {
      spies["onEditorViewReady"].calls += 1;
    },
    onEditorViewDestroy: () => {
      spies["onEditorViewDestroy"].calls += 1;
    },
    onRemoteSelectionsChange: (_selections: HiddenProseMirrorRemoteSelection[]) => {
      spies["onRemoteSelectionsChange"].calls += 1;
    },
    ...overrides,
  };

  return { deps, spies };
};

describe("createHiddenEditorManager", () => {
  test("starts with no view and uninitialized", () => {
    const { deps } = makeDeps();
    const manager = createHiddenEditorManager(deps);
    expect(manager.getView()).toBeNull();
    expect(manager.isInitialized()).toBe(false);
  });

  test("ensureView marks requested but creates nothing without a host", () => {
    const { deps, spies } = makeDeps({ getHost: () => null });
    const manager = createHiddenEditorManager(deps);
    manager.ensureView();
    expect(manager.isViewRequested()).toBe(true);
    expect(manager.getView()).toBeNull();
    expect(manager.isInitialized()).toBe(false);
    expect(spies["onEditorViewReady"].calls).toBe(0);
  });

  test("retryViewCreation does nothing until creation is requested", () => {
    const { deps, spies } = makeDeps({ getHost: () => null });
    const manager = createHiddenEditorManager(deps);
    manager.retryViewCreation();
    expect(manager.isViewRequested()).toBe(false);
    expect(manager.getView()).toBeNull();
    expect(spies["onEditorViewReady"].calls).toBe(0);
  });

  test("destroyView is a no-op without a view", () => {
    const { deps, spies } = makeDeps();
    const manager = createHiddenEditorManager(deps);
    manager.destroyView();
    expect(spies["onEditorViewDestroy"].calls).toBe(0);
  });

  test("syncExternalDocument is a no-op without a view", () => {
    const { deps, spies } = makeDeps();
    const manager = createHiddenEditorManager(deps);
    expect(() => {
      manager.syncExternalDocument();
    }).not.toThrow();
    expect(spies["onSelectionChange"].calls).toBe(0);
  });

  test("syncEditable is a no-op without a view", () => {
    const { deps } = makeDeps();
    const manager = createHiddenEditorManager(deps);
    expect(() => {
      manager.syncEditable();
    }).not.toThrow();
  });

  test("composes an API that shares the manager's (null) view owner", () => {
    const { deps } = makeDeps();
    const manager = createHiddenEditorManager(deps);
    expect(manager.api.getView()).toBeNull();
    expect(manager.api.getState()).toBeNull();
    expect(manager.api.getDocument()).toBeNull();
  });
});

describe("createHiddenEditorClipboardHandlers", () => {
  test("reports editable copy, cut, and paste without consuming the events", () => {
    const { deps, spies } = makeDeps();
    const handlers = createHiddenEditorClipboardHandlers(deps);
    const event = { preventDefault: () => undefined };

    expect(handlers.copy()).toBe(false);
    expect(handlers.cut(undefined, event)).toBe(false);
    expect(handlers.paste(undefined, event)).toBe(false);
    expect(spies["onCopy"].calls).toBe(1);
    expect(spies["onCut"].calls).toBe(1);
    expect(spies["onPaste"].calls).toBe(1);
    expect(spies["onReadOnlyEditAttempt"].calls).toBe(0);
  });

  test("allows copy but refuses mutating clipboard events in read-only mode", () => {
    const { deps, spies } = makeDeps({ getReadOnly: () => true });
    const handlers = createHiddenEditorClipboardHandlers(deps);
    let prevented = 0;
    const event = {
      preventDefault: () => {
        prevented += 1;
      },
    };

    expect(handlers.copy()).toBe(false);
    expect(handlers.cut(undefined, event)).toBe(true);
    expect(handlers.paste(undefined, event)).toBe(true);
    expect(spies["onCopy"].calls).toBe(1);
    expect(spies["onCut"].calls).toBe(0);
    expect(spies["onPaste"].calls).toBe(0);
    expect(spies["onReadOnlyEditAttempt"].calls).toBe(2);
    expect(prevented).toBe(2);
  });
});

test("canonical manager refuses transaction bypasses and shares one input journal", () => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  const source = createEmptyDocument({ initialText: "Start" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Expected paragraph fixture");
  paragraph.paraId = "12345678";
  const reasons: string[] = [];
  const { deps } = makeDeps({
    getHost: () => host,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    onSessionRefusal: (reason) => reasons.push(reason),
  });
  const manager = createHiddenEditorManager(deps);
  try {
    manager.ensureView();
    const view = manager.getView();
    expect(view).not.toBeNull();
    if (!view) return;
    const initial = manager.api.getCanonicalDocument();
    manager.api.setSelection(6);
    view.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        inputType: "insertText",
        data: "!",
        cancelable: true,
        bubbles: true,
      }),
    );
    const accepted = manager.api.getCanonicalDocument();
    expect(view.state.doc.textContent).toBe("Start!");
    expect(manager.api.canUndo()).toBe(true);
    const state = view.state;
    view.dispatch(view.state.tr.insertText("bypass", 1));
    expect(reasons).toHaveLength(1);
    expect(view.state).toBe(state);
    expect(manager.api.getCanonicalDocument()).toEqual(accepted);
    expect(manager.api.undo()).toBe(true);
    expect(manager.api.getDocument()).toEqual(initial);
    expect(view.state.doc.textContent).toBe("Start");
    expect(manager.api.canRedo()).toBe(true);
    expect(manager.api.redo()).toBe(true);
    expect(manager.api.getDocument()).toEqual(accepted);
  } finally {
    manager.destroyView();
    host.remove();
    GlobalRegistrator.unregister();
  }
});

test("a refused canonical activation reports once per loaded document across retries", () => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  const reasons: string[] = [];
  let identity = "first";
  const { deps } = makeDeps({
    getHost: () => host,
    getDocument: () => createEmptyDocument(),
    getDocumentIdentity: () => identity,
    getExperimentalSession: () => "canonical",
    onSessionRefusal: (reason) => reasons.push(reason),
  });
  const manager = createHiddenEditorManager(deps);
  try {
    for (let attempt = 0; attempt < 10; attempt++) {
      manager.ensureView();
      manager.retryViewCreation();
      manager.syncExternalDocument();
    }
    expect(manager.getView()).toBeNull();
    expect(reasons).toHaveLength(1);
    identity = "second";
    manager.retryViewCreation();
    expect(reasons).toHaveLength(2);
  } finally {
    manager.destroyView();
    GlobalRegistrator.unregister();
  }
});
