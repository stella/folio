import { assertExactModel } from "../../../../test/exactModel";
import { CANONICAL_GAP } from "../types/canonicalCapabilities";
import { describe, expect, spyOn, test } from "bun:test";
import { Fragment, Slice } from "prosemirror-model";
import { pasteWithoutFormatting } from "../prosemirror/commands/pastePlainText";
import { GlobalRegistrator } from "@happy-dom/global-registrator";
import fc from "fast-check";
import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import { TextSelection, type Command } from "prosemirror-state";
import { panic } from "better-result";
import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import { resolveCanonicalReviewRange } from "./canonicalReview";
import { createEmptyDocument } from "../utils/createDocument";
import { EditorState, Plugin } from "prosemirror-state";
import { EditorView } from "prosemirror-view";
import { OP_STORIES } from "@stll/docx-core/ops";
import { createHeaderFooterEditorManager } from "./headerFooterEditorManager";
import { createNoteEditorManager } from "./noteEditorManager";

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

test("canonical public batches publish the saved document and share human undo", () => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  const source = createEmptyDocument({ initialText: "Start" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Expected paragraph");
  paragraph.paraId = "12345678";
  const { deps, spies } = makeDeps({
    getHost: () => host,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
  });
  const manager = createHiddenEditorManager(deps);
  try {
    manager.ensureView();
    const view = manager.getView() ?? panic("Missing canonical view");
    const before = manager.api.getCanonicalDocument();
    const result = manager.api.applyCanonicalDocumentOperations({
      snapshot: createFolioAIEditSnapshot(view.state.doc),
      batch: {
        version: 1,
        mode: "direct",
        operations: [
          {
            id: "public",
            type: "replaceInBlock",
            blockId: "12345678",
            find: "Start",
            replace: "End",
          },
        ],
      },
    });
    expect(result?.applied).toEqual([{ id: "public" }]);
    expect(view.state.doc.textContent).toBe("End");
    expect(spies["onTransaction"].calls).toBe(1);
    expect(manager.api.getDocument()).toEqual(manager.api.getCanonicalDocument());
    expect(
      manager.api.undoCanonicalDocumentOperations(result?.undoHandle ?? panic("Missing handle"))
        ?.status,
    ).toBe("undone");
    expect(manager.api.getDocument()).toEqual(before);
    expect(manager.api.canUndo()).toBe(false);
  } finally {
    manager.destroyView();
    host.remove();
    GlobalRegistrator.unregister();
  }
});

test("canonical keyboard formatting and breaks publish journalled intents", () => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  const source = createEmptyDocument({ initialText: "ab" });
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
    if (!view) panic("Expected canonical editor view");

    const select = (from: number, to = from) => {
      manager.api.setSelection(from, to);
      expect(view.state.selection.from).toBe(from);
      expect(view.state.selection.to).toBe(to);
    };
    const expectProjection = () => {
      const projection = manager.api.getCanonicalStoryProjection(OP_STORIES.MAIN);
      if (!projection) panic("Expected canonical body projection");
      expect(view.state.doc.eq(projection)).toBe(true);
    };
    const dispatchKey = (key: string, modifiers: { ctrlKey?: boolean; metaKey?: boolean }) => {
      const event = new KeyboardEvent("keydown", {
        key,
        bubbles: true,
        cancelable: true,
        ...modifiers,
      });
      view.dom.dispatchEvent(event);
      expect(event.defaultPrevented).toBe(true);
    };
    const expectUndoRestores = (before: {
      document: ReturnType<typeof manager.api.getCanonicalDocument>;
      selection: ReturnType<typeof view.state.selection.toJSON>;
    }) => {
      expect(manager.api.undo()).toBe(true);
      expect(manager.api.getCanonicalDocument()).toEqual(before.document);
      expect(view.state.selection.toJSON()).toEqual(before.selection);
      expectProjection();
    };
    const modifiers = [
      { name: "Ctrl", ctrlKey: true },
      { name: "Meta", metaKey: true },
    ] as const;

    for (const modifier of modifiers) {
      for (const { key, property } of [
        { key: "b", property: "bold" },
        { key: "i", property: "italic" },
        { key: "u", property: "underline" },
      ] as const) {
        select(1, 3);
        const before = {
          document: manager.api.getCanonicalDocument(),
          selection: view.state.selection.toJSON(),
        };
        dispatchKey(key, modifier);

        const formatted = manager.api.getCanonicalDocument();
        const first = formatted?.package.document.content.at(0);
        expect(
          first?.type === "paragraph" &&
            first.content.some(
              (run) =>
                run.type === "run" &&
                (property === "underline"
                  ? run.formatting?.underline?.style === "single"
                  : run.formatting?.[property] === true),
            ),
        ).toBe(true);
        expect(reasons).toEqual([]);
        expectProjection();
        expectUndoRestores(before);
      }

      for (const { key, shiftKey, breakType } of [
        { key: "Enter", shiftKey: true, breakType: "textWrapping" },
        { key: "Enter", shiftKey: false, breakType: "page" },
      ] as const) {
        select(2);
        const before = {
          document: manager.api.getCanonicalDocument(),
          selection: view.state.selection.toJSON(),
        };
        const event = new KeyboardEvent("keydown", {
          key,
          bubbles: true,
          cancelable: true,
          ...(shiftKey ? {} : modifier),
          shiftKey,
        });
        view.dom.dispatchEvent(event);
        expect(event.defaultPrevented).toBe(true);

        const inserted = manager.api.getCanonicalDocument()?.package.document.content.at(0);
        expect(
          inserted?.type === "paragraph" &&
            inserted.content.some(
              (run) =>
                run.type === "run" &&
                run.content.some((leaf) => leaf.type === "break" && leaf.breakType === breakType),
            ),
        ).toBe(true);
        expect(reasons).toEqual([]);
        expectProjection();
        expectUndoRestores(before);
      }
    }
  } finally {
    manager.destroyView();
    host.remove();
    GlobalRegistrator.unregister();
  }
});

test.each(["body", "story"] as const)(
  "%s composition defers every story synchronizer while public snapshots remain blocked",
  async (owner) => {
    GlobalRegistrator.register();
    const host = document.createElement("div");
    document.body.append(host);
    const source = createEmptyDocument({ initialText: "Start" });
    const paragraph = source.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") panic("Expected paragraph fixture");
    paragraph.paraId = "12345678";
    const { deps } = makeDeps({
      getHost: () => host,
      getDocument: () => source,
      getDocumentContext: () => source,
      getExperimentalSession: () => "canonical",
    });
    const manager = createHiddenEditorManager(deps);
    const storyDeps = {
      getHost: () => host,
      // Synchronization must defer before either snapshot source is read.
      getDocument: () => manager.api.getDocument(),
      getCanonicalApi: () => manager.api,
      getExperimentalSession: () => "canonical" as const,
      getStyles: () => null,
      getTheme: () => null,
    };
    const synchronizers = [
      createHeaderFooterEditorManager(storyDeps),
      createNoteEditorManager(storyDeps),
    ];
    try {
      manager.ensureView();
      const view = manager.getView();
      if (!view) panic("Expected canonical view");
      const initial = manager.api.getCanonicalDocument();
      for (const synchronizer of synchronizers) synchronizer.sync();
      if (owner === "body") {
        view.dom.dispatchEvent(new CompositionEvent("compositionstart", { bubbles: true }));
      } else {
        expect(manager.api.updateCanonicalInputLifecycle("beginComposition")).toBe(true);
      }
      for (const read of [manager.api.getDocument, manager.api.getCanonicalDocument])
        expect(read).toThrow("Composition must finish before taking a snapshot.");
      expect(manager.api.getCanonicalStoryProjection(OP_STORIES.MAIN)).toBeNull();
      for (const synchronizer of synchronizers) expect(() => synchronizer.sync()).not.toThrow();
      // Synchronization cannot clear the shared pending boundary.
      expect(manager.api.getCanonicalDocument).toThrow();
      if (owner === "body") {
        view.dom.dispatchEvent(new CompositionEvent("compositionend", { bubbles: true }));
        await new Promise<void>((resolve) => setTimeout(resolve, 40));
      } else {
        manager.api.updateCanonicalInputLifecycle("endComposition");
      }
      expect(manager.api.getCanonicalDocument()).toEqual(initial);
      expect(manager.api.getCanonicalStoryProjection(OP_STORIES.MAIN)).not.toBeNull();
      for (const synchronizer of synchronizers) synchronizer.sync();
    } finally {
      for (const synchronizer of synchronizers) synchronizer.destroy();
      manager.destroyView();
      host.remove();
      GlobalRegistrator.unregister();
    }
  },
);

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

test("canonical suggesting uses the current author and preserves explicit mode across input", () => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  const source = createEmptyDocument({ initialText: "Start" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Expected paragraph fixture");
  paragraph.paraId = "12345678";
  let author = "First author";
  let mode = "suggesting";
  const reasons: string[] = [];
  const { deps } = makeDeps({
    getHost: () => host,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    getEditingMode: () => (mode === "suggesting" ? "suggesting" : "editing"),
    getSuggestionAuthor: () => author,
    onSessionRefusal: (reason) => reasons.push(reason),
  });
  const manager = createHiddenEditorManager(deps);
  try {
    manager.ensureView();
    const view = manager.getView();
    if (!view) panic("Expected canonical editor view");
    const type = (text: string) => {
      view.dom.dispatchEvent(
        new InputEvent("beforeinput", {
          inputType: "insertText",
          data: text,
          cancelable: true,
          bubbles: true,
        }),
      );
    };
    manager.api.setSelection(6);
    type("!");
    author = "Second author";
    type("?");
    const inserted = view.state.doc.nodeAt(6);
    const second = view.state.doc.nodeAt(7);
    expect(inserted?.marks.some((mark) => mark.attrs["author"] === "First author")).toBe(true);
    expect(second?.marks.some((mark) => mark.attrs["author"] === "Second author")).toBe(true);
    expect(manager.api.setCanonicalMode({ type: "suggesting", author: "Explicit author" })).toBe(
      true,
    );
    mode = "editing";
    type("+");
    expect(
      view.state.doc.nodeAt(8)?.marks.some((mark) => mark.attrs["author"] === "Explicit author"),
    ).toBe(true);
    expect(reasons).toEqual([]);
    expect(manager.api.undo()).toBe(true);
    expect(view.state.doc.textContent).toBe("Start!?");
    const suggested = manager.api.getCanonicalDocument();
    const revisionIds = new Set<number>();
    view.state.doc.descendants((node) => {
      for (const mark of node.marks) {
        const revisionId: unknown = mark.attrs["revisionId"];
        if (typeof revisionId === "number") revisionIds.add(revisionId);
      }
    });
    expect(revisionIds.size).toBe(2);
    expect(
      resolveCanonicalReviewRange({ editor: manager.api, from: 6, to: 8, resolution: "reject" }),
    ).toBe(true);
    expect(view.state.doc.textContent).toBe("Start");
    expect(manager.api.undo()).toBe(true);
    assertExactModel(manager.api.getCanonicalDocument(), suggested);
  } finally {
    manager.destroyView();
    host.remove();
    GlobalRegistrator.unregister();
  }
});

test.each([
  { kind: "header", rId: "rIdHeader1" },
  { kind: "footer", rId: "rIdFooter1" },
  { kind: "footnote", id: 12 },
] as const)("a mode change refuses direct secondary commits and history in %s", (story) => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  const storyHost = document.createElement("div");
  document.body.append(host, storyHost);
  const source = createEmptyDocument({ initialText: "Body" });
  const bodyParagraph = source.package.document.content.at(0);
  if (bodyParagraph?.type !== "paragraph") panic("Expected body paragraph fixture");
  bodyParagraph.paraId = "12345678";
  const content = [
    {
      ...bodyParagraph,
      paraId: "34567890",
      content: [{ type: "run", content: [{ type: "text", text: "Story" }] }],
    },
  ] satisfies typeof source.package.document.content;
  if (story.kind === "header")
    source.package.headers = new Map([
      [story.rId, { type: "header", hdrFtrType: "default", content }],
    ]);
  else if (story.kind === "footer")
    source.package.footers = new Map([
      [story.rId, { type: "footer", hdrFtrType: "default", content }],
    ]);
  else source.package.footnotes = [{ type: "footnote", id: story.id, content }];
  let mode: "editing" | "suggesting" = "editing";
  const reasons: string[] = [];
  const { deps } = makeDeps({
    getHost: () => host,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    getEditingMode: () => mode,
    onSessionRefusal: (reason) => reasons.push(reason),
  });
  const manager = createHiddenEditorManager(deps);
  let storyView: EditorView | undefined;
  try {
    manager.ensureView();
    const projection = manager.api.getCanonicalStoryProjection(story);
    if (!projection) panic("Expected secondary projection");
    storyView = new EditorView(storyHost, { state: EditorState.create({ doc: projection }) });
    const initialEnd = storyView.state.doc.content.size - 1;
    expect(
      manager.api.replaceCanonicalStoryText({
        view: storyView,
        story,
        intent: { from: initialEnd, to: initialEnd, text: "!" },
      }),
    ).toBe(true);
    const accepted = manager.api.getCanonicalDocument();
    const committedState = storyView.state;
    expect(committedState.doc.textContent).toBe("Story!");
    expect(manager.api.canUndo()).toBe(true);
    mode = "suggesting";
    const committedEnd = storyView.state.doc.content.size - 1;
    expect(
      manager.api.replaceCanonicalStoryText({
        view: storyView,
        story,
        intent: { from: committedEnd, to: committedEnd, text: "untracked" },
      }),
    ).toBe(false);
    expect(
      manager.api.applyCanonicalStoryHistory({ view: storyView, story, direction: "undo" }),
    ).toBe(false);
    expect(manager.api.getCanonicalDocument()).toEqual(accepted);
    expect(storyView.state).toBe(committedState);
    expect(manager.api.canUndo()).toBe(true);
    expect(manager.api.canRedo()).toBe(false);
    expect(reasons).toEqual([
      "Suggesting is unavailable in the experimental canonical session.",
      "Suggesting is unavailable in the experimental canonical session.",
    ]);
  } finally {
    storyView?.destroy();
    manager.destroyView();
    host.remove();
    storyHost.remove();
    GlobalRegistrator.unregister();
  }
});

test("canonical native paste and cut share one journal and clipboard failures cannot delete", () => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  const source = createEmptyDocument({ initialText: "abcd" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Expected clipboard fixture");
  paragraph.paraId = "12345678";
  const reasons: string[] = [];
  let copyTransforms = 0;
  const clipboardPlugin = new Plugin({
    props: {
      transformCopied: (_slice, view) => {
        copyTransforms += 1;
        return new Slice(
          Fragment.from(
            view.state.schema.node(
              "paragraph",
              null,
              view.state.schema.text("copiedX", [view.state.schema.marks["bold"].create()]),
            ),
          ),
          1,
          1,
        );
      },
    },
  });
  const { deps } = makeDeps({
    getHost: () => host,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    getExternalPlugins: () => [clipboardPlugin],
    onSessionRefusal: (reason) => reasons.push(reason),
  });
  const manager = createHiddenEditorManager(deps);
  try {
    manager.ensureView();
    const view = manager.getView();
    if (!view) panic("Expected canonical clipboard view");
    manager.api.setSelection(2, 4);
    const initial = manager.api.getCanonicalDocument();
    const initialSelection = view.state.selection.toJSON();
    const pasted = new Slice(
      Fragment.from(
        view.state.schema.node(
          "paragraph",
          { paraId: "4AFE0001" },
          view.state.schema.text("X", [view.state.schema.marks["bold"].create()]),
        ),
      ),
      1,
      1,
    );
    expect(
      view.someProp("handlePaste", (handler) => handler(view, new ClipboardEvent("paste"), pasted)),
    ).toBe(true);
    expect(view.state.doc.textContent).toBe("aXd");
    const afterPaste = manager.api.getCanonicalDocument();
    expect(manager.api.undo()).toBe(true);
    assertExactModel(manager.api.getCanonicalDocument(), initial);
    expect(view.state.selection.toJSON()).toEqual(initialSelection);
    expect(manager.api.canUndo()).toBe(false);
    expect(manager.api.redo()).toBe(true);
    assertExactModel(manager.api.getCanonicalDocument(), afterPaste);

    manager.api.setSelection(2, 3);
    const beforeCut = manager.api.getCanonicalDocument();
    const stateBeforeCut = view.state;
    const clipboard = new Map<string, string>();
    const rejected = new ClipboardEvent("cut", { cancelable: true, bubbles: true });
    Object.defineProperty(rejected, "clipboardData", {
      value: {
        setData: () => {
          throw new TypeError("Clipboard denied");
        },
      },
    });
    view.dom.dispatchEvent(rejected);
    expect(rejected.defaultPrevented).toBe(true);
    assertExactModel(manager.api.getCanonicalDocument(), beforeCut);
    expect(view.state).toBe(stateBeforeCut);
    expect(reasons).toHaveLength(1);

    const accepted = new ClipboardEvent("cut", { cancelable: true, bubbles: true });
    Object.defineProperty(accepted, "clipboardData", {
      value: { setData: (type: string, value: string) => clipboard.set(type, value) },
    });
    view.dom.dispatchEvent(accepted);
    expect(accepted.defaultPrevented).toBe(true);
    expect(clipboard.get("text/plain")).toBe("copiedX");
    expect(copyTransforms).toBe(2);
    expect(clipboard.get("text/html")).toContain("data-pm-slice");
    expect(clipboard.get("text/html")).toContain("<strong");
    expect(view.state.doc.textContent).toBe("ad");
    const afterCut = manager.api.getCanonicalDocument();
    expect(manager.api.undo()).toBe(true);
    assertExactModel(manager.api.getCanonicalDocument(), beforeCut);
    expect(manager.api.redo()).toBe(true);
    assertExactModel(manager.api.getCanonicalDocument(), afterCut);
  } finally {
    manager.destroyView();
    host.remove();
    GlobalRegistrator.unregister();
  }
});

test.each(["selection", "edit", "uncaptured"] as const)(
  "native drag captures its source across a later %s change",
  (change) => {
    GlobalRegistrator.register();
    const host = document.createElement("div");
    document.body.append(host);
    const source = createEmptyDocument({ initialText: "abcd" });
    const paragraph = source.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") panic("Expected drag source fixture.");
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
      if (!view) panic("Expected canonical drag view.");
      manager.api.setSelection(2, 4);
      const slice = view.state.selection.content();
      if (change !== "uncaptured")
        view.dom.dispatchEvent(new DragEvent("dragstart", { bubbles: true, cancelable: true }));
      manager.api.setSelection(5);
      if (change === "edit")
        view.dom.dispatchEvent(
          new InputEvent("beforeinput", {
            inputType: "insertText",
            data: "!",
            bubbles: true,
            cancelable: true,
          }),
        );
      const beforeDrop = manager.api.getCanonicalDocument();
      const selectionBeforeDrop = view.state.selection.toJSON();
      const stateBeforeDrop = view.state;
      const coordinates = spyOn(view, "posAtCoords").mockReturnValue({
        pos: change === "edit" ? 6 : 5,
        inside: 0,
      });
      expect(
        view.someProp("handleDrop", (handler) =>
          handler(view, new DragEvent("drop", { clientX: 10, clientY: 10 }), slice, true),
        ),
      ).toBe(true);
      coordinates.mockRestore();
      if (change === "edit") {
        expect(view.state.doc.textContent).toBe("abcd!");
        assertExactModel(manager.api.getCanonicalDocument(), beforeDrop);
        expect(view.state).toBe(stateBeforeDrop);
        assertExactModel(view.state.selection.toJSON(), selectionBeforeDrop);
        expect(reasons).toHaveLength(1);
        expect(reasons.at(0)).toContain("changed");
        expect(manager.api.undo()).toBe(true);
        expect(view.state.doc.textContent).toBe("abcd");
        expect(manager.api.canUndo()).toBe(false);
      } else if (change === "uncaptured") {
        assertExactModel(manager.api.getCanonicalDocument(), beforeDrop);
        expect(view.state).toBe(stateBeforeDrop);
        assertExactModel(view.state.selection.toJSON(), selectionBeforeDrop);
        expect(reasons).toHaveLength(1);
        expect(reasons.at(0)).toContain("captured source");
        expect(manager.api.canUndo()).toBe(false);
      } else {
        expect(view.state.doc.textContent).toBe("adbc");
        expect(reasons).toEqual([]);
        const moved = manager.api.getCanonicalDocument();
        const movedSelection = view.state.selection.toJSON();
        expect(manager.api.undo()).toBe(true);
        assertExactModel(manager.api.getCanonicalDocument(), beforeDrop);
        assertExactModel(view.state.selection.toJSON(), selectionBeforeDrop);
        expect(manager.api.canUndo()).toBe(false);
        expect(manager.api.redo()).toBe(true);
        assertExactModel(manager.api.getCanonicalDocument(), moved);
        assertExactModel(view.state.selection.toJSON(), movedSelection);
      }
    } finally {
      manager.destroyView();
      host.remove();
      GlobalRegistrator.unregister();
    }
  },
);

test("canonical drop moves atomically and pending clipboard reads respect read-only changes", async () => {
  GlobalRegistrator.register();
  const host = document.createElement("div");
  document.body.append(host);
  const source = createEmptyDocument({ initialText: "abcd" });
  const paragraph = source.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Expected drop fixture");
  paragraph.paraId = "12345678";
  let readOnly = false;
  const { deps, spies } = makeDeps({
    getHost: () => host,
    getDocument: () => source,
    getDocumentContext: () => source,
    getExperimentalSession: () => "canonical",
    getReadOnly: () => readOnly,
  });
  const manager = createHiddenEditorManager(deps);
  const navigatorDescriptor = Object.getOwnPropertyDescriptor(globalThis, "navigator");
  try {
    manager.ensureView();
    const view = manager.getView();
    if (!view) panic("Expected canonical drop view");
    manager.api.setSelection(2, 4);
    const initial = manager.api.getCanonicalDocument();
    const initialSelection = view.state.selection.toJSON();
    const slice = view.state.selection.content();
    view.dom.dispatchEvent(new DragEvent("dragstart", { bubbles: true, cancelable: true }));
    const coordinates = spyOn(view, "posAtCoords").mockReturnValue({ pos: 5, inside: 0 });
    expect(
      view.someProp("handleDrop", (handler) =>
        handler(view, new DragEvent("drop", { clientX: 10, clientY: 10 }), slice, true),
      ),
    ).toBe(true);
    coordinates.mockRestore();
    expect(view.state.doc.textContent).toBe("adbc");
    const moved = manager.api.getCanonicalDocument();
    expect(manager.api.undo()).toBe(true);
    assertExactModel(manager.api.getCanonicalDocument(), initial);
    expect(view.state.selection.toJSON()).toEqual(initialSelection);
    expect(manager.api.canUndo()).toBe(false);
    expect(manager.api.redo()).toBe(true);
    assertExactModel(manager.api.getCanonicalDocument(), moved);

    const pending = Promise.withResolvers<string>();
    Object.defineProperty(globalThis, "navigator", {
      configurable: true,
      value: { clipboard: { readText: () => pending.promise } },
    });
    expect(manager.api.executeCommand(pasteWithoutFormatting)).toBe(true);
    const beforeResolution = manager.api.getCanonicalDocument();
    const stateBeforeResolution = view.state;
    readOnly = true;
    pending.resolve("forbidden");
    await pending.promise;
    await Promise.resolve();
    assertExactModel(manager.api.getCanonicalDocument(), beforeResolution);
    expect(view.state).toBe(stateBeforeResolution);
    expect(spies["onReadOnlyEditAttempt"].calls).toBe(1);
  } finally {
    if (navigatorDescriptor) Object.defineProperty(globalThis, "navigator", navigatorDescriptor);
    else Reflect.deleteProperty(globalThis, "navigator");
    manager.destroyView();
    host.remove();
    GlobalRegistrator.unregister();
  }
});

test.each(["editing", "suggesting"] as const)(
  "a loaded session clears its explicit mode override for host %s",
  (hostMode) => {
    GlobalRegistrator.register();
    const host = document.createElement("div");
    document.body.append(host);
    const source = createEmptyDocument({ initialText: "Start" });
    const paragraph = source.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") panic("Expected paragraph fixture");
    paragraph.paraId = "12345678";
    let identity = "first";
    const { deps } = makeDeps({
      getHost: () => host,
      getDocument: () => source,
      getDocumentContext: () => source,
      getDocumentIdentity: () => identity,
      getExperimentalSession: () => "canonical",
      getEditingMode: () => hostMode,
      getSuggestionAuthor: () => "Current host author",
    });
    const manager = createHiddenEditorManager(deps);
    try {
      manager.ensureView();
      expect(
        manager.api.setCanonicalMode({ type: "suggesting", author: "Previous session author" }),
      ).toBe(true);
      identity = "second";
      manager.syncExternalDocument();
      const view = manager.getView();
      if (view === null) panic("Expected reseeded canonical editor view");
      manager.api.setSelection(6);
      view.dom.dispatchEvent(
        new InputEvent("beforeinput", {
          inputType: "insertText",
          data: "!",
          cancelable: true,
          bubbles: true,
        }),
      );
      const marks = view.state.doc.nodeAt(6)?.marks ?? [];
      expect(marks.some((mark) => mark.attrs["author"] === "Previous session author")).toBe(false);
      expect(marks.some((mark) => mark.attrs["author"] === "Current host author")).toBe(
        hostMode === "suggesting",
      );
      expect(view.state.doc.textContent).toBe("Start!");
    } finally {
      manager.destroyView();
      host.remove();
      GlobalRegistrator.unregister();
    }
  },
);

test(
  "canonical undescribed commands preserve non-document effects and refuse raw mutations",
  () => {
    GlobalRegistrator.register();
    const counts = { selection: 0, probe: 0, mutation: 0 };
    try {
      assertProperty(
        fc.property(
          fc.record({ text: fc.string({ minLength: 1, maxLength: 16 }), position: fc.nat() }),
          ({ text, position }) => {
            const host = document.createElement("div");
            document.body.append(host);
            const source = createEmptyDocument({ initialText: text });
            const paragraph = source.package.document.content.at(0);
            if (paragraph?.type !== "paragraph") panic("Expected generated paragraph");
            paragraph.paraId = "12345678";
            const reasons: string[] = [];
            const { deps } = makeDeps({
              getHost: () => host,
              getDocument: () => source,
              getDocumentContext: () => source,
              getExperimentalSession: () => "canonical",
              onSessionRefusal: (reason, gap) => {
                expect(Object.values(CANONICAL_GAP).includes(gap)).toBe(true);
                reasons.push(reason);
              },
            });
            const manager = createHiddenEditorManager(deps);
            try {
              manager.ensureView();
              const view = manager.getView();
              if (!view) panic("Expected generated canonical view");
              const baseline = manager.api.getCanonicalDocument();
              const target = 1 + (position % (text.length + 1));
              for (const kind of ["selection", "probe", "mutation"] as const) {
                let calls = 0;
                const command: Command = (state, dispatch) => {
                  calls++;
                  switch (kind) {
                    case "selection":
                      dispatch?.(state.tr.setSelection(TextSelection.create(state.doc, target)));
                      return true;
                    case "probe":
                      return false;
                    case "mutation":
                      dispatch?.(state.tr.insertText("unclassified", target));
                      return true;
                  }
                };
                const beforeState = view.state;
                const beforeRefusals = reasons.length;
                expect(manager.api.executeCommand(command)).toBe(kind !== "probe");
                expect(calls).toBe(1);
                counts[kind]++;
                expect(manager.api.getCanonicalDocument()).toEqual(baseline);
                const projection = manager.api.getCanonicalStoryProjection("main");
                if (!projection) panic("Expected generated canonical projection");
                expect(view.state.doc.eq(projection)).toBe(true);
                expect(manager.api.canUndo()).toBe(false);
                expect(manager.api.canRedo()).toBe(false);
                if (kind === "mutation") {
                  expect(view.state).toBe(beforeState);
                  expect(reasons.slice(beforeRefusals)).toEqual([
                    "Unclassified native text is unavailable in this session.",
                  ]);
                } else {
                  expect(reasons).toHaveLength(beforeRefusals);
                  expect(view.state.selection.from).toBe(target);
                }
              }
            } finally {
              manager.destroyView();
              host.remove();
            }
          },
        ),
      );
      for (const count of Object.values(counts)) expect(count).toBeGreaterThan(0);
    } finally {
      GlobalRegistrator.unregister();
    }
  },
  propertyTestTimeout(30_000),
);
