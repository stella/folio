import { panic } from "better-result";
import { expect, test } from "bun:test";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import type { Document } from "../types/document";
import {
  createHiddenEditorManager,
  type HiddenEditorManagerDeps,
} from "./hiddenEditorManager";

const source = (text = "Dear reader, hello."): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "12345678",
          content: [{ type: "run", content: [{ type: "text", text }] }],
        },
      ],
    },
  },
});

const mount = (session: "default" | "canonical", initial = source()) => {
  const releaseDom = acquireHarnessDom();
  const host = document.createElement("div");
  document.body.append(host);
  const state = {
    source: initial,
    identity: "first",
    readOnly: false,
    suggesting: false,
  };
  const refusals: string[] = [];
  const updates: Parameters<HiddenEditorManagerDeps["onTransaction"]>[0][] = [];
  const transactionHooks: HiddenEditorManagerDeps["onTransaction"][] = [];
  const manager = createHiddenEditorManager({
    getHost: () => host,
    getDocument: () => state.source,
    getDocumentContext: () => state.source,
    getDocumentIdentity: () => state.identity,
    getStyles: () => null,
    getExtensionManager: () => undefined,
    getExternalPlugins: () => [],
    getReadOnly: () => state.readOnly,
    getEditingMode: () => (state.suggesting ? "suggesting" : "editing"),
    getSuggestionAuthor: () => "Reviewer",
    onTransaction: (update) => {
      updates.push(update);
      for (const hook of transactionHooks) hook(update);
    },
    onCanonicalCompositionLayout: () => {},
    getEmbeddedFontFamilyMap: () => undefined,
    onSelectionChange: () => {},
    onKeyDown: () => false,
    onReadOnlyEditAttempt: () => {},
    onEditorViewReady: () => {},
    onEditorViewDestroy: () => {},
    onRemoteSelectionsChange: () => {},
    onSessionRefusal: (message) => refusals.push(message),
  } satisfies HiddenEditorManagerDeps);
  manager.ensureView();
  const view = manager.getView() ?? panic("Expected mounted editor");
  manager.api.setSelection(view.state.doc.content.size - 1);
  return {
    manager,
    view,
    state,
    refusals,
    updates,
    transactionHooks,
    destroy: () => {
      manager.destroyView();
      host.remove();
      releaseDom();
    },
  };
};

test.each(["default", "canonical"] as const)(
  "%s validates explicit anchors without replacing an active request",
  (session) => {
    const mounted = mount(session);
    const api = mounted.manager.api.autocomplete;
    try {
      expect(api.start({ requestId: "valid" })).toBe("valid");
      expect(api.append("valid", " next")).toBe(true);
      const before = api.get();
      for (const at of [
        { story: "main", blockId: "87654321", offset: 0 },
        { story: "main", blockId: "12345678", offset: -1 },
        { story: "main", blockId: "12345678", offset: 0.5 },
        { story: "main", blockId: "12345678", offset: 100 },
        {
          story: { kind: "header", rId: "rId1" },
          blockId: "12345678",
          offset: 0,
        },
      ] as const) {
        expect(api.start({ at, requestId: "invalid" })).toBeNull();
        expect(api.get()).toBe(before);
      }
      expect(mounted.refusals).toEqual([]);
    } finally {
      mounted.destroy();
    }
  },
);

test.each(["default", "canonical"] as const)(
  "%s word acceptance preserves remaining ghost and isolates each undo step",
  (session) => {
    const mounted = mount(session);
    const editor = mounted.manager.api;
    try {
      const initialVersion = editor.getCanonicalCommittedVersion();
      expect(editor.autocomplete.start({ requestId: "request" })).toBe(
        "request",
      );
      editor.autocomplete.append("request", " Next words.");
      editor.autocomplete.finish("request");
      expect(editor.getCanonicalCommittedVersion()).toBe(initialVersion);
      expect(editor.hasCanonicalPendingChanges()).toBe(false);
      expect(editor.canUndo()).toBe(false);
      expect(editor.autocomplete.accept("word")).toBe(true);
      expect(mounted.view.state.doc.textContent).toBe(
        "Dear reader, hello. Next ",
      );
      expect(editor.autocomplete.get()).toMatchObject({
        status: "shown",
        text: "words.",
      });
      expect(editor.autocomplete.accept()).toBe(true);
      expect(mounted.view.state.doc.textContent).toBe(
        "Dear reader, hello. Next words.",
      );
      expect(editor.autocomplete.get()).toEqual({ status: "idle" });
      expect(editor.undo()).toBe(true);
      expect(mounted.view.state.doc.textContent).toBe(
        "Dear reader, hello. Next ",
      );
      expect(editor.undo()).toBe(true);
      expect(mounted.view.state.doc.textContent).toBe("Dear reader, hello.");
      expect(mounted.refusals).toEqual([]);
    } finally {
      mounted.destroy();
    }
  },
);

test("canonical suggesting acceptance remains tracked and read-only acceptance cannot edit", () => {
  const mounted = mount("canonical");
  const editor = mounted.manager.api;
  try {
    mounted.state.suggesting = true;
    editor.autocomplete.start({ requestId: "tracked" });
    editor.autocomplete.append("tracked", " Next words.");
    editor.autocomplete.finish("tracked");
    mounted.state.readOnly = true;
    const version = editor.getCanonicalCommittedVersion();
    expect(editor.autocomplete.accept("word")).toBe(false);
    expect(editor.getCanonicalCommittedVersion()).toBe(version);
    mounted.state.readOnly = false;
    expect(editor.autocomplete.accept("word")).toBe(true);
    expect(editor.autocomplete.get()).toMatchObject({ text: "words." });
    expect(editor.autocomplete.accept()).toBe(true);
    const paragraph = editor
      .getCanonicalDocument()
      ?.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") panic("Expected accepted paragraph");
    expect(paragraph.content.some((item) => item.type === "insertion")).toBe(
      true,
    );
    expect(editor.undo()).toBe(true);
    expect(mounted.view.state.doc.textContent).toBe(
      "Dear reader, hello. Next ",
    );
    expect(mounted.refusals).toEqual([]);
  } finally {
    mounted.destroy();
  }
});

test.each(["default", "canonical"] as const)(
  "%s replaces documents and destroys views without retaining a stale stream",
  (session) => {
    const mounted = mount(session);
    const api = mounted.manager.api.autocomplete;
    try {
      api.start({ requestId: "old" });
      api.append("old", " next");
      mounted.state.source = source("New document");
      mounted.state.identity = "second";
      mounted.manager.syncExternalDocument();
      expect(api.get()).toEqual({ status: "idle" });
      expect(api.append("old", " stale")).toBe(false);
      expect(api.start({ requestId: "current" })).toBe("current");
      mounted.manager.destroyView();
      expect(api.get()).toEqual({ status: "idle" });
      expect(api.append("current", " stale")).toBe(false);
      expect(api.start()).toBeNull();
      expect(api.accept()).toBe(false);
    } finally {
      mounted.destroy();
    }
  },
);

test("canonical anchors retain exact paragraph identity when accepting actual typing", () => {
  const initial = source();
  const paragraph = initial.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Expected paragraph");
  paragraph.paraId = "ABCDEF12";
  const mounted = mount("canonical", initial);
  try {
    const editor = mounted.manager.api;
    expect(
      editor.autocomplete.start({
        at: { story: "main", blockId: "abcdef12", offset: 19 },
        requestId: "alias",
      }),
    ).toBeNull();
    expect(
      editor.autocomplete.start({
        at: { story: "main", blockId: "ABCDEF12", offset: 19 },
        requestId: "alias",
      }),
    ).toBe("alias");
    editor.autocomplete.append("alias", " next");
    expect(editor.autocomplete.get()).toMatchObject({
      anchor: { blockId: "ABCDEF12" },
    });
    expect(editor.autocomplete.accept("word")).toBe(true);
    expect(mounted.view.state.doc.textContent).toBe("Dear reader, hello. next");
    expect(mounted.refusals).toEqual([]);
  } finally {
    mounted.destroy();
  }
});

test.each(["default", "canonical"] as const)(
  "%s typing consumes only its exact marker gap and follows the producer caret",
  (session) => {
    for (const ghostGap of [0, 1, 2])
      for (const typedGap of [0, 1, 2]) {
        const mounted = mount(session, {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "12345678",
                  content: [
                    { type: "run", content: [{ type: "text", text: "a" }] },
                    { type: "bookmarkStart", id: 1, name: "_Ref" },
                    { type: "bookmarkEnd", id: 1 },
                    { type: "run", content: [{ type: "text", text: "bc" }] },
                  ],
                },
              ],
            },
          },
        });
        try {
          const editor = mounted.manager.api;
          const address = {
            story: "main",
            blockId: "12345678",
            offset: 1,
            zeroWidthBefore: ghostGap,
          } as const;
          editor.autocomplete.start({ at: address, requestId: "gap" });
          editor.autocomplete.append("gap", "xyz");
          const position = editor.editorPositionAt({
            ...address,
            zeroWidthBefore: typedGap,
          });
          if (position === null) panic("Expected represented marker gap");
          const typeAt = (at: number, text: string) => {
            editor.setSelection(at);
            return editor.replaceCanonicalStoryText({
              view: mounted.view,
              story: "main",
              intent: { from: at, to: at, text, semantic: "typing" },
            });
          };
          expect(typeAt(position, "x")).toBe(true);
          if (ghostGap !== typedGap) {
            expect(editor.autocomplete.get()).toEqual({ status: "idle" });
          } else {
            const current = editor.autocomplete.get();
            if (current.status === "idle")
              panic("Expected remaining ghost at typed gap");
            expect(current.text).toBe("yz");
            expect(editor.editorPositionAt(current.anchor)).toBe(
              mounted.view.state.selection.head,
            );
            const next = mounted.view.state.selection.head;
            expect(typeAt(next, "y")).toBe(true);
            expect(editor.autocomplete.get()).toMatchObject({
              text: "z",
              anchor: { offset: 3 },
            });
          }
          expect(mounted.refusals).toEqual([]);
        } finally {
          mounted.destroy();
        }
      }
  },
);
