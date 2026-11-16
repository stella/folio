import { panic } from "better-result";
import { expect, test } from "bun:test";
import { history } from "prosemirror-history";
import { Plugin } from "prosemirror-state";

import { acquireHarnessDom } from "../../../../test/harnessDom";
import {
  buildPlainTextSlice,
  pastePlainTextSlice,
} from "../prosemirror/commands/pastePlainText";
import { INLINE_CONTENT_CONTROL_NODE_NAME } from "../prosemirror/extensions/nodes/SdtExtension";
import { createSuggestionModePlugin } from "../prosemirror/plugins/suggestionMode";
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

const mount = (
  session: "default" | "canonical" | "prosemirror",
  initial = source(),
) => {
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
    getExternalPlugins: () =>
      session !== "prosemirror"
        ? []
        : [history(), createSuggestionModePlugin(false)],
    getCollaboration: () => undefined,
    getCollaborationModules: () => null,
    getPrecomputedInitialState: () => null,
    getReadOnly: () => state.readOnly,
    getExperimentalSession: () => {
      if (session === "default") return undefined;
      return session === "canonical" ? "canonical" : "legacy";
    },
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

test.each(["default", "canonical", "prosemirror"] as const)(
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

test.each(["default", "canonical", "prosemirror"] as const)(
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
      expect(editor.hasCanonicalPendingChanges()).toBe(
        session !== "prosemirror" ? false : null,
      );
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

test.each(["default", "canonical", "prosemirror"] as const)(
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

test.each(["default", "canonical", "prosemirror"] as const)(
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
            if (session !== "prosemirror")
              return editor.replaceCanonicalStoryText({
                view: mounted.view,
                story: "main",
                intent: { from: at, to: at, text, semantic: "typing" },
              });
            mounted.view.dispatch(mounted.view.state.tr.insertText(text, at));
            return true;
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

const mountLegacyPlugins = (plugins: readonly Plugin[]) => {
  const mounted = mount("prosemirror", source("ab"));
  mounted.view.updateState(
    mounted.view.state.reconfigure({
      plugins: [...mounted.view.state.plugins, ...plugins],
    }),
  );
  mounted.manager.api.setSelection(2);
  return mounted;
};

const showLegacySuggestion = (
  mounted: ReturnType<typeof mountLegacyPlugins>,
) => {
  const api = mounted.manager.api.autocomplete;
  expect(api.start({ requestId: "applied" })).toBe("applied");
  api.append("applied", "xy");
  api.finish("applied");
  return api;
};

test("a filtered legacy acceptance leaves its suggestion and document unchanged", () => {
  const mounted = mountLegacyPlugins([
    new Plugin({ filterTransaction: (transaction) => !transaction.docChanged }),
  ]);
  try {
    const api = showLegacySuggestion(mounted);
    const before = api.get();
    expect(api.accept("word")).toBe(false);
    expect(mounted.view.state.doc.textContent).toBe("ab");
    expect(api.get()).toBe(before);
    expect(mounted.manager.api.canUndo()).toBe(false);
  } finally {
    mounted.destroy();
  }
});

const APPENDED_AUTOCOMPLETE = "autocomplete-test-appended";
for (const appended of ["document", "metadata"] as const)
  test(`legacy ${appended} appends are accounted for before preserving a ghost`, () => {
    const mounted = mountLegacyPlugins([
      new Plugin({
        appendTransaction: (transactions, _before, after) => {
          if (
            !transactions.some((transaction) => transaction.docChanged) ||
            transactions.some((transaction) =>
              transaction.getMeta(APPENDED_AUTOCOMPLETE),
            )
          )
            return null;
          const transaction = after.tr.setMeta(APPENDED_AUTOCOMPLETE, true);
          return appended === "document"
            ? transaction.insertText("!", 1)
            : transaction;
        },
      }),
    ]);
    try {
      const api = showLegacySuggestion(mounted);
      mounted.view.dispatch(mounted.view.state.tr.insertText("x", 2));
      expect(mounted.view.state.doc.textContent).toBe(
        appended === "document" ? "!axb" : "axb",
      );
      if (appended === "document") {
        expect(api.get()).toEqual({ status: "idle" });
        expect(api.accept()).toBe(false);
      } else {
        expect(api.get()).toMatchObject({
          status: "shown",
          text: "y",
          anchor: { offset: 2 },
        });
        expect(api.accept()).toBe(true);
        expect(mounted.view.state.doc.textContent).toBe("axyb");
      }
    } finally {
      mounted.destroy();
    }
  });

test("legacy synchronous subscribers accept from the installed document", () => {
  const mounted = mountLegacyPlugins([]);
  try {
    const api = showLegacySuggestion(mounted);
    const observed: string[] = [];
    let accepted = false;
    const off = mounted.manager.api.autocomplete.subscribe((state) => {
      if (state.status !== "idle" && state.text === "y") {
        observed.push(mounted.view.state.doc.textContent);
        accepted = api.accept();
      }
    });
    mounted.view.dispatch(mounted.view.state.tr.insertText("x", 2));
    off();
    expect(observed).toEqual(["axb"]);
    expect(accepted).toBe(true);
    expect(mounted.view.state.doc.textContent).toBe("axyb");
    expect(mounted.updates.at(-1)?.newState.doc.textContent).toBe(
      mounted.view.state.doc.textContent,
    );
    expect(mounted.updates.at(-1)?.newState).toBe(mounted.view.state);
    expect(api.get()).toEqual({ status: "idle" });
  } finally {
    mounted.destroy();
  }
});

for (const paste of ["native", "plain"] as const)
  test(`${paste} legacy paste dismisses a matching ghost instead of consuming it`, () => {
    const mounted = mountLegacyPlugins([]);
    try {
      const api = showLegacySuggestion(mounted);
      if (paste === "native") expect(mounted.view.pasteText("x")).toBe(true);
      else
        pastePlainTextSlice(
          mounted.view,
          buildPlainTextSlice("x", mounted.view.state.schema),
        );
      expect(mounted.view.state.doc.textContent).toBe("axb");
      expect(api.get()).toEqual({ status: "idle" });
    } finally {
      mounted.destroy();
    }
  });

test("ambiguous control carets cannot trigger or replace an existing usable stream", () => {
  const mounted = mount("prosemirror", {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "12345678",
            content: [
              {
                type: "inlineSdt",
                properties: { sdtType: "richText", id: 7 },
                content: [
                  { type: "run", content: [{ type: "text", text: "in " }] },
                ],
              },
              { type: "run", content: [{ type: "text", text: "after " }] },
            ],
          },
        ],
      },
    },
  });
  try {
    const api = mounted.manager.api.autocomplete;
    expect(api.canTrigger()).toBe(true);
    expect(api.start({ requestId: "usable" })).toBe("usable");
    api.append("usable", "next");
    api.finish("usable");
    const before = api.get();
    const boundaries: number[] = [];
    mounted.view.state.doc.descendants((node, at) => {
      if (node.type.name === INLINE_CONTENT_CONTROL_NODE_NAME)
        boundaries.push(at + node.nodeSize - 1, at + node.nodeSize);
      return true;
    });
    expect(boundaries).toHaveLength(2);
    expect(
      api.start({
        at: { story: "main", blockId: "12345678", offset: 3 },
        requestId: "ambiguous",
      }),
    ).toBeNull();
    for (const at of boundaries) {
      mounted.manager.api.setSelection(at);
      expect(api.canTrigger()).toBe(false);
      expect(api.start({ requestId: "ambiguous" })).toBeNull();
      expect(api.get()).toBe(before);
    }
    expect(api.accept()).toBe(true);
    expect(mounted.view.state.doc.textContent).toBe("in after next");
    expect(mounted.refusals).toEqual([]);
  } finally {
    mounted.destroy();
  }
});

test("a legacy undo restoring a matching prefix dismisses the ghost", () => {
  const mounted = mount("prosemirror", source("axb"));
  try {
    mounted.view.dispatch(mounted.view.state.tr.delete(2, 3));
    mounted.manager.api.setSelection(2);
    const api = showLegacySuggestion(mounted);
    expect(mounted.manager.api.undo()).toBe(true);
    expect(mounted.view.state.doc.textContent).toBe("axb");
    expect(api.get()).toEqual({ status: "idle" });
  } finally {
    mounted.destroy();
  }
});

test("an earlier host callback reentering dispatch dismisses the pending ghost", () => {
  const mounted = mountLegacyPlugins([]);
  try {
    const api = showLegacySuggestion(mounted);
    mounted.transactionHooks.push(({ docChanged, newState }) => {
      if (docChanged)
        mounted.view.dispatch(
          newState.tr.setMeta("host-selection-refresh", true),
        );
    });
    mounted.view.dispatch(mounted.view.state.tr.insertText("x", 2));
    expect(mounted.view.state.doc.textContent).toBe("axb");
    expect(mounted.updates.at(-1)?.newState).toBe(mounted.view.state);
    expect(api.get()).toEqual({ status: "idle" });
    expect(api.accept()).toBe(false);
  } finally {
    mounted.destroy();
  }
});
