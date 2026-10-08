import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, spyOn, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { TextSelection } from "prosemirror-state";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import type { Document } from "@stll/folio-core/types/document";
import {
  createCanonicalHeaderFooterOperation,
  DOCUMENT_OP_TYPES,
  withCanonicalParagraphIds,
} from "@stll/folio-core/controller/canonicalOperations";
import * as headerFooterHook from "./hooks/useHeaderFooterEditor";
import { reviewDifferences } from "../../../../test/reviewDifferences";
import { CanonicalSaveDiagnosticError } from "@stll/folio-core/docx/canonicalSave";
import { CANONICAL_SAVE_FALLBACK_DIAGNOSTIC } from "../../../../test/canonicalSaveDiagnostics";
import type { SaveDiagnostic } from "@stll/folio-core/docx/saveDiagnostics";
import { describePackageDifferences } from "../../../../scripts/lib/corpus-invariants/model-equality";
import {
  CANONICAL_SAVE_SEEDS,
  canonicalSaveFeatureFlags,
  canonicalSaveFixture,
  canonicalSaveSequence,
  canonicalSaveParagraphXml,
} from "../../../../test/canonicalSaveSequence";
import {
  clearTrackedChanges,
  getChangedParagraphIds,
  hasStructuralChanges,
} from "@stll/folio-core/prosemirror/extensions/features/ParagraphChangeTrackerExtension";

import { DocxEditor } from "./DocxEditor";
import * as editorDialogs from "./DocxEditorDialogs";
import type { FootnotePropertiesMount } from "./DocxEditorDialogs";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const createHostMutationObserver = () => {
  let count = 0;
  return {
    onChange: (document: Document) => {
      count++;
      document.package.document.content = [];
    },
    get count() {
      return count;
    },
  };
};

for (const experimentalSession of [undefined, "canonical"] as const) {
  test(`document history shortcuts ${experimentalSession === "canonical" ? "leave canonical history to the editor" : "remain enabled by default"}`, async () => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const initialDocument = createEmptyDocument({ initialText: "Start" });
    try {
      await act(async () => {
        root.render(
          <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
            <DocxEditor
              document={initialDocument}
              {...(experimentalSession ? { experimentalSession } : {})}
              showToolbar={false}
            />
          </IntlProvider>,
        );
      });
      for (const shortcut of [
        { key: "z", ctrlKey: true },
        { key: "y", ctrlKey: true },
        { key: "Z", metaKey: true, shiftKey: true },
      ]) {
        const event = new KeyboardEvent("keydown", { ...shortcut, cancelable: true });
        await act(async () => {
          document.dispatchEvent(event);
        });
        expect(event.defaultPrevented).toBe(experimentalSession !== "canonical");
      }
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  });
}

test("canonical edits save and reopen the canonical text and paragraph identity", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const errors: Error[] = [];
  const footnoteAction: { apply: FootnotePropertiesMount["onApply"] | null } = { apply: null };
  const renderDialogs = editorDialogs.DocxEditorDialogs;
  const dialogs = spyOn(editorDialogs, "DocxEditorDialogs").mockImplementation((props) => {
    footnoteAction.apply = props.footnoteProperties.onApply;
    return renderDialogs(props);
  });
  const hostChanges = createHostMutationObserver();
  // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- This test renders the editor once.
  const onError = (error: Error) => errors.push(error);
  const bytes = await createDocx(createEmptyDocument({ initialText: "Start" }));
  try {
    await act(async () => {
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor
            ref={editor}
            documentBuffer={bytes}
            experimentalSession="canonical"
            onChange={hostChanges.onChange}
            onError={onError}
            showToolbar={false}
          />
        </IntlProvider>,
      );
    });
    await act(async () => editor.current?.loadDocumentBuffer(bytes));
    await act(async () => editor.current?.ensureEditorView({ focus: false }));
    const view = editor.current?.getEditor()?.getView() ?? panic("Expected body editor view");
    await act(async () => {
      view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)));
      view.dom.dispatchEvent(
        new InputEvent("beforeinput", {
          bubbles: true,
          cancelable: true,
          inputType: "insertText",
          data: " edited",
        }),
      );
    });
    const canonical = editor.current?.getDocument() ?? panic("Expected canonical document");
    expect(view.state.doc.textContent).toBe("Start edited");
    expect(editor.current?.hasPendingChanges()).toBe(true);
    await act(async () => await new Promise((resolve) => window.setTimeout(resolve, 300)));
    expect(hostChanges.count).toBeGreaterThan(0);
    let saved: ArrayBuffer | null | undefined;
    await act(async () => {
      saved = await editor.current?.save();
    });
    expect(errors).toEqual([]);
    if (!saved) panic("Expected saved canonical DOCX");
    const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
    expect(reviewDifferences(canonical, reopened)).toEqual({ messages: [], omitted: 0 });
    expect(editor.current?.hasPendingChanges()).toBe(false);
    await act(async () => {
      if (!footnoteAction.apply) panic("Expected mounted footnote properties action");
      footnoteAction.apply({ numStart: 3 }, { numStart: 4 });
    });
    expect(errors).toEqual([]);
    expect(editor.current?.getDocument()?.package.document.finalSectionProperties).toMatchObject({
      footnotePr: { numStart: 3 },
      endnotePr: { numStart: 4 },
    });
    expect(editor.current?.hasPendingChanges()).toBe(true);
    await act(async () => {
      expect(editor.current?.undo()).toBe(true);
    });
    expect(editor.current?.getDocument()).toEqual(canonical);
  } finally {
    await act(async () => root.unmount());
    dialogs.mockRestore();
    container.remove();
  }
});

test("canonical header edits and new footer and note stories survive adapter save and reopen", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const errors: Error[] = [];
  const bindings: { current: ReturnType<typeof headerFooterHook.useHeaderFooterEditor> | null } = {
    current: null,
  };
  const bindHeaderFooter = headerFooterHook.useHeaderFooterEditor;
  const hook = spyOn(headerFooterHook, "useHeaderFooterEditor").mockImplementation((options) => {
    const value = bindHeaderFooter(options);
    bindings.current = value;
    return value;
  });
  // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- This test renders the editor once.
  const onError = (error: Error) => errors.push(error);
  const seed = createEmptyDocument({ initialText: "Body" });
  const body = seed.package.document.content.at(0);
  if (!body || body.type !== "paragraph") panic("Expected seed paragraph");
  body.content.push({ type: "run", content: [{ type: "footnoteRef", id: 1 }] });
  seed.package.footnotes = [
    {
      type: "footnote",
      id: 1,
      content: [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "Existing note" }] }],
        },
      ],
    },
  ];
  seed.package.headers = new Map([
    [
      "rIdCanonicalHeader",
      {
        type: "header",
        hdrFtrType: "default",
        content: [
          {
            type: "paragraph",
            content: [{ type: "run", content: [{ type: "text", text: "Header" }] }],
          },
        ],
      },
    ],
  ]);
  seed.package.document.finalSectionProperties = {
    ...seed.package.document.finalSectionProperties,
    headerReferences: [{ type: "default", rId: "rIdCanonicalHeader" }],
  };
  const bytes = await createDocx(seed);
  try {
    await act(async () => {
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor
            ref={editor}
            documentBuffer={bytes}
            experimentalSession="canonical"
            onError={onError}
            showToolbar={false}
          />
        </IntlProvider>,
      );
    });
    await act(async () => editor.current?.loadDocumentBuffer(bytes));
    await act(async () => editor.current?.ensureEditorView({ focus: false }));
    await act(async () => bindings.current?.handleHeaderFooterDoubleClick("header"));
    const headerView =
      container.querySelector(".paged-editor__hidden-hf-pm .ProseMirror") ??
      panic("Expected mounted header editor");
    await act(async () => {
      headerView.dispatchEvent(
        new InputEvent("beforeinput", {
          inputType: "insertText",
          data: "Edited ",
          bubbles: true,
          cancelable: true,
        }),
      );
    });
    const api = editor.current?.getEditor() ?? panic("Expected editor API");
    const canonical = api.getCanonicalDocument() ?? panic("Expected canonical document");
    const mainParagraph = canonical.package.document.content.at(0);
    if (!mainParagraph || mainParagraph.type !== "paragraph" || !mainParagraph.paraId)
      panic("Expected identified body paragraph");
    const bodyParagraphId = mainParagraph.paraId;
    const newNote = withCanonicalParagraphIds(
      [
        {
          type: "paragraph",
          content: [{ type: "run", content: [{ type: "text", text: "New note" }] }],
        },
      ],
      canonical,
    );
    const footer = createCanonicalHeaderFooterOperation({
      document: canonical,
      position: "footer",
      referenceType: "first",
    });
    await act(async () => {
      expect(
        api.applyCanonicalOperations([
          {
            ...footer,
            content: withCanonicalParagraphIds(
              [
                {
                  type: "paragraph",
                  content: [{ type: "run", content: [{ type: "text", text: "First footer" }] }],
                },
              ],
              canonical,
            ),
          },
          {
            type: DOCUMENT_OP_TYPES.ADD_NOTE,
            at: { story: "main", blockId: bodyParagraphId, offset: 1 },
            note: { type: "footnote", id: 2, content: newNote },
          },
        ]),
      ).toBe(true);
    });
    expect(editor.current?.hasPendingChanges()).toBe(true);
    const snapshot = editor.current?.getDocument() ?? panic("Expected snapshot");
    const header = snapshot.package.headers?.get("rIdCanonicalHeader")?.content.at(0);
    expect(header?.type === "paragraph" ? header.content : null).toMatchObject([
      { type: "run", content: [{ type: "text", text: "Edited Header" }] },
    ]);
    expect(snapshot.package.document.finalSectionProperties?.titlePg).toBe(true);
    let saved: ArrayBuffer | null | undefined;
    await act(async () => {
      saved = await editor.current?.save();
    });
    if (!saved) panic("Expected saved stories");
    const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
    expect(reviewDifferences(snapshot, reopened)).toEqual({ messages: [], omitted: 0 });
    for (const collection of ["headers", "footers"] as const) {
      expect([...(reopened.package[collection]?.keys() ?? [])].sort()).toEqual(
        [...(snapshot.package[collection]?.keys() ?? [])].sort(),
      );
      for (const [rId, story] of snapshot.package[collection] ?? []) {
        const savedStory =
          reopened.package[collection]?.get(rId) ?? panic("Expected reopened story");
        expect(savedStory.type).toBe(story.type);
        expect(savedStory.hdrFtrType).toBe(story.hdrFtrType);
        expect(
          reviewDifferences(
            { package: { document: { content: story.content } } },
            { package: { document: { content: savedStory.content } } },
          ),
        ).toEqual({ messages: [], omitted: 0 });
      }
    }
    for (const collection of ["footnotes", "endnotes"] as const) {
      expect(reopened.package[collection]?.map(({ id }) => id) ?? []).toEqual(
        snapshot.package[collection]?.map(({ id }) => id) ?? [],
      );
      for (const story of snapshot.package[collection] ?? []) {
        const savedStory =
          reopened.package[collection]?.find(({ id }) => id === story.id) ??
          panic("Expected reopened note");
        expect(
          reviewDifferences(
            { package: { document: { content: story.content } } },
            { package: { document: { content: savedStory.content } } },
          ),
        ).toEqual({ messages: [], omitted: 0 });
      }
    }
    expect(reopened.package.document.finalSectionProperties).toEqual(
      snapshot.package.document.finalSectionProperties,
    );
    expect(reopened.package.footnotes?.map(({ id }) => id)).toEqual([1, 2]);
    expect(reopened.package.footers?.size).toBe(1);
    expect(errors).toHaveLength(1);
    for (const error of errors) {
      expect(error).toBeInstanceOf(CanonicalSaveDiagnosticError);
      if (!(error instanceof CanonicalSaveDiagnosticError)) panic("Expected typed save diagnostic");
      expect(error.gap).toBe("pm-save-projection");
      expect(error.diagnostic).toEqual(CANONICAL_SAVE_FALLBACK_DIAGNOSTIC);
    }
    expect(editor.current?.hasPendingChanges()).toBe(false);
  } finally {
    await act(async () => root.unmount());
    hook.mockRestore();
    container.remove();
  }
});

test("canonical toolbar capture shortcuts publish formatting and break intents with exact undo", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const errors: Error[] = [];
  // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- This test renders the editor once.
  const onError = (error: Error) => errors.push(error);
  const initialDocument = createEmptyDocument({ initialText: "ab" });
  const paragraph = initialDocument.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") panic("Expected keyboard fixture paragraph");
  paragraph.paraId = "12345678";
  let executed = 0;

  try {
    await act(async () => {
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor
            ref={editor}
            document={initialDocument}
            experimentalSession="canonical"
            onError={onError}
            showToolbar
          />
        </IntlProvider>,
      );
    });
    await act(async () => editor.current?.ensureEditorView({ focus: false }));
    const api = editor.current?.getEditor() ?? panic("Expected mounted editor API");
    const view = api.getView() ?? panic("Expected mounted body view");
    const expectProjection = () => {
      const projection = api.getCanonicalStoryProjection("main");
      if (!projection) panic("Expected canonical body projection");
      expect(view.state.doc.eq(projection)).toBe(true);
      expect(errors).toEqual([]);
    };
    const select = async (to: number) => {
      await act(async () => {
        api.setSelection(to === 3 ? 1 : 2, to);
        api.focus();
      });
      expect(view.hasFocus()).toBe(true);
    };
    const press = async (options: KeyboardEventInit) => {
      const event = new KeyboardEvent("keydown", {
        ...options,
        bubbles: true,
        cancelable: true,
      });
      await act(async () => {
        view.dom.dispatchEvent(event);
      });
      expect(event.defaultPrevented).toBe(true);
      executed++;
    };
    const undo = async (before: {
      document: ReturnType<typeof api.getCanonicalDocument>;
      selection: ReturnType<typeof view.state.selection.toJSON>;
    }) => {
      await act(async () => {
        expect(api.undo()).toBe(true);
      });
      expect(api.getCanonicalDocument()).toEqual(before.document);
      expect(view.state.selection.toJSON()).toEqual(before.selection);
      expectProjection();
    };

    for (const modifier of [{ ctrlKey: true }, { metaKey: true }]) {
      for (const { key, property } of [
        { key: "b", property: "bold" },
        { key: "i", property: "italic" },
        { key: "u", property: "underline" },
      ] as const) {
        await select(3);
        const before = {
          document: api.getCanonicalDocument(),
          selection: view.state.selection.toJSON(),
        };
        await press({ key, ...modifier });
        expectProjection();
        const formatted = api.getCanonicalDocument()?.package.document.content.at(0);
        expect(
          formatted?.type === "paragraph" &&
            formatted.content.some(
              (run) =>
                run.type === "run" &&
                (property === "underline"
                  ? run.formatting?.underline?.style === "single"
                  : run.formatting?.[property] === true),
            ),
        ).toBe(true);
        await undo(before);
      }
    }

    for (const shortcut of [
      { key: "Enter", shiftKey: true, breakType: "textWrapping" },
      { key: "Enter", ctrlKey: true, breakType: "page" },
      { key: "Enter", metaKey: true, breakType: "page" },
    ] as const) {
      await select(2);
      const before = {
        document: api.getCanonicalDocument(),
        selection: view.state.selection.toJSON(),
      };
      await press(shortcut);
      expectProjection();
      const content = api.getCanonicalDocument()?.package.document.content;
      expect(content).toHaveLength(1);
      const first = content?.at(0);
      const breaks =
        first?.type === "paragraph"
          ? first.content.flatMap((run) =>
              run.type === "run" ? run.content.filter((leaf) => leaf.type === "break") : [],
            )
          : [];
      expect(breaks).toEqual([{ type: "break", breakType: shortcut.breakType }]);
      await undo(before);
    }
    expect(executed).toBe(9);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});

const createModeControlObserver = () => {
  const modes: string[] = [];
  const errors: Error[] = [];
  return {
    modes,
    errors,
    onModeChange: (mode: string) => {
      modes.push(mode);
    },
    onError: (error: Error) => {
      errors.push(error);
    },
  };
};

test.each(["uncontrolled", "controlled"] as const)(
  "canonical track-changes control follows effective %s mode and saves authored revisions",
  async (control) => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const editor = createRef<DocxEditorRef>();
    const { modes, errors, onModeChange, onError } = createModeControlObserver();
    const bytes = await createDocx(createEmptyDocument({ initialText: "Start" }));
    const render = (mode: "editing" | "suggesting", readOnly = false) => (
      <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
        <DocxEditor
          ref={editor}
          documentBuffer={bytes}
          experimentalSession="canonical"
          mode={control === "controlled" ? mode : undefined}
          author="Control author"
          readOnly={readOnly}
          onModeChange={onModeChange}
          onError={onError}
          showReviewControls
        />
      </IntlProvider>
    );
    try {
      await act(async () => root.render(render("editing")));
      await act(async () => editor.current?.loadDocumentBuffer(bytes));
      await act(async () => editor.current?.ensureEditorView({ focus: false }));
      const view = editor.current?.getEditor()?.getView() ?? panic("Expected canonical view");
      if (control === "uncontrolled") {
        expect(
          editor.current
            ?.getEditor()
            ?.setCanonicalMode({ type: "suggesting", author: "Explicit author" }),
        ).toBe(true);
        // Stable host props preserve an explicit session override; a later effective mode change replaces it.
        await act(async () => root.render(render("editing")));
        await act(async () => {
          view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)));
          view.dom.dispatchEvent(
            new InputEvent("beforeinput", {
              bubbles: true,
              cancelable: true,
              inputType: "insertText",
              data: "override",
            }),
          );
        });
        expect(
          view.state.doc
            .nodeAt(6)
            ?.marks.some((mark) => mark.attrs["author"] === "Explicit author"),
        ).toBe(true);
        await act(async () => {
          editor.current?.getEditor()?.undo();
        });
        expect(view.state.doc.textContent).toBe("Start");
      }
      const getToggle = () =>
        container.querySelector<HTMLButtonElement>('button[aria-label="Toggle Track Changes"]') ??
        panic("Expected actual track-changes control");
      await act(async () => getToggle().click());
      expect(modes).toEqual(["suggesting"]);
      if (control === "controlled") {
        // An unacknowledged host request must not change the authoritative mode.
        expect(getToggle().getAttribute("aria-pressed")).toBe("false");
        await act(async () => {
          view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)));
          view.dom.dispatchEvent(
            new InputEvent("beforeinput", {
              bubbles: true,
              cancelable: true,
              inputType: "insertText",
              data: "!",
            }),
          );
        });
        expect(view.state.doc.nodeAt(6)?.marks.some((mark) => mark.type.name === "insertion")).toBe(
          false,
        );
        await act(async () => root.render(render("suggesting")));
      }
      expect(getToggle().getAttribute("aria-pressed")).toBe("true");
      const end = view.state.doc.content.size - 1;
      await act(async () => {
        view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, end)));
        view.dom.dispatchEvent(
          new InputEvent("beforeinput", {
            bubbles: true,
            cancelable: true,
            inputType: "insertText",
            data: "😀 tracked",
          }),
        );
      });
      expect(view.state.doc.nodeAt(end)?.marks).toContainEqual(
        expect.objectContaining({ attrs: expect.objectContaining({ author: "Control author" }) }),
      );
      const tracked = editor.current?.getDocument() ?? panic("Expected tracked document");
      await act(async () => {
        editor.current?.getEditor()?.undo();
      });
      expect(view.state.doc.textContent).toBe(control === "controlled" ? "Start!" : "Start");
      await act(async () => {
        editor.current?.getEditor()?.redo();
      });
      expect(editor.current?.getDocument()).toEqual(tracked);
      let saved: ArrayBuffer | null | undefined;
      await act(async () => {
        saved = await editor.current?.save();
      });
      if (!saved) panic("Expected tracked save");
      const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
      expect(reviewDifferences(tracked, reopened)).toEqual({ messages: [], omitted: 0 });
      expect(errors).toEqual([]);
      await act(async () => root.render(render("suggesting", true)));
      expect(container.querySelector('button[aria-label="Toggle Track Changes"]')).toBeNull();
      await act(async () => {
        view.dom.dispatchEvent(
          new InputEvent("beforeinput", {
            bubbles: true,
            cancelable: true,
            inputType: "insertText",
            data: "blocked",
          }),
        );
      });
      expect(editor.current?.getDocument()).toEqual(tracked);
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  },
);

test.each(CANONICAL_SAVE_SEEDS)(
  "generated canonical history %s saves independently of PM trackers",
  async (seed) => {
    const container = document.createElement("div");
    document.body.append(container);
    const root = createRoot(container);
    const editor = createRef<DocxEditorRef>();
    const errors: Error[] = [];
    const flags = canonicalSaveFeatureFlags(seed);
    let callbackCount = 0;
    // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- Each generated case mounts once.
    const onError = (error: Error) => errors.push(error);
    const bytes = await createDocx(canonicalSaveFixture(seed));
    try {
      await act(async () => {
        root.render(
          <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
            <DocxEditor
              ref={editor}
              documentBuffer={bytes}
              experimentalSession="canonical"
              onError={onError}
              featureFlags={flags}
              showToolbar={false}
            />
          </IntlProvider>,
        );
      });
      await act(async () => editor.current?.loadDocumentBuffer(bytes));
      await act(async () => editor.current?.ensureEditorView({ focus: false }));
      const adapter = editor.current ?? panic("Expected React adapter");
      const api = adapter.getEditor() ?? panic("Expected canonical controller");
      const view = api.getView() ?? panic("Expected canonical body view");
      const untouchedId = (seed + 3).toString(16).padStart(8, "0").toUpperCase();
      const baselineParagraphs = new Map<string, string>();
      for (let index = 0; index < 4; index++) {
        const paraId = (seed + index).toString(16).padStart(8, "0").toUpperCase();
        baselineParagraphs.set(paraId, await canonicalSaveParagraphXml(bytes, paraId));
      }
      await act(async () => {
        for (const [index, operations] of canonicalSaveSequence(seed).entries()) {
          const before = adapter.getDocument();
          await act(async () => {
            expect(api.applyCanonicalOperations(operations)).toBe(true);
          });
          const canonical = adapter.getDocument() ?? panic("Expected canonical snapshot");
          await act(async () => {
            expect(api.undo()).toBe(true);
          });
          expect(adapter.getDocument()).toEqual(before);
          await act(async () => {
            expect(api.redo()).toBe(true);
          });
          expect(adapter.getDocument()).toEqual(canonical);
          await act(async () => {
            view.dispatch(clearTrackedChanges(view.state));
          });
          expect(getChangedParagraphIds(view.state).size).toBe(0);
          const changedIds = new Set(
            api.captureCanonicalSave()?.changedBlockIds ?? panic("Expected canonical save capture"),
          );
          expect(changedIds.has(untouchedId)).toBe(false);
          expect(hasStructuralChanges(view.state)).toBe(false);
          let saved: ArrayBuffer | null = null;
          await act(async () => {
            saved = await adapter.save({ selective: index % 2 === 0 });
          });
          if (!saved) panic("Expected generated canonical save");
          const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
          expect(describePackageDifferences(canonical, reopened)).toEqual({
            messages: [],
            omitted: 0,
          });
          for (const [paraId, originalXml] of baselineParagraphs) {
            if (!changedIds.has(paraId))
              expect(await canonicalSaveParagraphXml(saved, paraId)).toBe(originalXml);
          }
          // Saving is not a history boundary: the same canonical journal remains reversible.
          await act(async () => {
            expect(api.undo()).toBe(true);
          });
          expect(adapter.getDocument()).toEqual(before);
          await act(async () => {
            expect(api.redo()).toBe(true);
          });
          const diagnostics: SaveDiagnostic[] = [];
          const errorsBeforeSerialization = errors.length;
          const repeated = await api.getDocx({
            mode: index % 2 === 0 ? "full" : "prefer-selective",
            onDiagnostic: (diagnostic) => diagnostics.push(diagnostic),
          });
          expect(errors.length).toBe(errorsBeforeSerialization);
          callbackCount += diagnostics.length;
          for (const diagnostic of diagnostics) {
            expect(diagnostic).toEqual(CANONICAL_SAVE_FALLBACK_DIAGNOSTIC);
          }
          if (!repeated) panic("Expected repeated canonical serialization");
          expect(
            describePackageDifferences(
              canonical,
              await parseDocx(repeated, { preloadFonts: false, detectVariables: false }),
            ),
          ).toEqual({ messages: [], omitted: 0 });
          for (const [paraId, originalXml] of baselineParagraphs) {
            if (!changedIds.has(paraId))
              expect(await canonicalSaveParagraphXml(repeated, paraId)).toBe(originalXml);
          }
        }
      });
      await act(async () => {
        const captured = api.captureCanonicalSave() ?? panic("Expected save snapshot");
        const pendingSave = adapter.save({ selective: false });
        expect(api.updateCanonicalInputLifecycle("beginComposition")).toBe(true);
        expect(api.isCanonicalSaveCurrent(captured.version)).toBe(false);
        const capturedSave = await pendingSave;
        if (!capturedSave) panic("Expected captured save during later composition");
        const capturedBytes = capturedSave;
        expect(
          describePackageDifferences(
            captured.document,
            await parseDocx(capturedBytes, {
              preloadFonts: false,
              detectVariables: false,
            }),
          ),
        ).toEqual({ messages: [], omitted: 0 });
        expect(api.updateCanonicalInputLifecycle("endComposition")).toBe(true);
      });
      if (seed % 2 !== 0) {
        expect(errors.length).toBeGreaterThan(0);
        expect(callbackCount).toBeGreaterThan(0);
      }
      for (const error of errors) {
        expect(error).toBeInstanceOf(CanonicalSaveDiagnosticError);
        if (!(error instanceof CanonicalSaveDiagnosticError))
          panic("Expected typed save diagnostic");
        expect(error.diagnostic).toEqual(CANONICAL_SAVE_FALLBACK_DIAGNOSTIC);
        expect(error.gap).toBe("pm-save-projection");
      }
    } finally {
      await act(async () => root.unmount());
      container.remove();
    }
  },
);
