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
    expect(errors).toEqual([]);
    expect(editor.current?.hasPendingChanges()).toBe(false);
  } finally {
    await act(async () => root.unmount());
    hook.mockRestore();
    container.remove();
  }
});
