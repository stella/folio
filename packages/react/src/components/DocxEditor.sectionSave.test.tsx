import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { TextSelection } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import { IntlProvider } from "use-intl";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "@stll/folio-core/document-operations";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { fromMarkdown } from "@stll/folio-core/markdown";
import { acceptAllChanges } from "@stll/folio-core/prosemirror/commands/comments";
import { FolioDocxReviewer } from "@stll/folio-core/server";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

const FIRST = "First section.";
const SECOND = "Second section.";

/** Two sections: the first, landscape, ends at the mark of its only paragraph. */
const twoSections = async (): Promise<ArrayBuffer> => {
  const model = fromMarkdown(`${FIRST}\n\n${SECOND}`);
  const [first] = model.package.document.content;
  if (first?.type !== "paragraph") panic("The fixture lost its first paragraph");
  first.sectionProperties = {
    ...structuredClone(model.package.document.finalSectionProperties),
    sectionStart: "nextPage",
    orientation: "landscape",
  };
  return await createDocx(model);
};

/** What the mounted editor reports: its body views and its errors. */
const createReports = () => {
  const views: EditorView[] = [];
  const errors: Error[] = [];
  const onEditorViewReady = (view: EditorView | null) => {
    if (view !== null) views.push(view);
  };
  const onError = (error: Error) => {
    errors.push(error);
  };
  return { views, errors, onEditorViewReady, onError };
};

/** Mount the editor over `bytes`, run `edit` on its body view, and save. */
const editAndSave = async (
  bytes: ArrayBuffer,
  edit: (view: EditorView, editor: DocxEditorRef) => void,
  mode: "editing" | "suggesting" = "editing",
): Promise<{ saved: ArrayBuffer | null; errors: Error[] }> => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const { views, errors, onEditorViewReady, onError } = createReports();
  try {
    await act(async () => {
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor
            ref={editor}
            documentBuffer={bytes}
            mode={mode}
            onEditorViewReady={onEditorViewReady}
            onError={onError}
            showToolbar={false}
          />
        </IntlProvider>,
      );
    });
    for (let attempt = 0; attempt < 50 && !views.at(-1)?.state.doc.textContent; attempt++) {
      await act(async () => {
        editor.current?.ensureEditorView({ focus: false });
        await new Promise((resolve) => setTimeout(resolve, 10));
      });
    }
    const view = views.at(-1) ?? panic("The editor did not report its body view");
    const ref = editor.current ?? panic("The editor did not expose its ref");
    await act(async () => edit(view, ref));
    let saved: ArrayBuffer | null = null;
    await act(async () => {
      saved = (await editor.current?.save()) ?? null;
    });
    return { saved, errors };
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
};

/** Press `key` in the body view, as the keyboard does. */
const press = (view: EditorView, key: "Backspace" | "Delete"): void => {
  const event = new KeyboardEvent("keydown", { key, bubbles: true, cancelable: true });
  expect(view.someProp("handleKeyDown", (handle) => handle(view, event))).toBe(true);
};

/** Put the selection in the body view, from `anchor` to `head`. */
const select = (view: EditorView, anchor: number, head = anchor): void => {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, anchor, head)));
};

/** Where the text of the body's second paragraph starts. */
const secondStart = (view: EditorView): number => view.state.doc.child(0).nodeSize + 1;

/** What the saved package holds: its paragraphs, and whether any is landscape. */
const savedSections = async (bytes: ArrayBuffer) => {
  const body = (await parseDocx(bytes)).package.document;
  return {
    carriers: body.content.filter(
      (block) => block.type === "paragraph" && block.sectionProperties !== undefined,
    ).length,
    finalOrientation: body.finalSectionProperties?.orientation ?? "portrait",
  };
};

const sectionCarriers = async (bytes: ArrayBuffer): Promise<number> =>
  (await parseDocx(bytes)).package.document.content.filter(
    (block) => block.type === "paragraph" && block.sectionProperties !== undefined,
  ).length;

test("the editor saves after a direct deleteBlock of the paragraph that ends a section", async () => {
  const source = await twoSections();
  expect(await sectionCarriers(source)).toBe(1);

  const { saved, errors } = await editAndSave(source, (_view, editor) => {
    const snapshot = editor.createAIEditSnapshot() ?? panic("The editor made no snapshot");
    const first = snapshot.blocks.at(0) ?? panic("The snapshot has no blocks");
    expect(first.text).toBe(FIRST);
    const result = editor.applyDocumentOperations({
      snapshot,
      batch: {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: [{ id: "delete", type: "deleteBlock", blockId: first.id }],
      },
    });
    expect(result.applied).toHaveLength(1);
  });

  expect(errors).toEqual([]);
  if (saved === null) panic("The save returned nothing");
  const reopened = await parseDocx(saved);
  expect(await sectionCarriers(saved)).toBe(0);
  expect(reopened.package.document.content).toHaveLength(1);
});

test("the editor saves after accepting a tracked deletion of that paragraph", async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await twoSections(), { author: "Reviewer" });
  const [first] = reviewer.getContent();
  if (!first) panic("The fixture has no blocks");
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [{ id: "delete", type: "deleteBlock", blockId: first.id }],
  });
  const pending = await reviewer.toBuffer();
  expect(await sectionCarriers(pending)).toBe(1);

  const { saved, errors } = await editAndSave(pending, (view) => {
    acceptAllChanges()(view.state, view.dispatch);
  });

  expect(errors).toEqual([]);
  if (saved === null) panic("The save returned nothing");
  expect(await sectionCarriers(saved)).toBe(0);
});

test("the editor saves after Backspace joins across the paragraph that ends a section", async () => {
  const { saved, errors } = await editAndSave(await twoSections(), (view) => {
    select(view, secondStart(view));
    press(view, "Backspace");
    expect(view.state.doc.childCount).toBe(1);
    expect(view.state.doc.textContent).toBe(FIRST + SECOND);
  });

  expect(errors).toEqual([]);
  if (saved === null) panic("The save returned nothing");
  // The first section is gone; its content took the following section's
  // properties, as a deleteBlock of the paragraph leaves them.
  expect(await savedSections(saved)).toEqual({ carriers: 0, finalOrientation: "portrait" });
});

test("the editor saves after deleting a selection that spans the section break", async () => {
  const { saved, errors } = await editAndSave(await twoSections(), (view) => {
    select(view, 1 + 5, secondStart(view) + 6);
    press(view, "Delete");
    expect(view.state.doc.childCount).toBe(1);
  });

  expect(errors).toEqual([]);
  if (saved === null) panic("The save returned nothing");
  expect(await savedSections(saved)).toEqual({ carriers: 0, finalOrientation: "portrait" });
});

test("in suggesting mode, Backspace across the break is tracked and saves once accepted", async () => {
  const { saved, errors } = await editAndSave(
    await twoSections(),
    (view) => {
      select(view, secondStart(view));
      press(view, "Backspace");
      // Tracked: both paragraphs are still there, the first one's mark deleted.
      expect(view.state.doc.childCount).toBe(2);
      expect(view.state.doc.child(0).attrs["pPrMark"]).toMatchObject({ kind: "del" });
      acceptAllChanges()(view.state, view.dispatch);
      expect(view.state.doc.childCount).toBe(1);
    },
    "suggesting",
  );

  expect(errors).toEqual([]);
  if (saved === null) panic("The save returned nothing");
  expect(await savedSections(saved)).toEqual({ carriers: 0, finalOrientation: "portrait" });
});
