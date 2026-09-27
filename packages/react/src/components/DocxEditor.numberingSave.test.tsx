import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { act, createRef } from "react";
import { createRoot } from "react-dom/client";
import type { EditorView } from "prosemirror-view";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { toggleBulletList } from "@stll/folio-core/prosemirror";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(globalThis, "IS_REACT_ACT_ENVIRONMENT");
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

test("save carries editor-created numbering definitions into the document package", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  const views: EditorView[] = [];
  let saveError: Error | null = null;
  // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- This test renders the editor once.
  const onEditorViewReady = (readyView: EditorView | null) => {
    if (readyView) {
      views.push(readyView);
    }
  };
  // oxlint-disable-next-line react-perf/jsx-no-new-function-as-prop -- This test renders the editor once.
  const onError = (error: Error) => {
    saveError = error;
  };
  try {
    const document = await parseDocx(
      await createDocx(createEmptyDocument({ initialText: "List item" })),
      { preloadFonts: false, detectVariables: false },
    );
    await act(async () => {
      root.render(
        <IntlProvider locale="en" timeZone="UTC" messages={getFolioMessages("en")}>
          <DocxEditor
            ref={editor}
            document={document}
            onEditorViewReady={onEditorViewReady}
            onError={onError}
            showToolbar={false}
          />
        </IntlProvider>,
      );
    });
    await act(async () => editor.current?.ensureEditorView({ focus: false }));
    const view = views.at(-1) ?? panic("Expected body editor view");
    await act(async () => {
      toggleBulletList(view.state, view.dispatch);
    });
    const current = editor.current?.getDocument() ?? panic("Expected current document");
    const paragraph = current.package.document.content.at(0);
    if (paragraph?.type !== "paragraph") {
      panic("Expected first paragraph");
    }
    const numId = paragraph.formatting?.numPr?.numId;
    expect(typeof numId).toBe("number");
    expect(current.package.numbering?.nums.some((instance) => instance.numId === numId)).toBe(true);
    let saved: ArrayBuffer | null | undefined;
    await act(async () => {
      saved = await editor.current?.save({ selective: false });
    });
    expect(saveError).toBeNull();
    expect(saved?.byteLength).toBeGreaterThan(0);
    if (!saved) {
      panic("Expected saved DOCX");
    }
    const reopened = await parseDocx(saved, { preloadFonts: false, detectVariables: false });
    const savedParagraph = reopened.package.document.content.at(0);
    expect(
      savedParagraph?.type === "paragraph" ? savedParagraph.formatting?.numPr?.numId : null,
    ).toBe(numId);
  } finally {
    await act(async () => root.unmount());
    container.remove();
  }
});
