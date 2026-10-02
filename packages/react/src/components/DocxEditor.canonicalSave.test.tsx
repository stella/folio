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
    const savedState = view.state;
    await act(async () => {
      if (!footnoteAction.apply) panic("Expected mounted footnote properties action");
      footnoteAction.apply({ numStart: 3 }, { numStart: 4 });
    });
    expect(errors).toHaveLength(1);
    expect(errors.at(0)?.message).toContain("Footnote and endnote properties are unavailable");
    expect(view.state).toBe(savedState);
    expect(editor.current?.getDocument()).toEqual(canonical);
    expect(editor.current?.hasPendingChanges()).toBe(false);
  } finally {
    await act(async () => root.unmount());
    dialogs.mockRestore();
    container.remove();
  }
});
