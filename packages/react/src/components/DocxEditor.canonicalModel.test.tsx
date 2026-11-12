import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { panic } from "better-result";
import { afterAll, expect, test } from "bun:test";
import { IntlProvider } from "use-intl";

import { getFolioMessages } from "@stll/folio-core/i18n/messages";
import type { Document } from "@stll/folio-core/types/document";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";

import { DocxEditor } from "./DocxEditor";
import type { DocxEditorRef } from "./DocxEditor.props";

const previousActEnvironment = Reflect.get(
  globalThis,
  "IS_REACT_ACT_ENVIRONMENT",
);
Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", true);
afterAll(() => {
  Reflect.set(globalThis, "IS_REACT_ACT_ENVIRONMENT", previousActEnvironment);
  GlobalRegistrator.unregister();
});

/** The documents the host's change listener received. */
const changes: Document[] = [];
const recordChange = (changed: Document) => {
  changes.push(changed);
};

/** Whether `value` is a whole document, not some of its content. */
const isDocument = (value: unknown): boolean =>
  typeof value === "object" &&
  value !== null &&
  "package" in value &&
  typeof value.package === "object" &&
  value.package !== null &&
  "document" in value.package;

test("a settled canonical edit hands the host one copy and keeps the session's document as the model", async () => {
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
  changes.length = 0;
  const original = globalThis.structuredClone;
  let copies = 0;
  try {
    await act(async () => {
      root.render(
        <IntlProvider
          locale="en"
          timeZone="UTC"
          messages={getFolioMessages("en")}
        >
          <DocxEditor
            ref={editor}
            accessibleName="Document editor"
            document={createEmptyDocument({ initialText: "Start" })}
            experimentalSession="canonical"
            showToolbar={false}
            onChange={recordChange}
          />
        </IntlProvider>,
      );
    });
    const view =
      editor.current?.getEditor()?.getView() ?? panic("Expected a body view");
    globalThis.structuredClone = ((
      value: unknown,
      options?: StructuredSerializeOptions,
    ) => {
      if (isDocument(value)) copies += 1;
      return original(value, options);
    }) as typeof structuredClone;
    await act(async () => {
      view.dom.dispatchEvent(
        new InputEvent("beforeinput", {
          bubbles: true,
          cancelable: true,
          inputType: "insertText",
          data: "N",
        }),
      );
      await new Promise((resolve) => setTimeout(resolve, 600));
    });
    globalThis.structuredClone = original;

    expect(view.state.doc.textContent).toContain("N");
    expect(changes).toHaveLength(1);
    // One copy for the host, none for the editor's own document state.
    expect(copies).toBe(1);
    const model =
      editor.current?.getEditor()?.getDocument() ??
      panic("Expected the canonical document");
    expect(changes[0]).toEqual(model);
    // The host owns its copy: emptying it leaves the editor's document alone.
    const notified = changes[0] ?? panic("Expected a notified document");
    notified.package.document.content.length = 0;
    expect(
      editor.current?.getEditor()?.getDocument()?.package.document.content
        .length,
    ).toBeGreaterThan(0);
  } finally {
    globalThis.structuredClone = original;
    await act(async () => root.unmount());
    container.remove();
  }
});
