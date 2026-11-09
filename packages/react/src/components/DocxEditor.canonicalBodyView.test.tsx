import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { act, createRef } from "react";
import { createRoot } from "react-dom/client";

import { afterAll, expect, spyOn, test } from "bun:test";
import { IntlProvider } from "use-intl";

import * as headerFooterManagers from "@stll/folio-core/controller/headerFooterEditorManager";
import * as noteManagers from "@stll/folio-core/controller/noteEditorManager";
import { getFolioMessages } from "@stll/folio-core/i18n/messages";
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

test("a canonical editor has its body view after mount, whatever the story editors read", async () => {
  // The header, footer and note editors must not be what creates the body
  // view: with their syncs doing nothing, the view still exists after mount.
  const createParts = headerFooterManagers.createHeaderFooterEditorManager;
  const createNotes = noteManagers.createNoteEditorManager;
  const parts = spyOn(
    headerFooterManagers,
    "createHeaderFooterEditorManager",
  ).mockImplementation((deps) => ({ ...createParts(deps), sync: () => {} }));
  const notes = spyOn(
    noteManagers,
    "createNoteEditorManager",
  ).mockImplementation((deps) => ({ ...createNotes(deps), sync: () => {} }));
  const container = document.createElement("div");
  document.body.append(container);
  const root = createRoot(container);
  const editor = createRef<DocxEditorRef>();
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
          />
        </IntlProvider>,
      );
    });
    expect(parts).toHaveBeenCalled();
    expect(notes).toHaveBeenCalled();
    expect(editor.current?.getEditor()?.getView()).not.toBeNull();
  } finally {
    await act(async () => root.unmount());
    container.remove();
    parts.mockRestore();
    notes.mockRestore();
  }
});
