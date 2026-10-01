import { GlobalRegistrator } from "@happy-dom/global-registrator";

GlobalRegistrator.register();

import { afterAll, expect, test } from "bun:test";
import { panic } from "better-result";
import { TextSelection } from "prosemirror-state";

const { createApp, defineComponent, h, shallowRef } = await import("vue");

import { parseDocx } from "@stll/folio-core/docx/parser";
import { createDocx } from "@stll/folio-core/docx/rezip";
import { createEmptyDocument } from "@stll/folio-core/utils/createDocument";
import { reviewDifferences } from "../../../../test/reviewDifferences";

const { useDocxEditor } = await import("./useDocxEditor");

// The save oracle exercises the composable's real hidden manager and serialization,
// rather than rebuilding the expected document from its PM projection.
afterAll(() => GlobalRegistrator.unregister());

test("canonical edits save and reopen the canonical text and paragraph identity", async () => {
  const container = document.createElement("div");
  const hidden = document.createElement("div");
  const pages = document.createElement("div");
  document.body.append(container, hidden, pages);
  const errors: Error[] = [];
  let hostChanges = 0;
  const holder: { editor: ReturnType<typeof import("./useDocxEditor").useDocxEditor> | null } = {
    editor: null,
  };
  const app = createApp(
    defineComponent({
      setup() {
        holder.editor = useDocxEditor({
          hiddenContainer: shallowRef(hidden),
          pagesContainer: shallowRef(pages),
          experimentalSession: "canonical",
          onError: (error) => errors.push(error),
          onChange: (document) => {
            hostChanges++;
            document.package.document.content = [];
          },
        });
        return () => h("div");
      },
    }),
  );
  app.mount(container);
  try {
    const editor = holder.editor ?? panic("Expected mounted Vue editor");
    const bytes = await createDocx(createEmptyDocument({ initialText: "Start" }));
    await editor.loadBuffer(bytes);
    const detachedRead = editor.getDocument() ?? panic("Expected canonical document snapshot");
    detachedRead.package.document.content = [];
    expect(editor.getDocument()?.package.document.content).not.toHaveLength(0);
    const unsupportedWrite = editor.getDocument() ?? panic("Expected canonical document snapshot");
    unsupportedWrite.package.document.content = [];
    editor.setDocument(unsupportedWrite);
    expect(editor.getDocument()?.package.document.content).not.toHaveLength(0);
    expect(errors.at(0)?.message).toContain("Direct document model changes are unavailable");
    const view = editor.editorView.value ?? panic("Expected body editor view");
    view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, 6)));
    view.dom.dispatchEvent(
      new InputEvent("beforeinput", {
        bubbles: true,
        cancelable: true,
        inputType: "insertText",
        data: " edited",
      }),
    );
    const canonical = editor.getDocument() ?? panic("Expected canonical document");
    expect(view.state.doc.textContent).toBe("Start edited");
    expect(editor.isDirty.value).toBe(true);
    await new Promise((resolve) => window.setTimeout(resolve, 300));
    expect(hostChanges).toBeGreaterThan(0);
    const saved = await editor.save();
    expect(errors).toHaveLength(1);
    if (!saved) panic("Expected saved canonical DOCX");
    const reopened = await parseDocx(await saved.arrayBuffer(), {
      preloadFonts: false,
      detectVariables: false,
    });
    expect(reviewDifferences(canonical, reopened)).toEqual({ messages: [], omitted: 0 });
    expect(editor.isDirty.value).toBe(false);
  } finally {
    app.unmount();
    container.remove();
    hidden.remove();
    pages.remove();
  }
});
